"""
shell_telematico.py
Incapsula un file XML GloBE nella struttura Shell Messaggio Telematico AdE.

Struttura corretta (da sample TDAC9 + XSD fornituraGIRT_v1_0_0 / telematico_v1):

    <tm:Messaggio xsi:schemaLocation="..." xmlns:tm="..." xmlns:girtel="..." ...>
      <tm:Intestazione>
        <tm:CodiceFiscaleFornitore>{CF}</tm:CodiceFiscaleFornitore>
        <tm:SpazioServizioTelematico/>
      </tm:Intestazione>
      <tm:Contenuto CodiceFornitura="GIR26">
        <girtel:Fornitura>
          <globe:GLOBE_OECD version="1.0">   ← XML GloBE embedded direttamente, NO CDATA
            ...
          </globe:GLOBE_OECD>
        </girtel:Fornitura>
      </tm:Contenuto>
    </tm:Messaggio>

Changelog rispetto alla versione precedente:
- Eliminati: IdProtocollo, DataOraCreazione, FornitoreServizio, NomeSoftware,
             VersioneSoftware, PeriodoRiferimento, DescrizioneFornitura,
             NomeFile, DimensioneFile, HashFile, CDATA wrapper.
- Aggiunto:  SpazioServizioTelematico (vuoto, obbligatorio da XSD).
- Namespace: prefisso tm: (non msg:), girtel: per Fornitura.
- XML GloBE: embedded come nodo XML nativo (non serializzato come stringa).
"""

from __future__ import annotations

import re
import uuid
import shutil
from datetime import datetime
from pathlib import Path

from lxml import etree

# ---------------------------------------------------------------------------
# Namespace
# ---------------------------------------------------------------------------
NS_TM     = "www.agenziaentrate.gov.it:specificheTecniche:telent:v1"
NS_GIRTEL = "urn:www.agenziaentrate.gov.it:specificheTecniche:girtel"
NS_GLOBE  = "urn:oecd:ties:globe:v2"
NS_STF    = "urn:oecd:ties:globestf:v5"
NS_ISO    = "urn:oecd:ties:isoglobetypes:v1"
NS_XS     = "http://www.w3.org/2001/XMLSchema"
NS_XSI    = "http://www.w3.org/2001/XMLSchema-instance"

NSMAP_ROOT = {
    "tm":     NS_TM,
    "girtel": NS_GIRTEL,
    "globe":  NS_GLOBE,
    "stf":    NS_STF,
    "iso":    NS_ISO,
    "xs":     NS_XS,
    "xsi":    NS_XSI,
}

SCHEMA_LOCATION = (
    "urn:www.agenziaentrate.gov.it:specificheTecniche:girtel "
    "../telematico/fornituraGIRT_v1.0.0.xsd"
)

CODICE_FORNITURA = "GIR26"

# Regex di validazione CodiceFiscaleFornitore (da telematico_v1.xsd: [0-9]{11}|[A-Z0-9]{16})
# Esteso a [0-9]{10,11} per PIVA italiane che possono avere 10 cifre significative
# (Intercos: 5813780961 = 10 cifre; sample TDAC9: 01212121212 = 11 cifre con leading 0)
_CF_RE = re.compile(r"^([0-9]{10,11}|[A-Z0-9]{16})$")


# ---------------------------------------------------------------------------
# Funzione principale
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Risoluzione placeholder {Guid:D}
# ---------------------------------------------------------------------------

NS_GLOBE_TAG = f"{{{NS_GLOBE}}}"
NS_STF_TAG   = f"{{{NS_STF}}}"

def _new_guid() -> str:
    """Genera un UUID4 nel formato xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx."""
    return str(uuid.uuid4())

def _resolve_guid_placeholders(globe_root) -> None:
    """
    Sostituisce tutti i valori che contengono il placeholder {Guid:D}
    (o varianti come {Guid}) in MessageRefId e DocRefId con UUID4 reali.

    Questo garantisce che la shell MSG generata passi i check AdE 60001 e 60011,
    i quali rifiutano qualsiasi riferimento con placeholder non risolti.
    """
    _GUID_RE = re.compile(r"\{Guid(?::[A-Za-z])?\}")

    # MessageRefId (in MessageSpec, namespace globe:)
    for tag in (f"{NS_GLOBE_TAG}MessageRefId",):
        for el in globe_root.iter(tag):
            if el.text and _GUID_RE.search(el.text):
                el.text = _GUID_RE.sub(_new_guid(), el.text)

    # DocRefId e CorrDocRefId (in DocSpec, namespace stf:)
    for tag in (f"{NS_STF_TAG}DocRefId", f"{NS_STF_TAG}CorrDocRefId"):
        for el in globe_root.iter(tag):
            if el.text and _GUID_RE.search(el.text):
                el.text = _GUID_RE.sub(_new_guid(), el.text)



def encapsulate(
    xml_path: str | Path,
    codice_fiscale: str,
    output_dir: str | Path | None = None,
    archive_source: bool = False,
) -> Path:
    """
    Incapsula *xml_path* (file XML GloBE) nella shell telematica AdE.

    Args:
        xml_path:        Percorso del file XML GloBE sorgente.
        codice_fiscale:  CF numerico (11 cifre) o alfanumerico (16 char) del fornitore.
        output_dir:      Directory di output. Default: stessa directory di xml_path.
        archive_source:  Se True, sposta xml_path in uploads/archivio/ dopo l'elaborazione.

    Returns:
        Path del file _MSG.xml generato.

    Raises:
        ValueError: se codice_fiscale non rispetta il pattern XSD.
        FileNotFoundError: se xml_path non esiste.
        etree.XMLSyntaxError: se il file sorgente non è XML valido.
    """
    xml_path = Path(xml_path)
    if not xml_path.exists():
        raise FileNotFoundError(f"File non trovato: {xml_path}")

    cf = codice_fiscale.strip().upper()
    if not _CF_RE.match(cf):
        raise ValueError(
            f"CodiceFiscaleFornitore non valido: '{cf}'. "
            "Deve essere 11 cifre numeriche o 16 caratteri alfanumerici."
        )

    # Directory di output
    if output_dir is None:
        out_dir = xml_path.parent
    else:
        out_dir = Path(output_dir)
        out_dir.mkdir(parents=True, exist_ok=True)

    stem = xml_path.stem
    out_path = out_dir / f"{stem}_MSG.xml"

    # --- Parse XML GloBE sorgente ---
    parser = etree.XMLParser(remove_blank_text=False)
    globe_tree = etree.parse(str(xml_path), parser)
    globe_root = globe_tree.getroot()  # <globe:GLOBE_OECD>

    # --- Risolve placeholder {Guid:D} in MessageRefId e DocRefId ---
    _resolve_guid_placeholders(globe_root)

    # --- Costruisci struttura shell ---
    root = etree.Element(f"{{{NS_TM}}}Messaggio", nsmap=NSMAP_ROOT)
    root.set(f"{{{NS_XSI}}}schemaLocation", SCHEMA_LOCATION)

    # <tm:Intestazione>
    intestazione = etree.SubElement(root, f"{{{NS_TM}}}Intestazione")

    cf_el = etree.SubElement(intestazione, f"{{{NS_TM}}}CodiceFiscaleFornitore")
    cf_el.text = cf

    # <tm:SpazioServizioTelematico/> — elemento vuoto obbligatorio da XSD
    etree.SubElement(intestazione, f"{{{NS_TM}}}SpazioServizioTelematico")

    # <tm:Contenuto CodiceFornitura="GIR26">
    contenuto = etree.SubElement(root, f"{{{NS_TM}}}Contenuto")
    contenuto.set("CodiceFornitura", CODICE_FORNITURA)

    # <girtel:Fornitura>
    fornitura = etree.SubElement(contenuto, f"{{{NS_GIRTEL}}}Fornitura")

    # XML GloBE embedded direttamente come nodo XML (NO CDATA, NO stringa)
    fornitura.append(globe_root)

    # --- Serializza ---
    out_bytes = etree.tostring(
        root,
        xml_declaration=True,
        encoding="utf-8",
        pretty_print=True,
    )
    out_path.write_bytes(out_bytes)

    print(f"  [shell] Creato: {out_path.name}  ({len(out_bytes):,} byte)")

    # --- Archiviazione sorgente (solo se richiesta e il file è in uploads/) ---
    if archive_source:
        _archive_file(xml_path)

    return out_path


# ---------------------------------------------------------------------------
# Batch su uploads/
# ---------------------------------------------------------------------------

def encapsulate_uploads(
    uploads_dir: str | Path,
    codice_fiscale: str,
) -> list[Path]:
    """
    Processa tutti gli XML in *uploads_dir* (non in archivio/).
    Ogni file viene incapsulato e poi archiviato.

    Returns:
        Lista dei file _MSG.xml generati.
    """
    uploads_dir = Path(uploads_dir)
    xml_files = [
        f for f in uploads_dir.glob("*.xml")
        if f.parent.name != "archivio"
    ]

    if not xml_files:
        print(f"  [shell] Nessun XML in {uploads_dir}")
        return []

    results = []
    for xf in sorted(xml_files):
        print(f"  [shell] Elaborazione: {xf.name}")
        try:
            msg = encapsulate(
                xml_path=xf,
                codice_fiscale=codice_fiscale,
                output_dir=xf.parent,
                archive_source=True,
            )
            results.append(msg)
        except Exception as exc:
            print(f"  [shell] ERRORE su {xf.name}: {exc}")

    return results


# ---------------------------------------------------------------------------
# Helper: archiviazione
# ---------------------------------------------------------------------------

def _archive_file(file_path: Path) -> None:
    """Sposta *file_path* in <parent>/archivio/<YYYYMMDD_HHMMSS>_<nome>."""
    archive_dir = file_path.parent / "archivio"
    archive_dir.mkdir(exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = archive_dir / f"{ts}_{file_path.name}"
    shutil.move(str(file_path), str(dest))
    print(f"  [shell] Archiviato: {dest.name}")


# ---------------------------------------------------------------------------
# CLI rapido (test)
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Shell Telematico GIR26")
    parser.add_argument("--xml", help="File XML GloBE da incapsulare")
    parser.add_argument("--batch", help="Directory uploads/ per elaborazione batch")
    parser.add_argument("--cf", required=True, help="Codice Fiscale fornitore")
    parser.add_argument("--output", help="Directory di output (default: stessa del file)")
    args = parser.parse_args()

    if args.xml:
        result = encapsulate(args.xml, args.cf, output_dir=args.output)
        print(f"Output: {result}")
    elif args.batch:
        results = encapsulate_uploads(args.batch, args.cf)
        print(f"Generati {len(results)} file MSG.")
    else:
        parser.error("Specificare --xml o --batch")
