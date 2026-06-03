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

import copy
import re
import shutil
import uuid
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
_CF_RE = re.compile(r"^([0-9]{11}|[A-Z0-9]{16})$")

# Regex placeholder GUID
_PLACEHOLDER_RE = re.compile(r"\{Guid:D\}", re.IGNORECASE)


def _resolve_guid_placeholders(root: etree._Element) -> etree._Element:
    """
    Restituisce un deepcopy dell'albero XML con tutti i placeholder {Guid:D}
    sostituiti con UUID4 reali. Il file sorgente non viene modificato.
    """
    root = copy.deepcopy(root)
    for el in root.iter():
        if el.text and _PLACEHOLDER_RE.search(el.text):
            el.text = _PLACEHOLDER_RE.sub(lambda _: str(uuid.uuid4()), el.text)
        if el.tail and _PLACEHOLDER_RE.search(el.tail):
            el.tail = _PLACEHOLDER_RE.sub(lambda _: str(uuid.uuid4()), el.tail)
        for attr_name in list(el.attrib):
            val = el.attrib[attr_name]
            if _PLACEHOLDER_RE.search(val):
                el.attrib[attr_name] = _PLACEHOLDER_RE.sub(
                    lambda _: str(uuid.uuid4()), val
                )
    return root


# ---------------------------------------------------------------------------
# Helper: sanitize nome file
# ---------------------------------------------------------------------------

def _sanitize_filename(name: str) -> str:
    """
    Rimuove o sostituisce caratteri non ammessi nel nome file.
    Spazi e parentesi → underscore; caratteri speciali rimossi;
    underscore multipli collassati; strip trattini/punti iniziali/finali.
    """
    name = re.sub(r"[\s\(\)]+", "_", name)   # spazi e parentesi → _
    name = re.sub(r"[^\w\-.]", "", name)       # rimuovi tutto tranne alfanumerici, _ - .
    name = re.sub(r"_+", "_", name)            # collassa underscore multipli
    name = name.strip(".-_")                   # rimuovi _ - . iniziali/finali
    return name or "output"

# ---------------------------------------------------------------------------
# Funzione principale
# ---------------------------------------------------------------------------

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

    stem = _sanitize_filename(xml_path.stem)
    out_path = out_dir / f"{stem}_MSG.xml"

    # --- Parse XML GloBE sorgente ---
    parser = etree.XMLParser(remove_blank_text=False)
    globe_tree = etree.parse(str(xml_path), parser)
    globe_root = globe_tree.getroot()  # <globe:GLOBE_OECD>

    # --- Risolvi placeholder {Guid:D} in-memory (il sorgente non viene modificato) ---
    globe_root = _resolve_guid_placeholders(globe_root)

    # --- Auto-fix version: normalizza l'attributo version su GLOBE_OECD ---
    # Il file sorgente può avere: nessun version, version="x", globe:version="x", o entrambi.
    # Lo XSD AdE richiede solo version="..." senza prefisso namespace.
    ns_globe = f"{{{NS_GLOBE}}}"
    ver = globe_root.get("version", "") or globe_root.get(f"{ns_globe}version", "")
    for attr in list(globe_root.attrib):
        if "version" in attr and attr != "version":
            del globe_root.attrib[attr]
    globe_root.set("version", ver if ver else "1.0")
    if not ver:
        print("  [shell] ⚠ GLOBE_OECD/@version assente: impostato automaticamente a '1.0'.")
    
    _NS_MAP_GLOBE = {
        "globe": NS_GLOBE,
        "stf":   NS_STF,
        "iso":   NS_ISO,
        "xsi":   NS_XSI,
    }
    _raw = etree.tostring(globe_root)
    globe_root = etree.fromstring(_raw)
    etree.cleanup_namespaces(globe_root, top_nsmap=_NS_MAP_GLOBE)

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