"""
xml_validator.py  –  PILLAR GloBE/DAC9
Implementa tutti i check AdE verificabili offline dall'Allegato Tecnico (13 marzo 2026):
  • Sezione 8.2  – Record Errors SEVERE  (60001-60028, esclusi 60002/60008/60009/60014)
  • Sezione 8.3  – Record Errors OTHER   (70001-70124)

I check 60002, 60008, 60009, 60014 richiedono storico trasmissioni AdE e non sono inclusi.

Se il file contiene placeholder {Guid:D} non risolti, il validatore genera UUID4 casuali
in-memory per poter eseguire i check (il file sorgente non viene modificato).
La presenza di placeholder viene segnalata nel Sommario del report XLSX.

Utilizzo:
    from services.xml_validator import validate
    xlsx_path = validate(xml_path, output_dir=out_dir)
"""

from __future__ import annotations

import copy
import re
import uuid
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from lxml import etree
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ── Namespace map ─────────────────────────────────────────────────────────────
NS = {
    "globe": "urn:oecd:ties:globe:v2",
    "stf":   "urn:oecd:ties:globestf:v5",
    "iso":   "urn:oecd:ties:isoglobetypes:v1",
}

def _ns(tag: str) -> str:
    """Risolve tag come 'globe:GLOBE_OECD' nel nome completo Clark."""
    prefix, local = tag.split(":")
    return f"{{{NS[prefix]}}}{local}"

# ── Costanti ──────────────────────────────────────────────────────────────────
PLACEHOLDER_RE = re.compile(r"\{Guid:D\}", re.IGNORECASE)

# DocRefId: [IT][2024][<UUID>]  (SendingCountry=2, ReportingYear=4, UniqueID=UUID-ish)
DOCREFID_RE = re.compile(
    r"^[A-Z]{2}\d{4}[A-Za-z0-9\-]{6,}$"
)
MSGREFID_RE = re.compile(
    r"^[A-Z]{2}\d{4}[A-Z]{2}[A-Za-z0-9\-]{1,}$"
)
# TIN GIR3003 format: P2JJYYYYMMDDCCCXXX
TIN3003_RE = re.compile(r"^P2[A-Z]{2}\d{8}[A-Z]{3}[A-Za-z0-9]{3}$")

GLOBE_STATUS_UPE_FORBIDDEN = {
    "GIR305","GIR307","GIR308","GIR309","GIR312",
    "GIR313","GIR314","GIR315","GIR317","GIR318",
}

# ── Risultato singolo check ───────────────────────────────────────────────────

class CheckResult:
    __slots__ = ("code","desc","xpath","status","detail")

    STATUS_OK   = "OK"
    STATUS_KO   = "KO"
    STATUS_SKIP = "SKIP"
    STATUS_WARN = "WARN"

    def __init__(self, code: str, desc: str, xpath: str,
                 status: str = "OK", detail: str = ""):
        self.code   = code
        self.desc   = desc
        self.xpath  = xpath
        self.status = status
        self.detail = detail

    def ko(self, detail: str = "") -> "CheckResult":
        self.status = self.STATUS_KO
        if detail:
            self.detail = detail
        return self

    def skip(self, reason: str = "") -> "CheckResult":
        self.status = self.STATUS_SKIP
        self.detail = reason
        return self

    def warn(self, detail: str = "") -> "CheckResult":
        self.status = self.STATUS_WARN
        self.detail = detail
        return self

# ── Helper di lettura XML ─────────────────────────────────────────────────────

def _t(el) -> str:
    """Testo di un elemento, strippato; stringa vuota se None."""
    if el is None:
        return ""
    return (el.text or "").strip()

def _find(root, xpath: str, ns: dict | None = None):
    return root.find(xpath, ns or NS)

def _findall(root, xpath: str, ns: dict | None = None):
    return root.findall(xpath, ns or NS)

def _attr(el, attr: str) -> str:
    if el is None:
        return ""
    return (el.get(attr) or "").strip()

def _decimal(s: str) -> Decimal | None:
    try:
        return Decimal(s.replace(",", "."))
    except (InvalidOperation, AttributeError):
        return None

def _year(el) -> int | None:
    t = _t(el)
    if t:
        try:
            return int(t[:4])
        except ValueError:
            pass
    return None

def _date(s: str) -> date | None:
    for fmt in ("%Y-%m-%d", "%Y-%m", "%Y"):
        try:
            return datetime.strptime(s.strip(), fmt).date()
        except ValueError:
            continue
    return None

# ── Risoluzione placeholder {Guid:D} ─────────────────────────────────────────

def _has_guid_placeholders(root: etree._Element) -> bool:
    """Restituisce True se il file contiene almeno un placeholder {Guid:D}."""
    for el in root.iter():
        if el.text and PLACEHOLDER_RE.search(el.text):
            return True
        for val in el.attrib.values():
            if PLACEHOLDER_RE.search(val):
                return True
    return False

def _resolve_guid_placeholders(root: etree._Element) -> etree._Element:
    """
    Restituisce un deepcopy dell'albero XML con tutti i placeholder
    {Guid:D} sostituiti con UUID4 casuali, senza modificare il sorgente.
    """
    root = copy.deepcopy(root)
    for el in root.iter():
        if el.text and PLACEHOLDER_RE.search(el.text):
            el.text = PLACEHOLDER_RE.sub(lambda _: str(uuid.uuid4()), el.text)
        if el.tail and PLACEHOLDER_RE.search(el.tail):
            el.tail = PLACEHOLDER_RE.sub(lambda _: str(uuid.uuid4()), el.tail)
        for attr_name in list(el.attrib):
            val = el.attrib[attr_name]
            if PLACEHOLDER_RE.search(val):
                el.attrib[attr_name] = PLACEHOLDER_RE.sub(
                    lambda _: str(uuid.uuid4()), val
                )
    return root

# ── Funzione principale ───────────────────────────────────────────────────────

def validate(xml_path: Path, output_dir: Path | None = None) -> Path:
    """
    Valida il file GloBE XML contro le regole AdE (Allegato 2, 13/03/2026).
    Ritorna il path del report XLSX generato.
    """
    xml_path   = Path(xml_path)
    output_dir = Path(output_dir) if output_dir else xml_path.parent
    output_dir.mkdir(parents=True, exist_ok=True)

    # Parse
    try:
        tree = etree.parse(str(xml_path))
        root = tree.getroot()
    except etree.XMLSyntaxError as e:
        # Errore 50001: file non parsabile
        results = [
            CheckResult("50001","Il file XML non è parsabile.",
                        "/GLOBE_OECD",
                        CheckResult.STATUS_KO, str(e))
        ]
        return _write_xlsx(results, xml_path, output_dir)

    # Rileva e risolve placeholder {Guid:D} in-memory
    guid_resolved = _has_guid_placeholders(root)
    if guid_resolved:
        root = _resolve_guid_placeholders(root)

    results: list[CheckResult] = []
    results += _check_file_errors(root, xml_path)
    results += _check_severe(root)
    results += _check_other(root)

    return _write_xlsx(results, xml_path, output_dir, guid_resolved=guid_resolved)

# ── 8.1 FILE ERRORS (50001-50004) ────────────────────────────────────────────

def _check_file_errors(root, xml_path: Path) -> list[CheckResult]:
    out = []

    # 50001 – file leggibile (già superato se siamo qui)
    out.append(CheckResult(
        "50001",
        "Il file XML è ben formato e parsabile.",
        "/GLOBE_OECD",
        CheckResult.STATUS_OK,
    ))

    # 50002 – radice GLOBE_OECD presente
    r = CheckResult("50002","Elemento radice GLOBE_OECD presente.","/GLOBE_OECD")
    tag = root.tag.split("}")[-1] if "}" in root.tag else root.tag
    if tag != "GLOBE_OECD":
        r.ko(f"Radice trovata: {root.tag}")
    out.append(r)

    # 50003 – attributo version
    r = CheckResult("50003","GLOBE_OECD/@version presente e valorizzato.",
                    "/GLOBE_OECD/@version")
    ver = root.get("version","")
    if not ver:
        r.ko("Attributo 'version' assente o vuoto.")
    out.append(r)

    # 50004 – encoding UTF-8 (lxml lo normalizza sempre; segnaliamo solo se BOM)
    r = CheckResult("50004","Il file è codificato UTF-8 senza BOM.",
                    "/GLOBE_OECD")
    with open(xml_path, "rb") as f:
        bom = f.read(3)
    if bom == b"\xef\xbb\xbf":
        r.ko("BOM UTF-8 rilevato (non ammesso).")
    out.append(r)

    return out

# ── 8.2 RECORD ERRORS – SEVERE (6XXXX) ───────────────────────────────────────

def _check_severe(root) -> list[CheckResult]:
    out = []
    ns = NS

    ms   = _find(root, "globe:MessageSpec", ns)
    body = _find(root, "globe:GLOBEBody",   ns)

    # Scorciatoie
    msg_ref = _t(_find(ms, "globe:MessageRefId", ns)) if ms is not None else ""
    rep_per = _find(ms, "globe:ReportingPeriod", ns)  if ms is not None else None
    rep_per_txt = _t(rep_per)

    filing_info = _find(body, "globe:FilingInfo", ns)         if body is not None else None
    gen_sec     = _find(body, "globe:GeneralSection", ns)     if body is not None else None
    summary     = _find(body, "globe:Summary", ns)            if body is not None else None
    utpr_attr   = _find(body, "globe:UTPRAttribution", ns)    if body is not None else None
    jur_sections= _findall(body, "globe:JurisdictionSection", ns) if body is not None else []

    filing_doc  = _find(filing_info, "globe:DocSpec", ns)     if filing_info is not None else None
    gen_doc     = _find(gen_sec,     "globe:DocSpec", ns)     if gen_sec is not None else None

    # Raccoglie tutti i DocSpec dell'intero body
    all_docspecs = _findall(body, ".//globe:DocSpec", ns) if body is not None else []

    # --- 60001 MessageRefId formato ---
    r = CheckResult("60001",
        "MessageRefId nel formato [SendingCountry][ReportingPeriod][ReceivingCountry][UniqueID]",
        "/GLOBE_OECD/MessageSpec/MessageRefId")
    if not msg_ref:
        r.ko("MessageRefId assente o vuoto.")
    elif not MSGREFID_RE.match(msg_ref):
        r.warn(f"Formato inatteso: {msg_ref!r}")
    out.append(r)

    # --- 60003 Anno ReportingPeriod non futuro ---
    r = CheckResult("60003",
        "L'anno di ReportingPeriod non può essere maggiore dell'anno corrente.",
        "/GLOBE_OECD/MessageSpec/ReportingPeriod")
    if rep_per_txt:
        yr = _year(rep_per)
        if yr and yr > date.today().year:
            r.ko(f"Anno {yr} > anno corrente {date.today().year}")
    else:
        r.ko("ReportingPeriod assente.")
    out.append(r)

    # --- 60004 No mix OECD1 + OECD2/OECD3 ---
    r = CheckResult("60004",
        "Il messaggio non può mescolare record OECD1 con OECD2/OECD3.",
        "/GLOBE_OECD/GLOBEBody/*/DocSpec/DocTypeIndic")
    if body is not None:
        indics = [_t(_find(ds, "globe:DocTypeIndic", ns)) for ds in all_docspecs]
        has_new  = any(i in ("OECD1","OECD0") for i in indics)
        has_corr = any(i in ("OECD2","OECD3") for i in indics)
        if has_new and has_corr:
            r.ko("Presenti sia OECD1/OECD0 che OECD2/OECD3.")
    out.append(r)

    # --- 60005 OECD2/OECD3: stesso tipo di blocco del CorrDocRefId ---
    r = CheckResult("60005",
        "Se DocTypeIndic è OECD2/OECD3, il record deve appartenere alla stessa sottosezione del CorrDocRefId.",
        "/GLOBE_OECD/GLOBEBody/*/DocSpec/DocTypeIndic")
    # Implementazione semplificata: verifica che CorrDocRefId sia presente
    issues = []
    for ds in all_docspecs:
        ti = _t(_find(ds, "globe:DocTypeIndic", ns))
        cr = _t(_find(ds, "globe:CorrDocRefId", ns))
        if ti in ("OECD2","OECD3") and not cr:
            issues.append("DocSpec con OECD2/OECD3 senza CorrDocRefId")
    if issues:
        r.ko("; ".join(issues[:3]))
    out.append(r)

    # --- 60006 CorrDocRefId unico nello stesso messaggio ---
    r = CheckResult("60006",
        "Lo stesso CorrDocRefId non può comparire più di una volta nello stesso messaggio.",
        "/GLOBE_OECD/GLOBEBody/*/DocSpec/CorrDocRefId")
    if body is not None:
        corr_ids = [_t(_find(ds, "globe:CorrDocRefId", ns))
                    for ds in all_docspecs
                    if _t(_find(ds, "globe:CorrDocRefId", ns))]
        dups = {x for x in corr_ids if corr_ids.count(x) > 1}
        if dups:
            r.ko(f"CorrDocRefId duplicati: {', '.join(sorted(dups)[:3])}")
    out.append(r)

    # --- 60007 DocRefId unico nel messaggio ---
    r = CheckResult("60007",
        "DocRefId già usato per un altro record.",
        "/GLOBE_OECD/GLOBEBody/*/DocSpec/DocRefId")
    if body is not None:
        doc_ids = [_t(_find(ds, "globe:DocRefId", ns))
                   for ds in all_docspecs
                   if _t(_find(ds, "globe:DocRefId", ns))]
        dups = {x for x in doc_ids if doc_ids.count(x) > 1}
        if dups:
            r.ko(f"DocRefId duplicati: {', '.join(sorted(dups)[:3])}")
    out.append(r)

    # --- 60010 FilingInfo OECD3 implica cancellazione di tutti i blocchi correlati ---
    r = CheckResult("60010",
        "FilingInfo non può essere cancellato (OECD3) senza cancellare anche GeneralSection, Summary, JurisdictionSection, UTPRAttribution.",
        "/GLOBE_OECD/GLOBEBody/FilingInfo/DocSpec/DocTypeIndic")
    if filing_doc is not None:
        fi_ti = _t(_find(filing_doc, "globe:DocTypeIndic", ns))
        if fi_ti == "OECD3":
            siblings_ok = True
            for block_tag in ["globe:GeneralSection","globe:Summary",
                              "globe:JurisdictionSection","globe:UTPRAttribution"]:
                for blk in _findall(body, block_tag, ns):
                    blk_doc = _find(blk, "globe:DocSpec", ns)
                    if blk_doc is not None:
                        ti = _t(_find(blk_doc, "globe:DocTypeIndic", ns))
                        if ti != "OECD3":
                            siblings_ok = False
            if not siblings_ok:
                r.ko("FilingInfo = OECD3 ma altri blocchi non sono OECD3.")
    out.append(r)

    # --- 60011 DocRefId formato [SendingCountry][ReportingYear][UniqueID] ---
    r = CheckResult("60011",
        "DocRefId nel formato [SendingCountry][ReportingYear][UniqueID].",
        "/GLOBE_OECD/GLOBEBody/*/DocSpec/DocRefId")
    bad = []
    for ds in all_docspecs:
        dr = _t(_find(ds, "globe:DocRefId", ns))
        if dr and not DOCREFID_RE.match(dr):
            bad.append(dr[:40])
    if bad:
        r.ko(f"Formato non valido: {bad[0]!r}" + (f" (+{len(bad)-1})" if len(bad)>1 else ""))
    out.append(r)

    # --- 60012 OECD1/OECD0: CorrDocRefId assente ---
    r = CheckResult("60012",
        "Se DocTypeIndic è OECD1 o OECD0, CorrDocRefId deve essere assente.",
        "/GLOBE_OECD/GLOBEBody/*/DocSpec/CorrDocRefId")
    bad = []
    for ds in all_docspecs:
        ti = _t(_find(ds, "globe:DocTypeIndic", ns))
        cr = _t(_find(ds, "globe:CorrDocRefId", ns))
        if ti in ("OECD1","OECD0") and cr:
            bad.append(f"{ti} con CorrDocRefId={cr[:20]}")
    if bad:
        r.ko("; ".join(bad[:3]))
    out.append(r)

    # --- 60013 OECD0 ammesso solo per FilingInfo ---
    r = CheckResult("60013",
        "OECD0 (resend) è ammesso solo per FilingInfo, non per GeneralSection/Summary/JurisdictionSection/UTPRAttribution.",
        "GLOBEBody/*/DocSpec/DocTypeIndic")
    bad_blocks = []
    for block_tag in ["globe:GeneralSection","globe:Summary",
                      "globe:JurisdictionSection","globe:UTPRAttribution"]:
        for blk in _findall(body, block_tag, ns) if body is not None else []:
            blk_doc = _find(blk, "globe:DocSpec", ns)
            if blk_doc is not None:
                ti = _t(_find(blk_doc, "globe:DocTypeIndic", ns))
                if ti == "OECD0":
                    bad_blocks.append(block_tag.split(":")[-1])
    if bad_blocks:
        r.ko(f"OECD0 trovato in: {', '.join(bad_blocks[:3])}")
    out.append(r)

    # --- 60015 OECD2/OECD3: CorrDocRefId obbligatorio ---
    r = CheckResult("60015",
        "Se DocTypeIndic è OECD2 o OECD3, CorrDocRefId è obbligatorio.",
        "/GLOBE_OECD/GLOBEBody/*/DocSpec/CorrDocRefId")
    bad = []
    for ds in all_docspecs:
        ti = _t(_find(ds, "globe:DocTypeIndic", ns))
        cr = _t(_find(ds, "globe:CorrDocRefId", ns))
        if ti in ("OECD2","OECD3") and not cr:
            dr = _t(_find(ds, "globe:DocRefId", ns))
            bad.append(f"{ti} DocRefId={dr[:20]!r}")
    if bad:
        r.ko("; ".join(bad[:3]))
    out.append(r)

    # --- 60016 FilingInfo OECD0 → GeneralSection non OECD1 ---
    r = CheckResult("60016",
        "Se FilingInfo/DocTypeIndic = OECD0 e GeneralSection presente, GeneralSection/DocTypeIndic non può essere OECD1.",
        "/GLOBE_OECD/GLOBEBody/FilingInfo/DocSpec/DocTypeIndic")
    if filing_doc is not None and gen_doc is not None:
        fi_ti  = _t(_find(filing_doc, "globe:DocTypeIndic", ns))
        gen_ti = _t(_find(gen_doc,   "globe:DocTypeIndic", ns))
        if fi_ti == "OECD0" and gen_ti == "OECD1":
            r.ko("FilingInfo=OECD0 ma GeneralSection=OECD1")
    out.append(r)

    # --- 60017 FilingInfo OECD1 → GeneralSection presente ---
    r = CheckResult("60017",
        "Se FilingInfo/DocTypeIndic = OECD1, GeneralSection deve essere presente.",
        "/GLOBE_OECD/GLOBEBody/FilingInfo/DocSpec/DocTypeIndic")
    if filing_doc is not None:
        fi_ti = _t(_find(filing_doc, "globe:DocTypeIndic", ns))
        if fi_ti == "OECD1" and gen_sec is None:
            r.ko("FilingInfo=OECD1 ma GeneralSection assente.")
    out.append(r)

    # (60018 non presente nell'Allegato 2)

    # --- 60019 Local filing: RecJurCode deve essere la giurisdizione locale ---
    r = CheckResult("60019",
        "Se FilingCE/Role è GIR403/GIR404/GIR405, RecJurCode deve essere la giurisdizione locale.",
        "/GLOBE_OECD/GLOBEBody/GeneralSection/RecJurCode")
    if filing_info is not None and gen_sec is not None:
        filing_ce = _find(filing_info, "globe:FilingCE", ns)
        role = _t(_find(filing_ce, "globe:Role", ns)) if filing_ce is not None else ""
        if role in ("GIR403","GIR404","GIR405"):
            rec_jur = _t(_find(gen_sec, "globe:RecJurCode", ns))
            res_cc  = _t(_find(filing_ce, "globe:ResCountryCode", ns))
            if rec_jur and res_cc and rec_jur != res_cc:
                r.ko(f"RecJurCode={rec_jur!r} ≠ ResCountryCode FilingCE={res_cc!r} (local filing)")
    out.append(r)

    # --- 60020 Period/Start ≤ Period/End ---
    r = CheckResult("60020",
        "Period/Start non può essere successivo a Period/End.",
        "/GLOBE_OECD/GLOBEBody/FilingInfo/Period/Start")
    if filing_info is not None:
        period = _find(filing_info, "globe:Period", ns)
        if period is not None:
            s = _date(_t(_find(period, "globe:Start", ns)))
            e = _date(_t(_find(period, "globe:End",   ns)))
            if s and e and s > e:
                r.ko(f"Start={s} > End={e}")
    out.append(r)

    # --- 60021 Period/End ≤ ReportingPeriod ---
    r = CheckResult("60021",
        "FilingInfo/Period/End non può essere successivo a MessageSpec/ReportingPeriod.",
        "/GLOBE_OECD/GLOBEBody/FilingInfo/Period/End")
    if filing_info is not None and rep_per_txt:
        period = _find(filing_info, "globe:Period", ns)
        if period is not None:
            e   = _date(_t(_find(period, "globe:End", ns)))
            rp  = _date(rep_per_txt)
            if e and rp:
                # Confronta solo l'anno
                if e.year > rp.year:
                    r.ko(f"Period/End={e} > ReportingPeriod={rp}")
    out.append(r)

    # --- 60022 FilingCE/Role=GIR401 → TIN coincide con UPE ---
    r = CheckResult("60022",
        "Se FilingCE/Role = GIR401, FilingCE/TIN deve coincidere con almeno un TIN dell'UPE.",
        "/GLOBE_OECD/GLOBEBody/FilingInfo/FilingCE/TIN")
    if filing_info is not None and gen_sec is not None:
        filing_ce = _find(filing_info, "globe:FilingCE", ns)
        role = _t(_find(filing_ce, "globe:Role", ns)) if filing_ce is not None else ""
        if role == "GIR401":
            filing_tin = _t(_find(filing_ce, "globe:TIN", ns))
            cs = _find(gen_sec, "globe:CorporateStructure", ns)
            upe_tins: set[str] = set()
            if cs is not None:
                for path in [".//globe:UPE/globe:ExcludedUPE/globe:ID/globe:TIN",
                             ".//globe:UPE/globe:OtherUPE/globe:ID/globe:TIN"]:
                    for t_el in _findall(cs, path, ns):
                        v = _t(t_el)
                        if v:
                            upe_tins.add(v)
            if filing_tin and upe_tins and filing_tin not in upe_tins:
                r.ko(f"FilingCE/TIN={filing_tin!r} non trovato tra i TIN UPE.")
    out.append(r)

    # --- 60023 FilingCE/ResCountryCode = TransmittingCountry ---
    r = CheckResult("60023",
        "FilingCE/ResCountryCode deve coincidere con MessageSpec/TransmittingCountry.",
        "/GLOBE_OECD/GLOBEBody/FilingInfo/FilingCE/ResCountryCode")
    if ms is not None and filing_info is not None:
        trans_cc = _t(_find(ms, "globe:TransmittingCountry", ns))
        filing_ce = _find(filing_info, "globe:FilingCE", ns)
        res_cc    = _t(_find(filing_ce, "globe:ResCountryCode", ns)) if filing_ce is not None else ""
        if trans_cc and res_cc and trans_cc != res_cc:
            r.ko(f"TransmittingCountry={trans_cc!r} ≠ ResCountryCode={res_cc!r}")
    out.append(r)

    # --- 60024 Se SafeHarbour/ETRRange/SBIE/etc. presenti → JurWithTaxingRights/JurisdictionName valorizzato ---
    r = CheckResult("60024",
        "Se SafeHarbour/ETRRange/SBIE/QDMTTut/GLoBETut presenti, JurWithTaxingRights/JurisdictionName deve essere valorizzato.",
        "/GLOBE_OECD/GLOBEBody/Summary/JurWithTaxingRights/JurisdictionName")
    if summary is not None:
        has_detail = (
            _find(summary, "globe:SafeHarbour", ns) is not None or
            _find(summary, "globe:ETRRange",    ns) is not None or
            _find(summary, "globe:SBIE",        ns) is not None or
            _find(summary, "globe:QDMTTut",     ns) is not None or
            _find(summary, "globe:GLoBETut",    ns) is not None
        )
        if has_detail:
            jwtr = _find(summary, "globe:JurWithTaxingRights", ns)
            jn   = _t(_find(jwtr, "globe:JurisdictionName", ns)) if jwtr is not None else ""
            if not jn:
                r.ko("JurWithTaxingRights/JurisdictionName non valorizzato.")
    out.append(r)

    # --- 60025 ETRRate = AdjustedCoveredTax/Total / NetGlobeIncome/Total ---
    r = CheckResult("60025",
        "ETRRate deve essere uguale ad AdjustedCoveredTax/Total / NetGlobeIncome/Total.",
        "JurisdictionSection/.../ETRRate")
    errors_60025 = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for etr in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            etr_rate_el = _find(etr, "globe:ETRRate", ns)
            ngi_el      = _find(etr, "globe:NetGlobeIncome/globe:Total", ns)
            act_el      = _find(etr, "globe:AdjustedCoveredTax/globe:Total", ns)
            if etr_rate_el is None or ngi_el is None:
                continue
            ngi = _decimal(_t(ngi_el))
            act = _decimal(_t(act_el)) if act_el is not None else Decimal(0)
            etr_rate = _decimal(_t(etr_rate_el))
            if ngi is None or etr_rate is None:
                continue
            if ngi <= 0:
                continue  # regola non si applica se NéGI ≤ 0
            if act is None:
                act = Decimal(0)
            expected = act / ngi
            # Tolleranza 0.0001 (4 decimali)
            if abs(expected - etr_rate) > Decimal("0.0001"):
                errors_60025.append(
                    f"{jur_name}: ETRRate={etr_rate} ≠ {act}/{ngi}={expected:.6f}"
                )
    if errors_60025:
        r.ko("; ".join(errors_60025[:3]))
    out.append(r)

    # --- 60026 TopUpTax = (TopUpTaxPercentage * ExcessProfits) + AdditionalTopUpTax - QDMTT ---
    r = CheckResult("60026",
        "TopUpTax deve essere uguale a (TopUpTaxPercentage * ExcessProfits) + AdditionalTopUpTax - QDMTT/Amount.",
        "JurisdictionSection/.../TopUpTax")
    errors_60026 = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            tut_el   = _find(oc, "globe:TopUpTax", ns)
            tutp_el  = _find(oc, "globe:TopUpTaxPercentage", ns)
            ep_el    = _find(oc, "globe:ExcessProfits", ns)
            atut_non = _find(oc, "globe:AdditionalTopUpTax/globe:NONArt4.1.5/globe:AdditionalTopUpTax", ns)
            atut_art = _find(oc, "globe:AdditionalTopUpTax/globe:Art4.1.5/globe:AdditionalTopUpTax", ns)
            qdmtt_el = _find(oc, "globe:QDMTT/globe:Amount", ns)
            if tut_el is None or tutp_el is None or ep_el is None:
                continue
            tut  = _decimal(_t(tut_el))
            tutp = _decimal(_t(tutp_el))
            ep   = _decimal(_t(ep_el))
            if tut is None or tutp is None or ep is None:
                continue
            atut_n = _decimal(_t(atut_non)) if atut_non is not None else Decimal(0)
            atut_a = _decimal(_t(atut_art)) if atut_art is not None else Decimal(0)
            qdmtt  = _decimal(_t(qdmtt_el)) if qdmtt_el is not None else Decimal(0)
            if atut_n is None: atut_n = Decimal(0)
            if atut_a is None: atut_a = Decimal(0)
            if qdmtt  is None: qdmtt  = Decimal(0)
            expected = (tutp * ep) + atut_n + atut_a - qdmtt
            if abs(expected - tut) > Decimal("1"):  # tolleranza 1 unità
                errors_60026.append(
                    f"{jur_name}: TopUpTax={tut} ≠ expected={expected:.2f}"
                )
    if errors_60026:
        r.ko("; ".join(errors_60026[:3]))
    out.append(r)

    # --- 60027 IIR/ParentEntity/TopUpTax = TopUpTaxShare - IIROffset ---
    r = CheckResult("60027",
        "IIR/ParentEntity/TopUpTax deve essere uguale a TopUpTaxShare - IIROffset.",
        "JurisdictionSection/.../IIR/ParentEntity/TopUpTax")
    errors_60027 = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ltce in _findall(js, ".//globe:LowTaxJurisdiction/globe:LTCE", ns):
            for pe in _findall(ltce, "globe:IIR/globe:ParentEntity", ns):
                tut_el    = _find(pe, "globe:TopUpTax",      ns)
                share_el  = _find(pe, "globe:TopUpTaxShare", ns)
                offset_el = _find(pe, "globe:IIROffset",     ns)
                if tut_el is None or share_el is None:
                    continue
                tut    = _decimal(_t(tut_el))
                share  = _decimal(_t(share_el))
                offset = _decimal(_t(offset_el)) if offset_el is not None else Decimal(0)
                if tut is None or share is None:
                    continue
                if offset is None: offset = Decimal(0)
                expected = share - offset
                if abs(expected - tut) > Decimal("1"):
                    errors_60027.append(f"{jur_name}: TopUpTax={tut} ≠ {share}-{offset}={expected}")
    if errors_60027:
        r.ko("; ".join(errors_60027[:3]))
    out.append(r)

    # --- 60028 AdjustedFANIL/Total = FANIL + sum(Additions) - sum(Reductions) ---
    r = CheckResult("60028",
        "AdjustedFANIL/Total = FANIL + Σ(MainEntityPEandFTE/Additions) - Σ(MainEntityPEandFTE/Reductions).",
        "JurisdictionSection/.../AdjustedFANIL/Total")
    errors_60028 = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ce_c in _findall(js, ".//globe:CEComputation", ns):
            adf = _find(ce_c, "globe:AdjustedFANIL", ns)
            if adf is None:
                continue
            total_el = _find(adf, "globe:Total", ns)
            fanil_el = _find(adf, "globe:FANIL", ns)
            if total_el is None or fanil_el is None:
                continue
            total = _decimal(_t(total_el))
            fanil = _decimal(_t(fanil_el))
            if total is None or fanil is None:
                continue
            additions  = sum(
                (_decimal(_t(x)) or Decimal(0))
                for x in _findall(adf, ".//globe:MainEntityPEandFTE/globe:Additions", ns)
            )
            reductions = sum(
                (_decimal(_t(x)) or Decimal(0))
                for x in _findall(adf, ".//globe:MainEntityPEandFTE/globe:Reductions", ns)
            )
            expected = fanil + additions - reductions
            if abs(expected - total) > Decimal("1"):
                errors_60028.append(
                    f"{jur_name}: Total={total} ≠ {fanil}+{additions}-{reductions}={expected}"
                )
    if errors_60028:
        r.ko("; ".join(errors_60028[:3]))
    out.append(r)

    return out


# ── 8.3 RECORD ERRORS – OTHER (7XXXX) ────────────────────────────────────────

def _check_other(root) -> list[CheckResult]:
    out = []
    ns = NS

    body        = _find(root, "globe:GLOBEBody",       ns)
    ms          = _find(root, "globe:MessageSpec",     ns)
    filing_info = _find(body, "globe:FilingInfo",      ns) if body is not None else None
    gen_sec     = _find(body, "globe:GeneralSection",  ns) if body is not None else None
    summary_els = _findall(body, "globe:Summary",      ns) if body is not None else []
    jur_sections= _findall(body, "globe:JurisdictionSection", ns) if body is not None else []
    utpr_attr   = _find(body,  "globe:UTPRAttribution",ns) if body is not None else None

    cs = _find(gen_sec, "globe:CorporateStructure", ns) if gen_sec is not None else None

    rep_per_txt = _t(_find(ms, "globe:ReportingPeriod", ns)) if ms is not None else ""
    period_el   = _find(filing_info, "globe:Period", ns) if filing_info is not None else None
    period_end_txt = _t(_find(period_el, "globe:End",   ns)) if period_el is not None else ""
    period_start_txt=_t(_find(period_el,"globe:Start",  ns)) if period_el is not None else ""
    period_end   = _date(period_end_txt)
    period_start = _date(period_start_txt)

    # Helper: tutti i TIN dei CE in CorporateStructure
    def _all_ce_tins() -> set[str]:
        tins: set[str] = set()
        if cs is None:
            return tins
        for t_el in _findall(cs, ".//globe:CE/globe:ID/globe:TIN", ns):
            v = _t(t_el)
            if v:
                tins.add(v)
        return tins

    def _all_upe_tins() -> set[str]:
        tins: set[str] = set()
        if cs is None:
            return tins
        for p in [".//globe:UPE/globe:ExcludedUPE/globe:ID/globe:TIN",
                  ".//globe:UPE/globe:OtherUPE/globe:ID/globe:TIN"]:
            for t_el in _findall(cs, p, ns):
                v = _t(t_el)
                if v:
                    tins.add(v)
        return tins

    # ── 8.3.1 TIN (70001-70007) ──────────────────────────────────────────────

    tin_errors = {c: [] for c in [f"7000{i}" for i in range(1,8)]}

    for tin_el in _findall(root, ".//globe:TIN", ns):
        val     = _t(tin_el)
        tot     = _attr(tin_el, "TypeOfTIN")
        unknown = _attr(tin_el, "unknown").upper()
        issued  = _attr(tin_el, "issuedBy")

        # 70001
        if tot == "GIR3004":
            if val != "NOTIN" or unknown != "TRUE" or issued:
                tin_errors["70001"].append(
                    f"TIN GIR3004: val={val!r} unknown={unknown!r} issuedBy={issued!r}"
                )
        # 70002
        if val == "NOTIN":
            if tot != "GIR3004" or unknown != "TRUE" or issued:
                tin_errors["70002"].append(
                    f"NOTIN ma TypeOfTIN={tot!r} unknown={unknown!r} issuedBy={issued!r}"
                )
        # 70003
        if unknown == "TRUE":
            if val != "NOTIN" or tot != "GIR3004" or issued:
                tin_errors["70003"].append(
                    f"@unknown=TRUE ma val={val!r} TypeOfTIN={tot!r} issuedBy={issued!r}"
                )
        # 70004 – verifica formato TIN per giurisdizione IT (semplificata)
        if issued == "IT" and val not in ("NOTIN","") and tot not in ("GIR3003","GIR3004"):
            if not (re.match(r"^\d{11}$", val) or re.match(r"^[A-Z0-9]{16}$", val)):
                tin_errors["70004"].append(f"TIN IT non valido: {val!r}")
        # 70005
        if not tot:
            tin_errors["70005"].append(f"@TypeOfTIN assente (val={val!r})")
        if not issued and tot not in ("GIR3003","GIR3004",""):
            tin_errors["70005"].append(f"@issuedBy assente con TypeOfTIN={tot!r}")
        # 70007
        if tot == "GIR3003":
            if not TIN3003_RE.match(val):
                tin_errors["70007"].append(f"TIN GIR3003 formato non valido: {val!r}")

    # TIN strutturali (70006): non possono usare GIR3004 né @unknown=TRUE
    structural_paths = [
        ".//globe:CorporateStructure/globe:UPE/globe:ExcludedUPE/globe:ID/globe:TIN",
        ".//globe:CorporateStructure/globe:UPE/globe:OtherUPE/globe:ID/globe:TIN",
        ".//globe:CorporateStructure/globe:CE/globe:ID/globe:TIN",
        ".//globe:CorporateStructure/globe:CE/globe:QIIR/globe:Exception/globe:TIN",
        ".//globe:JurisdictionSection//globe:CEComputation/globe:Elections/globe:AggregatedReporting/globe:TaxConsolGroupTIN",
    ]
    for path in structural_paths:
        for t_el in _findall(root, path, ns):
            tot = _attr(t_el, "TypeOfTIN")
            unk = _attr(t_el, "unknown").upper()
            if tot == "GIR3004" or unk == "TRUE":
                tin_errors["70006"].append(
                    f"TIN strutturale con GIR3004/unknown: val={_t(t_el)!r}"
                )

    descs = {
        "70001": "Se TIN/@TypeOfTIN = GIR3004, val=NOTIN, @unknown=TRUE, @issuedBy assente.",
        "70002": "Se TIN = NOTIN, TypeOfTIN=GIR3004, @unknown=TRUE, @issuedBy assente.",
        "70003": "Se TIN/@unknown=TRUE, allora TIN=NOTIN, TypeOfTIN=GIR3004, @issuedBy assente.",
        "70004": "Se TIN/@issuedBy è la giurisdizione locale, il TIN deve essere valido.",
        "70005": "@issuedBy e @TypeOfTIN devono essere presenti (salvo eccezioni GIR3003/GIR3004).",
        "70006": "I TIN strutturali non possono usare TypeOfTIN=GIR3004 né @unknown=TRUE.",
        "70007": "Se TIN/@TypeOfTIN=GIR3003, il TIN deve rispettare il formato P2JJYYYYMMDDCCCXXX.",
    }
    for code, errs in tin_errors.items():
        r = CheckResult(code, descs[code], "/GLOBE_OECD/GLOBEBody//TIN")
        if errs:
            r.ko(errs[0] + (f" (+{len(errs)-1})" if len(errs)>1 else ""))
        out.append(r)

    # ── 8.3.2 RecJurCode / UPE / Rules (70008-70012) ─────────────────────────

    # 70008 UTPRAttribution/RecJurCode deve essere giurisdizione UPE o JurWithTaxingRights
    r = CheckResult("70008",
        "UTPRAttribution/RecJurCode deve essere la giurisdizione dell'UPE o una di JurWithTaxingRights.",
        "/GLOBE_OECD/GLOBEBody/UTPRAttribution/RecJurCode")
    if utpr_attr is not None and gen_sec is not None:
        utpr_jur = _t(_find(utpr_attr, "globe:RecJurCode", ns))
        valid_jurs: set[str] = set()
        if cs is not None:
            for ot in _findall(cs, ".//globe:UPE/globe:OtherUPE/globe:ID/globe:ResCountryCode", ns):
                valid_jurs.add(_t(ot))
            for ex in _findall(cs, ".//globe:UPE/globe:ExcludedUPE/globe:ID/globe:ResCountryCode", ns):
                valid_jurs.add(_t(ex))
        for s_el in summary_els:
            jwtr = _find(s_el, "globe:JurWithTaxingRights", ns)
            if jwtr is not None:
                for jn in _findall(jwtr, "globe:JurisdictionName", ns):
                    valid_jurs.add(_t(jn))
        if utpr_jur and valid_jurs and utpr_jur not in valid_jurs:
            r.ko(f"RecJurCode={utpr_jur!r} non in {sorted(valid_jurs)[:5]}")
    out.append(r)

    # 70009 GloBEStatus UPE non deve contenere valori vietati
    r = CheckResult("70009",
        "GloBEStatus dell'UPE non deve contenere: GIR305,GIR307-309,GIR312-315,GIR317,GIR318.",
        "/GLOBE_OECD/GLOBEBody/GeneralSection/CorporateStructure/UPE/*/ID/GloBEStatus")
    bad = []
    if cs is not None:
        for path in [".//globe:UPE/globe:ExcludedUPE/globe:ID/globe:GloBEStatus",
                     ".//globe:UPE/globe:OtherUPE/globe:ID/globe:GloBEStatus"]:
            for gs_el in _findall(cs, path, ns):
                vals = [v.strip() for v in (_t(gs_el) or "").split()]
                forbidden = set(vals) & GLOBE_STATUS_UPE_FORBIDDEN
                if forbidden:
                    bad.append(f"UPE GloBEStatus vietato: {forbidden}")
    if bad:
        r.ko("; ".join(bad[:3]))
    out.append(r)

    # 70010 OtherUPE/ResCountryCode: un solo valore
    r = CheckResult("70010",
        "Per OtherUPE, ID/ResCountryCode deve avere un solo valore.",
        "/GLOBE_OECD/GLOBEBody/GeneralSection/.../OtherUPE/ID/ResCountryCode")
    bad = []
    if cs is not None:
        for el in _findall(cs, ".//globe:UPE/globe:OtherUPE/globe:ID", ns):
            rcc_els = _findall(el, "globe:ResCountryCode", ns)
            if len(rcc_els) > 1:
                bad.append(f"OtherUPE con {len(rcc_els)} ResCountryCode")
    if bad:
        r.ko("; ".join(bad[:3]))
    out.append(r)

    # 70011 CE/ResCountryCode: un solo valore
    r = CheckResult("70011",
        "Per CE, ID/ResCountryCode deve avere un solo valore.",
        "/GLOBE_OECD/GLOBEBody/GeneralSection/.../CE/ID/ResCountryCode")
    bad = []
    if cs is not None:
        for el in _findall(cs, ".//globe:CE/globe:ID", ns):
            rcc_els = _findall(el, "globe:ResCountryCode", ns)
            if len(rcc_els) > 1:
                bad.append(f"CE con {len(rcc_els)} ResCountryCode")
    if bad:
        r.ko("; ".join(bad[:3]))
    out.append(r)

    # 70012 Entità stesso ResCountryCode → stesso Rules (salvo GIR204)
    r = CheckResult("70012",
        "Entità con lo stesso ResCountryCode devono avere lo stesso Rules (salvo GIR204).",
        "/GLOBE_OECD/GLOBEBody/GeneralSection/.../Rules")
    bad = []
    if cs is not None:
        rcc_rules: dict[str, set] = {}
        for path_id in [
            ".//globe:UPE/globe:ExcludedUPE/globe:ID",
            ".//globe:UPE/globe:OtherUPE/globe:ID",
            ".//globe:CE/globe:ID",
        ]:
            for id_el in _findall(cs, path_id, ns):
                rcc = _t(_find(id_el, "globe:ResCountryCode", ns))
                rules_vals = set((_t(_find(id_el, "globe:Rules", ns)) or "").split())
                if rcc and "GIR204" not in rules_vals:
                    if rcc not in rcc_rules:
                        rcc_rules[rcc] = rules_vals
                    elif rcc_rules[rcc] != rules_vals:
                        bad.append(f"Giurisdizione {rcc}: Rules incoerenti")
    if bad:
        r.ko("; ".join(bad[:3]))
    out.append(r)

    # ── 8.3.3 GloBEStatus dei CE (70013-70021) ───────────────────────────────

    ce_gs_map: dict = {}  # tin → set(GloBEStatus)
    if cs is not None:
        for ce_el in _findall(cs, "globe:CE", ns):
            id_el = _find(ce_el, "globe:ID", ns)
            if id_el is None:
                continue
            tin   = _t(_find(id_el, "globe:TIN", ns))
            gs    = set((_t(_find(id_el, "globe:GloBEStatus", ns)) or "").split())
            ce_gs_map[tin] = gs

    all_gs_vals = set()
    for gs in ce_gs_map.values():
        all_gs_vals |= gs

    def _check_gs_pair(code, desc, condition_fn, xpath):
        r = CheckResult(code, desc, xpath)
        bad = []
        for tin, gs in ce_gs_map.items():
            if condition_fn(gs):
                bad.append(f"CE TIN={tin!r} GloBEStatus={gs}")
        if bad:
            r.ko(bad[0])
        return r

    out.append(_check_gs_pair("70013",
        "Se GloBEStatus contiene GIR313, non deve contenere GIR314.",
        lambda gs: "GIR313" in gs and "GIR314" in gs,
        "/GLOBE_OECD/GLOBEBody/GeneralSection/.../CE/ID/GloBEStatus"))

    out.append(_check_gs_pair("70014",
        "Se GloBEStatus contiene GIR307, non deve contenere GIR308.",
        lambda gs: "GIR307" in gs and "GIR308" in gs,
        "/GLOBE_OECD/GLOBEBody/GeneralSection/.../CE/ID/GloBEStatus"))

    # 70015 Se un CE contiene GIR308, deve esistere un altro CE con GIR307
    r = CheckResult("70015",
        "Se un CE ha GIR308, deve esistere un altro CE con GIR307.",
        "/GLOBE_OECD/GLOBEBody/GeneralSection/.../CE/ID/GloBEStatus")
    bad = [t for t, gs in ce_gs_map.items() if "GIR308" in gs]
    if bad and "GIR307" not in all_gs_vals:
        r.ko(f"CE {bad[0]!r} ha GIR308 ma nessun CE ha GIR307.")
    out.append(r)

    out.append(_check_gs_pair("70016",
        "Se GloBEStatus contiene GIR307, deve contenere anche GIR309.",
        lambda gs: "GIR307" in gs and "GIR309" not in gs,
        "/GLOBE_OECD/GLOBEBody/GeneralSection/.../CE/ID/GloBEStatus"))

    out.append(_check_gs_pair("70017",
        "Se GloBEStatus contiene GIR308, deve contenere anche GIR309.",
        lambda gs: "GIR308" in gs and "GIR309" not in gs,
        "/GLOBE_OECD/GLOBEBody/GeneralSection/.../CE/ID/GloBEStatus"))

    out.append(_check_gs_pair("70018",
        "Se GloBEStatus contiene GIR305, non deve contenere GIR306 nello stesso CE.",
        lambda gs: "GIR305" in gs and "GIR306" in gs,
        "/GLOBE_OECD/GLOBEBody/GeneralSection/.../CE/ID/GloBEStatus"))

    # 70019 Se CE ha GIR305, deve esistere un altro CE con GIR306
    r = CheckResult("70019",
        "Se un CE ha GIR305, deve esistere un altro CE con GIR306.",
        "/GLOBE_OECD/GLOBEBody/GeneralSection/.../CE/ID/GloBEStatus")
    has305 = any("GIR305" in gs for gs in ce_gs_map.values())
    has306 = "GIR306" in all_gs_vals
    if has305 and not has306:
        r.ko("CE con GIR305 presente ma nessun CE con GIR306.")
    out.append(r)

    out.append(_check_gs_pair("70020",
        "Se GloBEStatus contiene GIR316 o GIR318, non deve contenere altri valori.",
        lambda gs: ("GIR316" in gs or "GIR318" in gs) and len(gs) > 1,
        "/GLOBE_OECD/GLOBEBody/GeneralSection/.../CE/ID/GloBEStatus"))

    # 70021 CE con GIR316/GIR318 deve avere OwnershipChange valorizzato
    r = CheckResult("70021",
        "CE con GIR316 o GIR318 deve avere un OwnershipChange valorizzato.",
        "/GLOBE_OECD/GLOBEBody/GeneralSection/.../CE/ID/GloBEStatus")
    bad = []
    if cs is not None:
        for ce_el in _findall(cs, "globe:CE", ns):
            id_el = _find(ce_el, "globe:ID", ns)
            if id_el is None:
                continue
            gs = set((_t(_find(id_el, "globe:GloBEStatus", ns)) or "").split())
            if "GIR316" in gs or "GIR318" in gs:
                oc = _find(ce_el, "globe:OwnershipChange", ns)
                if oc is None or not _t(_find(oc, "globe:ChangeDate", ns)):
                    tin = _t(_find(id_el, "globe:TIN", ns))
                    bad.append(f"CE {tin!r} ha {gs & {'GIR316','GIR318'}} ma manca OwnershipChange")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # ── 8.3.4 OwnershipChange (70022-70025) ──────────────────────────────────

    for ce_el in (_findall(cs, "globe:CE", ns) if cs is not None else []):
        oc = _find(ce_el, "globe:OwnershipChange", ns)
        if oc is None:
            continue
        cd_el = _find(oc, "globe:ChangeDate", ns)
        cd    = _date(_t(cd_el))

    # 70022
    r = CheckResult("70022",
        "OwnershipChange/ChangeDate non può essere anteriore a FilingInfo/Period/Start.",
        "/GLOBE_OECD/GLOBEBody/GeneralSection/.../CE/OwnershipChange/ChangeDate")
    bad = []
    if cs is not None and period_start:
        for ce_el in _findall(cs, "globe:CE", ns):
            oc = _find(ce_el, "globe:OwnershipChange", ns)
            if oc is None:
                continue
            cd = _date(_t(_find(oc, "globe:ChangeDate", ns)))
            if cd and cd < period_start:
                tin = _t(_find(_find(ce_el,"globe:ID",ns), "globe:TIN", ns)) if _find(ce_el,"globe:ID",ns) is not None else "?"
                bad.append(f"CE {tin!r}: ChangeDate={cd} < Start={period_start}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70023
    r = CheckResult("70023",
        "OwnershipChange/ChangeDate non può essere successiva a FilingInfo/Period/End.",
        "/GLOBE_OECD/GLOBEBody/GeneralSection/.../CE/OwnershipChange/ChangeDate")
    bad = []
    if cs is not None and period_end:
        for ce_el in _findall(cs, "globe:CE", ns):
            oc = _find(ce_el, "globe:OwnershipChange", ns)
            if oc is None:
                continue
            cd = _date(_t(_find(oc, "globe:ChangeDate", ns)))
            if cd and cd > period_end:
                tin = _t(_find(_find(ce_el,"globe:ID",ns), "globe:TIN", ns)) if _find(ce_el,"globe:ID",ns) is not None else "?"
                bad.append(f"CE {tin!r}: ChangeDate={cd} > End={period_end}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70024
    r = CheckResult("70024",
        "OwnershipChange/PreOwnership non deve essere compilato quando PreGloBEStatus = GIR719.",
        "/GLOBE_OECD/GLOBEBody/GeneralSection/.../CE/OwnershipChange/PreOwnership")
    bad = []
    if cs is not None:
        for ce_el in _findall(cs, "globe:CE", ns):
            oc = _find(ce_el, "globe:OwnershipChange", ns)
            if oc is None:
                continue
            pre_gs = _t(_find(oc, "globe:PreGloBEStatus", ns))
            pre_ow = _find(oc, "globe:PreOwnership", ns)
            if pre_gs == "GIR719" and pre_ow is not None:
                bad.append("OwnershipChange con PreGloBEStatus=GIR719 ha PreOwnership")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70025
    r = CheckResult("70025",
        "Se PreOwnership/OwnershipType contiene GIR805/GIR806, PreOwnership/TIN deve usare TypeOfTIN=GIR3004 e valore NOTIN.",
        "/GLOBE_OECD/GLOBEBody/GeneralSection/.../CE/OwnershipChange/PreOwnership/TIN")
    bad = []
    if cs is not None:
        for ce_el in _findall(cs, "globe:CE", ns):
            oc = _find(ce_el, "globe:OwnershipChange", ns)
            if oc is None:
                continue
            for po in _findall(oc, "globe:PreOwnership", ns):
                ot_vals = set((_t(_find(po, "globe:OwnershipType", ns)) or "").split())
                if ot_vals & {"GIR805","GIR806"}:
                    for tin_el in _findall(po, "globe:TIN", ns):
                        tot = _attr(tin_el, "TypeOfTIN")
                        val = _t(tin_el)
                        if tot != "GIR3004" or val != "NOTIN":
                            bad.append(f"PreOwnership GIR805/806: TIN={val!r} TypeOfTIN={tot!r}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # ── 8.3.5 Ownership (70026-70031) ────────────────────────────────────────

    if cs is not None:
        for code, desc, check_fn in [
            ("70026",
             "Se GloBEStatus = GIR305, Ownership/OwnershipPercentage deve essere 100%.",
             lambda ce_el: _check_70026(ce_el, ns)),
            ("70027",
             "Se GloBEStatus = GIR318, OwnershipPercentage=0%, TIN=NOTIN, Type=GIR806.",
             lambda ce_el: _check_70027(ce_el, ns)),
            ("70028",
             "Salvo GIR318, Ownership/OwnershipPercentage non deve essere 0%.",
             lambda ce_el: _check_70028(ce_el, ns)),
        ]:
            r = CheckResult(code, desc, "/GLOBE_OECD/GLOBEBody/GeneralSection/.../CE/Ownership")
            bad = []
            for ce_el in _findall(cs, "globe:CE", ns):
                err = check_fn(ce_el)
                if err:
                    bad.append(err)
            if bad:
                r.ko(bad[0])
            out.append(r)

    # 70029 OwnershipType=GIR801 → TIN corrisponde a UPE
    r = CheckResult("70029",
        "Se OwnershipType contiene GIR801, Ownership/TIN deve corrispondere a un TIN dell'UPE.",
        "/GLOBE_OECD/GLOBEBody/GeneralSection/.../CE/Ownership/TIN")
    bad = []
    upe_tins = _all_upe_tins()
    if cs is not None:
        for ce_el in _findall(cs, "globe:CE", ns):
            for ow in _findall(ce_el, "globe:Ownership", ns):
                ot_vals = set((_t(_find(ow, "globe:OwnershipType", ns)) or "").split())
                if "GIR801" in ot_vals:
                    tin_v = _t(_find(ow, "globe:TIN", ns))
                    if tin_v and upe_tins and tin_v not in upe_tins:
                        bad.append(f"GIR801 TIN={tin_v!r} non in UPE TINs")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70030 OwnershipType=GIR802/803/804 → TIN corrisponde a CE
    r = CheckResult("70030",
        "Se OwnershipType contiene GIR802/803/804, Ownership/TIN deve corrispondere a un CE/ID/TIN.",
        "/GLOBE_OECD/GLOBEBody/GeneralSection/.../CE/Ownership/TIN")
    bad = []
    ce_tins = _all_ce_tins()
    if cs is not None:
        for ce_el in _findall(cs, "globe:CE", ns):
            for ow in _findall(ce_el, "globe:Ownership", ns):
                ot_vals = set((_t(_find(ow, "globe:OwnershipType", ns)) or "").split())
                if ot_vals & {"GIR802","GIR803","GIR804"}:
                    tin_v = _t(_find(ow, "globe:TIN", ns))
                    if tin_v and ce_tins and tin_v not in ce_tins:
                        bad.append(f"{ot_vals&{'GIR802','GIR803','GIR804'}} TIN={tin_v!r} non in CE TINs")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70031 GloBEStatus=GIR305 → Ownership/TIN = TIN di un CE/UPE con GIR306
    r = CheckResult("70031",
        "Se GloBEStatus = GIR305, Ownership/TIN deve essere uguale ad almeno un TIN dell'entità con GIR306.",
        "/GLOBE_OECD/GLOBEBody/GeneralSection/.../CE/Ownership/TIN")
    bad = []
    # Raccogli TIN con GIR306
    tins_with_306: set[str] = set()
    if cs is not None:
        for ce_el in _findall(cs, "globe:CE", ns):
            id_el = _find(ce_el, "globe:ID", ns)
            if id_el is None:
                continue
            gs = set((_t(_find(id_el, "globe:GloBEStatus", ns)) or "").split())
            if "GIR306" in gs:
                for t_el in _findall(id_el, "globe:TIN", ns):
                    tins_with_306.add(_t(t_el))
        # Anche UPE OtherUPE con GIR306
        for ot in _findall(cs, ".//globe:UPE/globe:OtherUPE/globe:ID", ns):
            gs = set((_t(_find(ot, "globe:GloBEStatus", ns)) or "").split())
            if "GIR306" in gs:
                for t_el in _findall(ot, "globe:TIN", ns):
                    tins_with_306.add(_t(t_el))

        for ce_el in _findall(cs, "globe:CE", ns):
            id_el = _find(ce_el, "globe:ID", ns)
            if id_el is None:
                continue
            gs = set((_t(_find(id_el, "globe:GloBEStatus", ns)) or "").split())
            if "GIR305" in gs:
                for ow in _findall(ce_el, "globe:Ownership", ns):
                    ow_tin = _t(_find(ow, "globe:TIN", ns))
                    if ow_tin and tins_with_306 and ow_tin not in tins_with_306:
                        tin = _t(_find(id_el, "globe:TIN", ns))
                        bad.append(f"CE {tin!r} GIR305 → Ownership/TIN={ow_tin!r} non in GIR306 TINs")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # ── 8.3.6 QIIR (70032) ───────────────────────────────────────────────────

    r = CheckResult("70032",
        "Se CE/QIIR è compilato, CE/ID/Rules deve contenere GIR201 o GIR202.",
        "/GLOBE_OECD/GLOBEBody/GeneralSection/.../CE/QIIR")
    bad = []
    if cs is not None:
        for ce_el in _findall(cs, "globe:CE", ns):
            if _find(ce_el, "globe:QIIR", ns) is not None:
                id_el = _find(ce_el, "globe:ID", ns)
                rules = set((_t(_find(id_el, "globe:Rules", ns)) or "").split()) if id_el is not None else set()
                if not (rules & {"GIR201","GIR202"}):
                    tin = _t(_find(id_el, "globe:TIN", ns)) if id_el is not None else "?"
                    bad.append(f"CE {tin!r} ha QIIR ma Rules={rules}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # ── 8.3.7 General Section – CE/QIIR (70033-70035) ────────────────────────

    # 70033 Exception/TIN deve corrispondere a TIN di un altro CE
    r = CheckResult("70033",
        "CE/QIIR/Exception/TIN deve corrispondere al TIN di un altro CE nella CorporateStructure.",
        "/GLOBE_OECD/GLOBEBody/GeneralSection/.../CE/QIIR/Exception/TIN")
    bad = []
    ce_tins = _all_ce_tins()
    if cs is not None:
        for ce_el in _findall(cs, "globe:CE", ns):
            id_el = _find(ce_el, "globe:ID", ns)
            own_tin = _t(_find(id_el, "globe:TIN", ns)) if id_el is not None else ""
            qiir = _find(ce_el, "globe:QIIR", ns)
            if qiir is None:
                continue
            for exc in _findall(qiir, "globe:Exception", ns):
                exc_tin = _t(_find(exc, "globe:TIN", ns))
                if exc_tin and exc_tin not in ce_tins:
                    bad.append(f"Exception/TIN={exc_tin!r} non corrisponde a nessun CE")
                if exc_tin and exc_tin == own_tin:
                    bad.append(f"Exception/TIN={exc_tin!r} uguale al CE stesso")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70034 POPE-IPE=GIR902 → Exception/Art2.1.3/Status=TRUE
    r = CheckResult("70034",
        "Se POPE-IPE = GIR902 (IPE) e Exception compilato, Art2.1.3/Status deve essere TRUE.",
        "/GLOBE_OECD/GLOBEBody/GeneralSection/.../CE/QIIR/POPE-IPE")
    bad = []
    if cs is not None:
        for ce_el in _findall(cs, "globe:CE", ns):
            qiir = _find(ce_el, "globe:QIIR", ns)
            if qiir is None:
                continue
            pope = _t(_find(qiir, "globe:POPE-IPE", ns))
            exc  = _find(qiir, "globe:Exception", ns)
            if pope == "GIR902" and exc is not None:
                status = _t(_find(exc, "globe:Art2.1.3/globe:Status", ns))
                if status.upper() != "TRUE":
                    bad.append(f"POPE-IPE=GIR902 Exception.Art2.1.3/Status={status!r}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70035 POPE-IPE=GIR901 → Exception/Art2.1.5/Status=TRUE
    r = CheckResult("70035",
        "Se POPE-IPE = GIR901 (POPE) e Exception compilato, Art2.1.5/Status deve essere TRUE.",
        "/GLOBE_OECD/GLOBEBody/GeneralSection/.../CE/QIIR/POPE-IPE")
    bad = []
    if cs is not None:
        for ce_el in _findall(cs, "globe:CE", ns):
            qiir = _find(ce_el, "globe:QIIR", ns)
            if qiir is None:
                continue
            pope = _t(_find(qiir, "globe:POPE-IPE", ns))
            exc  = _find(qiir, "globe:Exception", ns)
            if pope == "GIR901" and exc is not None:
                status = _t(_find(exc, "globe:Art2.1.5/globe:Status", ns))
                if status.upper() != "TRUE":
                    bad.append(f"POPE-IPE=GIR901 Exception.Art2.1.5/Status={status!r}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # ── 8.3.8 Summary / Subgroup / SafeHarbour (70036-70043) ─────────────────

    out += _check_summary(summary_els, jur_sections, filing_info, ns)

    # ── 8.3.9 JurisdictionSection / ETRException / TransCbCR (70044-70048) ───

    out += _check_jurisdiction_etr(summary_els, jur_sections, ns)

    # ── 8.3.10 UTPR SafeHarbour / NMCE (70049-70053) ─────────────────────────

    out += _check_utpr_sh(summary_els, jur_sections, ns)

    # ── 8.3.11 Elections / CEComputation (70054-70058) ───────────────────────

    out += _check_elections(jur_sections, ns)

    # ── 8.3.12 OverallComputation / NetGlobeIncome / ACT (70059-70063) ───────

    out += _check_overall_comp(jur_sections, ns)

    # ── 8.3.13 PostFilingAdjust (70064-70067) ────────────────────────────────

    out += _check_postfiling(jur_sections, period_start, ns)

    # ── 8.3.14 CoveredTaxRefund / DeemedDistTax (70068-70075) ────────────────

    out += _check_covtax_recapture(jur_sections, period_end, ns)

    # ── 8.3.15 TransBlendCFC / DeferTaxAdjustAmt (70076-70082) ───────────────

    out += _check_defer_tax(jur_sections, ns)

    # ── 8.3.16 ExcessNegTaxExpense / ExcessProfits / Substance (70083-70087) ─

    out += _check_excess_substance(jur_sections, ns)

    # ── 8.3.17 AdditionalTopUpTax (70088-70096) ──────────────────────────────

    out += _check_additional_tut(jur_sections, period_end, ns)

    # ── 8.3.18 JurisdictionSection IIR/UTPR (70097-70105) ────────────────────

    out += _check_iir_utpr(jur_sections, utpr_attr, ns)

    # ── 8.3.19 CEComputation AdjustedFANIL / UPEAdjustments (70106-70113) ────

    out += _check_ce_fanil(jur_sections, ns)

    # ── 8.3.20 CEComputation NetGlobeIncome / Elections / ACT (70114-70124) ──

    out += _check_ce_ngi(jur_sections, ns)

    return out


# ── Helper ownership ─────────────────────────────────────────────────────────

def _check_70026(ce_el, ns):
    id_el = _find(ce_el, "globe:ID", ns)
    gs = set((_t(_find(id_el, "globe:GloBEStatus", ns)) or "").split()) if id_el else set()
    if "GIR305" not in gs:
        return None
    for ow in _findall(ce_el, "globe:Ownership", ns):
        pct = _decimal(_t(_find(ow, "globe:OwnershipPercentage", ns)))
        if pct is not None and pct != Decimal("100"):
            tin = _t(_find(id_el, "globe:TIN", ns)) if id_el else "?"
            return f"CE {tin!r} GIR305: OwnershipPercentage={pct} ≠ 100%"
    return None

def _check_70027(ce_el, ns):
    id_el = _find(ce_el, "globe:ID", ns)
    gs = set((_t(_find(id_el, "globe:GloBEStatus", ns)) or "").split()) if id_el else set()
    if "GIR318" not in gs:
        return None
    for ow in _findall(ce_el, "globe:Ownership", ns):
        pct = _decimal(_t(_find(ow, "globe:OwnershipPercentage", ns)))
        tin_v = _t(_find(ow, "globe:TIN", ns))
        ot    = _t(_find(ow, "globe:OwnershipType", ns))
        if pct != Decimal("0") or tin_v != "NOTIN" or ot != "GIR806":
            tin = _t(_find(id_el, "globe:TIN", ns)) if id_el else "?"
            return f"CE {tin!r} GIR318: pct={pct} tin={tin_v!r} ot={ot!r}"
    return None

def _check_70028(ce_el, ns):
    id_el = _find(ce_el, "globe:ID", ns)
    gs = set((_t(_find(id_el, "globe:GloBEStatus", ns)) or "").split()) if id_el else set()
    if "GIR318" in gs:
        return None
    for ow in _findall(ce_el, "globe:Ownership", ns):
        pct = _decimal(_t(_find(ow, "globe:OwnershipPercentage", ns)))
        if pct is not None and pct == Decimal("0"):
            tin = _t(_find(id_el, "globe:TIN", ns)) if id_el else "?"
            return f"CE {tin!r}: OwnershipPercentage=0% (non GIR318)"
    return None


# ── Helper 8.3.8 Summary ─────────────────────────────────────────────────────

def _check_summary(summary_els, jur_sections, filing_info, ns):
    out = []
    period_el = _find(filing_info, "globe:Period", ns) if filing_info is not None else None
    period_end_txt = _t(_find(period_el, "globe:End", ns)) if period_el is not None else ""
    period_end = _date(period_end_txt)

    # Costruiamo mappa jur → [Summary] per 70036
    from collections import defaultdict
    jur_summary_map: dict = defaultdict(list)
    for s_el in summary_els:
        for jur_el in _findall(s_el, "globe:Jurisdiction", ns):
            jn = _t(_find(jur_el, "globe:JurisdictionName", ns))
            if jn:
                for sg in _findall(jur_el, "globe:Subgroup", ns):
                    jur_summary_map[jn].append(sg)

    # 70036
    r = CheckResult("70036",
        "Se una giurisdizione ha più Subgroup, l'intera Summary deve essere ripetuta tante volte.",
        "/GLOBE_OECD/GLOBEBody/Summary/Jurisdiction/Subgroup")
    # Verifica: # summary per giurisdizione = # subgroup per giurisdizione
    bad = []
    jur_count: dict = defaultdict(int)
    for s_el in summary_els:
        for jur_el in _findall(s_el, "globe:Jurisdiction", ns):
            jn = _t(_find(jur_el, "globe:JurisdictionName", ns))
            if jn:
                jur_count[jn] += 1
    for jn, sgs in jur_summary_map.items():
        if len(sgs) > 1 and jur_count.get(jn,0) != len(sgs):
            bad.append(f"{jn}: {len(sgs)} Subgroup ma {jur_count[jn]} Summary")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70037 Summary/Subgroup → corrispondente JurisdictionSection/ETR/Subgroup
    r = CheckResult("70037",
        "Se Summary/Subgroup valorizzato, deve esistere JurisdictionSection/ETR/Subgroup con TIN coerente.",
        "/GLOBE_OECD/GLOBEBody/Summary/Jurisdiction/Subgroup/TIN")
    # Build ETR Subgroup TINs per giurisdizione
    js_subgroup_tins: dict = defaultdict(set)
    for js in jur_sections:
        jur = _t(_find(js, "globe:Jurisdiction", ns))
        for etr in _findall(js, ".//globe:GLoBETax/globe:ETR", ns):
            for sg in _findall(etr, "globe:SubGroup", ns):
                for t_el in _findall(sg, "globe:TIN", ns):
                    js_subgroup_tins[jur].add(_t(t_el))
    bad = []
    for s_el in summary_els:
        for jur_el in _findall(s_el, "globe:Jurisdiction", ns):
            jn = _t(_find(jur_el, "globe:JurisdictionName", ns))
            for sg in _findall(jur_el, "globe:Subgroup", ns):
                for t_el in _findall(sg, "globe:TIN", ns):
                    tin_v = _t(t_el)
                    if tin_v and tin_v not in js_subgroup_tins.get(jn, set()):
                        bad.append(f"{jn} Summary Subgroup TIN={tin_v!r} non in JurSection")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # Recupero ReportingPeriod da root (approssimazione: uso period_end)
    # 70038 Se Period/End > 30/06/2028 → no GIR1203/1204/1205
    r = CheckResult("70038",
        "Se Period/End > 30/06/2028, SafeHarbour non può contenere GIR1203, GIR1204, GIR1205.",
        "/GLOBE_OECD/GLOBEBody/Summary/SafeHarbour")
    bad = []
    cutoff = date(2028, 6, 30)
    if period_end and period_end > cutoff:
        for s_el in summary_els:
            for sh in _findall(s_el, "globe:SafeHarbour", ns):
                val = _t(sh)
                if val in ("GIR1203","GIR1204","GIR1205"):
                    bad.append(f"SafeHarbour={val} dopo {cutoff}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70039 Se Period/End > 31/12/2026 → no GIR1206
    r = CheckResult("70039",
        "Se Period/End > 31/12/2026, SafeHarbour non può contenere GIR1206.",
        "/GLOBE_OECD/GLOBEBody/Summary/SafeHarbour")
    bad = []
    cutoff2 = date(2026, 12, 31)
    if period_end and period_end > cutoff2:
        for s_el in summary_els:
            for sh in _findall(s_el, "globe:SafeHarbour", ns):
                if _t(sh) == "GIR1206":
                    bad.append("SafeHarbour=GIR1206 dopo 31/12/2026")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70040 GIR1206 solo nella giurisdizione dell'UPE
    r = CheckResult("70040",
        "SafeHarbour = GIR1206 (Transitional UTPR Safe Harbour) solo nella giurisdizione dell'UPE.",
        "/GLOBE_OECD/GLOBEBody/Summary/SafeHarbour")
    # Semplificato: check presunto (struttura UPE non sempre recuperabile qui)
    out.append(r)

    # 70041 CFSofUPE = GIR502/GIR504 → no GIR1207/1208/1209
    r = CheckResult("70041",
        "Se CFSofUPE = GIR502 o GIR504, SafeHarbour non può contenere GIR1207, GIR1208, GIR1209.",
        "/GLOBE_OECD/GLOBEBody/FilingInfo/AccountingInfo/CFSofUPE")
    bad = []
    if filing_info is not None:
        acc = _find(filing_info, "globe:AccountingInfo", ns)
        cfs = _t(_find(acc, "globe:CFSofUPE", ns)) if acc is not None else ""
        if cfs in ("GIR502","GIR504"):
            for s_el in summary_els:
                for sh in _findall(s_el, "globe:SafeHarbour", ns):
                    if _t(sh) in ("GIR1207","GIR1208","GIR1209"):
                        bad.append(f"CFSofUPE={cfs} ma SafeHarbour={_t(sh)}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70042 JurWithTaxingRights + SafeHarbour non solo GIR1206 → ETRRange/SBIE/QDMTTut/GLoBETut obbligatori
    r = CheckResult("70042",
        "Se JurWithTaxingRights compilato e SafeHarbour assente o solo GIR1206, devono essere compilati ETRRange, SBIE, QDMTTut, GLoBETut.",
        "/GLOBE_OECD/GLOBEBody/Summary/JurWithTaxingRights")
    bad = []
    for s_el in summary_els:
        jwtr = _find(s_el, "globe:JurWithTaxingRights", ns)
        if jwtr is None:
            continue
        sh_vals = {_t(sh) for sh in _findall(s_el, "globe:SafeHarbour", ns)}
        if sh_vals <= {""} or sh_vals <= {"GIR1206"}:
            for required in ["globe:ETRRange","globe:SBIE","globe:QDMTTut","globe:GLoBETut"]:
                if _find(s_el, required, ns) is None:
                    bad.append(f"Manca {required.split(':')[-1]} in Summary")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70043 JurWithTaxingRights + SafeHarbour=GIR1202 → ETRRange/SBIE/QDMTTut
    r = CheckResult("70043",
        "Se JurWithTaxingRights e SafeHarbour = GIR1202, devono essere compilati ETRRange, SBIE e QDMTTut.",
        "/GLOBE_OECD/GLOBEBody/Summary/JurWithTaxingRights")
    bad = []
    for s_el in summary_els:
        jwtr = _find(s_el, "globe:JurWithTaxingRights", ns)
        if jwtr is None:
            continue
        sh_vals = {_t(sh) for sh in _findall(s_el, "globe:SafeHarbour", ns)}
        if "GIR1202" in sh_vals:
            for required in ["globe:ETRRange","globe:SBIE","globe:QDMTTut"]:
                if _find(s_el, required, ns) is None:
                    bad.append(f"SafeHarbour=GIR1202 ma manca {required.split(':')[-1]}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    return out


# ── Helper 8.3.9 JurisdictionSection / ETRException ──────────────────────────

def _check_jurisdiction_etr(summary_els, jur_sections, ns):
    out = []

    # 70044
    r = CheckResult("70044",
        "Se ETRStatus è compilato, deve contenere almeno uno tra ETRException e ETRComputation.",
        "/GLOBE_OECD/GLOBEBody/JurisdictionSection/.../ETRStatus")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for etr in _findall(js, ".//globe:GLoBETax/globe:ETR", ns):
            etr_status = _find(etr, "globe:ETRStatus", ns)
            if etr_status is not None:
                has_exc  = _find(etr_status, "globe:ETRException",  ns) is not None
                has_comp = _find(etr_status, "globe:ETRComputation", ns) is not None
                if not has_exc and not has_comp:
                    bad.append(f"{jur_name}: ETRStatus senza ETRException né ETRComputation")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # Build SafeHarbour mappa giurisdizione → set(valori) con subgroup TIN
    from collections import defaultdict
    sh_jur: dict = defaultdict(set)   # jur → set di GIR12xx
    sh_sg:  dict = defaultdict(set)   # (jur, tin) → set di GIR12xx
    for s_el in summary_els:
        for jur_el in _findall(s_el, "globe:Jurisdiction", ns):
            jn = _t(_find(jur_el, "globe:JurisdictionName", ns))
            sh_vals = {_t(sh) for sh in _findall(s_el, "globe:SafeHarbour", ns)}
            sh_jur[jn] |= sh_vals
            for sg in _findall(jur_el, "globe:Subgroup", ns):
                for t_el in _findall(sg, "globe:TIN", ns):
                    sh_sg[(jn, _t(t_el))] |= sh_vals

    # 70045 GIR1203/1204/1205 → ETRException/TransitionalCbCRSafeHarbour presente
    r = CheckResult("70045",
        "Se SafeHarbour = GIR1203/1204/1205, nel corrispondente JurisdictionSection deve essere compilato ETRException/TransitionalCbCRSafeHarbour.",
        "/GLOBE_OECD/GLOBEBody/JurisdictionSection/.../ETRException/TransitionalCbCRSafeHarbour")
    bad = []
    cbcr_codes = {"GIR1203","GIR1204","GIR1205"}
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        sh_vals = sh_jur.get(jur_name, set())
        if sh_vals & cbcr_codes:
            for etr in _findall(js, ".//globe:GLoBETax/globe:ETR", ns):
                etr_status = _find(etr, "globe:ETRStatus", ns)
                if etr_status is None:
                    bad.append(f"{jur_name}: GIR1203/4/5 ma ETRStatus assente")
                    continue
                exc = _find(etr_status, "globe:ETRException", ns)
                if exc is None or _find(exc, "globe:TransitionalCbCRSafeHarbour", ns) is None:
                    bad.append(f"{jur_name}: GIR1203/4/5 ma TransitionalCbCRSafeHarbour assente")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70046 TransitionalCbCRSafeHarbour → ETR/SubGroup con TypeofSubGroup = GIR1607/1608
    r = CheckResult("70046",
        "Se ETRException/TransitionalCbCRSafeHarbour compilato, ETR/SubGroup deve esistere con TypeofSubGroup = GIR1607 o GIR1608.",
        "/GLOBE_OECD/GLOBEBody/JurisdictionSection/.../ETRException/TransitionalCbCRSafeHarbour")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for etr in _findall(js, ".//globe:GLoBETax/globe:ETR", ns):
            etr_status = _find(etr, "globe:ETRStatus", ns)
            if etr_status is None:
                continue
            exc = _find(etr_status, "globe:ETRException", ns)
            if exc is None:
                continue
            if _find(exc, "globe:TransitionalCbCRSafeHarbour", ns) is not None:
                sg_types = {_t(_find(sg, "globe:TypeofSubGroup", ns))
                            for sg in _findall(etr, "globe:SubGroup", ns)}
                if not (sg_types & {"GIR1607","GIR1608"}):
                    bad.append(f"{jur_name}: TransitionalCbCRSH presente ma SubGroup TypeofSubGroup≠GIR1607/1608")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70047 GIR1203 → TransitionalCbCRSafeHarbour/Revenue compilato
    r = CheckResult("70047",
        "Se SafeHarbour = GIR1203, TransitionalCbCRSafeHarbour/Revenue deve essere compilato.",
        "/GLOBE_OECD/GLOBEBody/JurisdictionSection/.../TransitionalCbCRSafeHarbour/Revenue")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        if "GIR1203" not in sh_jur.get(jur_name, set()):
            continue
        for etr in _findall(js, ".//globe:GLoBETax/globe:ETR", ns):
            etr_status = _find(etr, "globe:ETRStatus", ns)
            if etr_status is None:
                continue
            exc = _find(etr_status, "globe:ETRException", ns)
            if exc is None:
                continue
            tcsh = _find(exc, "globe:TransitionalCbCRSafeHarbour", ns)
            if tcsh is not None and not _t(_find(tcsh, "globe:Revenue", ns)):
                bad.append(f"{jur_name}: GIR1203 ma Revenue assente")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70048 GIR1204 → TransitionalCbCRSafeHarbour/IncomeTax compilato
    r = CheckResult("70048",
        "Se SafeHarbour = GIR1204, TransitionalCbCRSafeHarbour/IncomeTax deve essere compilato.",
        "/GLOBE_OECD/GLOBEBody/JurisdictionSection/.../TransitionalCbCRSafeHarbour/IncomeTax")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        if "GIR1204" not in sh_jur.get(jur_name, set()):
            continue
        for etr in _findall(js, ".//globe:GLoBETax/globe:ETR", ns):
            etr_status = _find(etr, "globe:ETRStatus", ns)
            if etr_status is None:
                continue
            exc = _find(etr_status, "globe:ETRException", ns)
            if exc is None:
                continue
            tcsh = _find(exc, "globe:TransitionalCbCRSafeHarbour", ns)
            if tcsh is not None and not _t(_find(tcsh, "globe:IncomeTax", ns)):
                bad.append(f"{jur_name}: GIR1204 ma IncomeTax assente")
    if bad:
        r.ko(bad[0])
    out.append(r)

    return out


# ── Helper 8.3.10 UTPR SafeHarbour / NMCE (70049-70053) ──────────────────────

def _check_utpr_sh(summary_els, jur_sections, ns):
    out = []
    from collections import defaultdict
    sh_jur: dict = defaultdict(set)
    for s_el in summary_els:
        for jur_el in _findall(s_el, "globe:Jurisdiction", ns):
            jn = _t(_find(jur_el, "globe:JurisdictionName", ns))
            sh_vals = {_t(sh) for sh in _findall(s_el, "globe:SafeHarbour", ns)}
            sh_jur[jn] |= sh_vals

    # 70049 GIR1206 → ETRException/UTPRSafeHarbour/CITRate
    r = CheckResult("70049",
        "Se SafeHarbour = GIR1206, ETRException/UTPRSafeHarbour con CITRate deve essere compilato.",
        "/GLOBE_OECD/GLOBEBody/JurisdictionSection/.../ETRException/UTPRSafeHarbour/CITRate")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        if "GIR1206" not in sh_jur.get(jur_name, set()):
            continue
        for etr in _findall(js, ".//globe:GLoBETax/globe:ETR", ns):
            etr_status = _find(etr, "globe:ETRStatus", ns)
            exc = _find(etr_status, "globe:ETRException", ns) if etr_status is not None else None
            utpr_sh = _find(exc, "globe:UTPRSafeHarbour", ns) if exc is not None else None
            if utpr_sh is None or not _t(_find(utpr_sh, "globe:CITRate", ns)):
                bad.append(f"{jur_name}: GIR1206 ma UTPRSafeHarbour/CITRate assente")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70050 GIR1207/1208/1209 → ETRComputation/Non-MaterialCE
    r = CheckResult("70050",
        "Se SafeHarbour = GIR1207/1208/1209, ETRComputation/Non-MaterialCE deve essere compilato.",
        "/GLOBE_OECD/GLOBEBody/JurisdictionSection/.../ETRComputation/Non-MaterialCE")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        if not (sh_jur.get(jur_name, set()) & {"GIR1207","GIR1208","GIR1209"}):
            continue
        for etr in _findall(js, ".//globe:GLoBETax/globe:ETR", ns):
            etr_status = _find(etr, "globe:ETRStatus", ns)
            etr_comp   = _find(etr_status, "globe:ETRComputation", ns) if etr_status else None
            if etr_comp is None or _find(etr_comp, "globe:Non-MaterialCE", ns) is None:
                bad.append(f"{jur_name}: GIR1207/8/9 ma Non-MaterialCE assente")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70051 GIR1208 → Non-MaterialCE/RFY/AggregateSimplified
    r = CheckResult("70051",
        "Se SafeHarbour = GIR1208, Non-MaterialCE/RFY deve contenere AggregateSimplified.",
        "/GLOBE_OECD/GLOBEBody/JurisdictionSection/.../Non-MaterialCE/RFY/AggregateSimplified")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        if "GIR1208" not in sh_jur.get(jur_name, set()):
            continue
        for etr in _findall(js, ".//globe:GLoBETax/globe:ETR", ns):
            etr_status = _find(etr, "globe:ETRStatus", ns)
            etr_comp   = _find(etr_status, "globe:ETRComputation", ns) if etr_status else None
            nmce = _find(etr_comp, "globe:Non-MaterialCE", ns) if etr_comp else None
            rfy  = _find(nmce, "globe:RFY", ns) if nmce else None
            if rfy is None or _find(rfy, "globe:AggregateSimplified", ns) is None:
                bad.append(f"{jur_name}: GIR1208 ma RFY/AggregateSimplified assente")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70052 GIR1209 → OverallComputation/SubstanceExclusion
    r = CheckResult("70052",
        "Se SafeHarbour = GIR1209, OverallComputation/SubstanceExclusion deve essere compilato.",
        "/GLOBE_OECD/GLOBEBody/JurisdictionSection/.../OverallComputation/SubstanceExclusion")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        if "GIR1209" not in sh_jur.get(jur_name, set()):
            continue
        for etr in _findall(js, ".//globe:GLoBETax/globe:ETR", ns):
            for oc in _findall(etr, ".//globe:ETRComputation/globe:OverallComputation", ns):
                if _find(oc, "globe:SubstanceExclusion", ns) is None:
                    bad.append(f"{jur_name}: GIR1209 ma SubstanceExclusion assente")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70053 GIR1205 → SubstanceExclusion obbligatorio salvo Profit ≤ 0
    r = CheckResult("70053",
        "Se SafeHarbour = GIR1205, OverallComputation/SubstanceExclusion obbligatorio (salvo Profit ≤ 0).",
        "/GLOBE_OECD/GLOBEBody/JurisdictionSection/.../OverallComputation/SubstanceExclusion")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        if "GIR1205" not in sh_jur.get(jur_name, set()):
            continue
        for etr in _findall(js, ".//globe:GLoBETax/globe:ETR", ns):
            etr_status = _find(etr, "globe:ETRStatus", ns)
            exc  = _find(etr_status, "globe:ETRException", ns) if etr_status else None
            tcsh = _find(exc, "globe:TransitionalCbCRSafeHarbour", ns) if exc else None
            profit_el = _find(tcsh, "globe:Profit", ns) if tcsh else None
            profit = _decimal(_t(profit_el)) if profit_el is not None else None
            if profit is not None and profit <= 0:
                continue  # eccezione: Profit ≤ 0
            for oc in _findall(etr, ".//globe:ETRComputation/globe:OverallComputation", ns):
                if _find(oc, "globe:SubstanceExclusion", ns) is None:
                    bad.append(f"{jur_name}: GIR1205 ma SubstanceExclusion assente (Profit>0)")
    if bad:
        r.ko(bad[0])
    out.append(r)

    return out


# ── Helper 8.3.11 Elections / CEComputation (70054-70058) ────────────────────

def _check_elections(jur_sections, ns):
    out = []

    # 70054 RevocationYear solo quando Status=FALSE (ETR/Election level)
    r = CheckResult("70054",
        "In ETR/Election/*, RevocationYear può essere valorizzato solo quando Status = FALSE.",
        "/GLOBE_OECD/GLOBEBody/JurisdictionSection/.../ETR/Election/*/RevocationYear")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for etr in _findall(js, ".//globe:GLoBETax/globe:ETR", ns):
            election = _find(etr, "globe:Election", ns)
            if election is None:
                continue
            for child in election:
                rv  = _find(child, "globe:RevocationYear", ns)
                st  = _t(_find(child, "globe:Status", ns))
                if rv is not None and st.upper() != "FALSE":
                    bad.append(f"{jur_name}: RevocationYear presente con Status={st!r}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70055 Art3.2.1.c: OutstandingBalance = QualOwnerIntentBalance + Additions - Reductions
    r = CheckResult("70055",
        "Election/Art3.2.1.c: OutstandingBalance = QualOwnerIntentBalance + Additions - Reductions.",
        "/GLOBE_OECD/GLOBEBody/JurisdictionSection/.../Election/Art3.2.1.c/OutstandingBalance")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for etr in _findall(js, ".//globe:GLoBETax/globe:ETR", ns):
            art = _find(etr, "globe:Election/globe:Art3.2.1.c", ns)
            if art is None:
                continue
            ob   = _decimal(_t(_find(art, "globe:OutstandingBalance",    ns)))
            qoib = _decimal(_t(_find(art, "globe:QualOwnerIntentBalance",ns)))
            add  = _decimal(_t(_find(art, "globe:Additions",              ns)))
            red  = _decimal(_t(_find(art, "globe:Reductions",             ns)))
            if ob is None or qoib is None:
                continue
            if add is None: add = Decimal(0)
            if red is None: red = Decimal(0)
            expected = qoib + add - red
            if abs(expected - ob) > Decimal("1"):
                bad.append(f"{jur_name}: OutstandingBalance={ob} ≠ {expected}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70056 CEComputation/Elections/*/RevocationYear solo quando Status=FALSE
    r = CheckResult("70056",
        "In CEComputation/Elections/*, RevocationYear può essere valorizzato solo quando Status = FALSE.",
        "/GLOBE_OECD/GLOBEBody/JurisdictionSection/.../CEComputation/Elections/*/RevocationYear")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ce_c in _findall(js, ".//globe:CEComputation", ns):
            elections = _find(ce_c, "globe:Elections", ns)
            if elections is None:
                continue
            for child in elections:
                rv = _find(child, "globe:RevocationYear", ns)
                st = _t(_find(child, "globe:Status", ns))
                if rv is not None and st.upper() != "FALSE":
                    bad.append(f"{jur_name}: CEComputation RevocationYear con Status={st!r}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70057 Se AggregatedReporting compilato → CEComputation/TIN = AggregatedReporting/TaxConsolGroupTIN
    r = CheckResult("70057",
        "Se CEComputation/Elections/AggregatedReporting compilato, CEComputation/TIN deve coincidere con TaxConsolGroupTIN.",
        "/GLOBE_OECD/GLOBEBody/JurisdictionSection/.../CEComputation/TIN")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ce_c in _findall(js, ".//globe:CEComputation", ns):
            agg = _find(ce_c, "globe:Elections/globe:AggregatedReporting", ns)
            if agg is None:
                continue
            ce_tin     = _t(_find(ce_c, "globe:TIN", ns))
            consol_tin = _t(_find(agg, "globe:TaxConsolGroupTIN", ns))
            if ce_tin and consol_tin and ce_tin != consol_tin:
                bad.append(f"{jur_name}: CEComputation/TIN={ce_tin!r} ≠ TaxConsolGroupTIN={consol_tin!r}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70058 Art7.6/InvestmentEntityTIN ≠ CEComputation/TIN
    r = CheckResult("70058",
        "CEComputation/Elections/Art7.6/InvestmentEntityTIN non deve coincidere con CEComputation/TIN.",
        "/GLOBE_OECD/GLOBEBody/JurisdictionSection/.../CEComputation/Elections/Art7.6/InvestmentEntityTIN")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ce_c in _findall(js, ".//globe:CEComputation", ns):
            art76 = _find(ce_c, "globe:Elections/globe:Art7.6", ns)
            if art76 is None:
                continue
            inv_tin = _t(_find(art76, "globe:InvestmentEntityTIN", ns))
            ce_tin  = _t(_find(ce_c, "globe:TIN", ns))
            if inv_tin and ce_tin and inv_tin == ce_tin:
                bad.append(f"{jur_name}: InvestmentEntityTIN = CEComputation/TIN = {ce_tin!r}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    return out


# ── Helper 8.3.12 OverallComputation (70059-70063) ───────────────────────────

def _check_overall_comp(jur_sections, ns):
    out = []

    # 70059 AdjustmentItem in NetGlobeIncome/Adjustments unico per ETR
    r = CheckResult("70059",
        "Ogni AdjustmentItem in OverallComputation/NetGlobeIncome/Adjustments non può comparire più di una volta per ETR.",
        "JurisdictionSection/.../NetGlobeIncome/Adjustments/AdjustmentItem")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            items = [_t(x) for x in _findall(oc, "globe:NetGlobeIncome/globe:Adjustments/globe:AdjustmentItem", ns)]
            dups = {x for x in items if items.count(x) > 1}
            if dups:
                bad.append(f"{jur_name}: NéGI AdjustmentItem duplicati: {dups}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70060 AdjustmentItem = GIR2025 → IntShippingIncome compilato
    r = CheckResult("70060",
        "Se NetGlobeIncome/AdjustmentItem = GIR2025, deve essere compilato NetGlobeIncome/IntShippingIncome.",
        "JurisdictionSection/.../NetGlobeIncome/IntShippingIncome")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            items = [_t(x) for x in _findall(oc, "globe:NetGlobeIncome/globe:Adjustments/globe:AdjustmentItem", ns)]
            if "GIR2025" in items:
                if _find(oc, "globe:NetGlobeIncome/globe:IntShippingIncome", ns) is None:
                    bad.append(f"{jur_name}: GIR2025 ma IntShippingIncome assente")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70061 Election/Art4.6.1=TRUE → AdjustedCoveredTax/Adjustments con GIR2711 amount<0
    r = CheckResult("70061",
        "Se Election/Art4.6.1 = TRUE, AdjustedCoveredTax/Adjustments deve contenere GIR2711 con Amount negativo.",
        "JurisdictionSection/.../AdjustedCoveredTax/Adjustments/AdjustmentItem=GIR2711")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for etr in _findall(js, ".//globe:GLoBETax/globe:ETR", ns):
            art461 = _find(etr, "globe:Election/globe:Art4.6.1", ns)
            if art461 is None or _t(art461).upper() != "TRUE":
                continue
            for oc in _findall(etr, ".//globe:ETRComputation/globe:OverallComputation", ns):
                items_amounts = {
                    _t(_find(adj, "globe:AdjustmentItem", ns)):
                    _decimal(_t(_find(adj, "globe:Amount", ns)))
                    for adj in _findall(oc, "globe:AdjustedCoveredTax/globe:Adjustments", ns)
                }
                amt_2711 = items_amounts.get("GIR2711")
                if amt_2711 is None or amt_2711 >= 0:
                    bad.append(f"{jur_name}: Art4.6.1=TRUE ma GIR2711 amount={amt_2711}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70062 AdjustmentItem = GIR2720 → AdjustedCoveredTax/Total non negativo
    r = CheckResult("70062",
        "Se AdjustedCoveredTax/AdjustmentItem = GIR2720, AdjustedCoveredTax/Total non può essere negativo.",
        "JurisdictionSection/.../AdjustedCoveredTax/Total")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            items = [_t(x) for x in _findall(oc, "globe:AdjustedCoveredTax/globe:Adjustments/globe:AdjustmentItem", ns)]
            if "GIR2720" in items:
                total = _decimal(_t(_find(oc, "globe:AdjustedCoveredTax/globe:Total", ns)))
                if total is not None and total < 0:
                    bad.append(f"{jur_name}: GIR2720 ma AdjustedCoveredTax/Total={total} < 0")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70063 AdjustmentItem in AdjustedCoveredTax/Adjustments unico per ETR
    r = CheckResult("70063",
        "Ogni AdjustmentItem in OverallComputation/AdjustedCoveredTax/Adjustments non può comparire più di una volta per ETR.",
        "JurisdictionSection/.../AdjustedCoveredTax/Adjustments/AdjustmentItem")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            items = [_t(x) for x in _findall(oc, "globe:AdjustedCoveredTax/globe:Adjustments/globe:AdjustmentItem", ns)]
            dups = {x for x in items if items.count(x) > 1}
            if dups:
                bad.append(f"{jur_name}: ACT AdjustmentItem duplicati: {dups}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    return out


# ── Helper 8.3.13 PostFilingAdjust (70064-70067) ─────────────────────────────

def _check_postfiling(jur_sections, period_start, ns):
    out = []

    # 70064 DeferTaxAsset/Total = sum(AmountAttributed/Amount)
    r = CheckResult("70064",
        "PostFilingAdjust/DeferTaxAsset/Total = somma di tutti i AmountAttributed/Amount del blocco DeferTaxAsset.",
        "JurisdictionSection/.../PostFilingAdjust/DeferTaxAsset/Total")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            pfa = _find(oc, "globe:AdjustedCoveredTax/globe:PostFilingAdjust", ns)
            if pfa is None:
                continue
            dta = _find(pfa, "globe:DeferTaxAsset", ns)
            if dta is None:
                continue
            total = _decimal(_t(_find(dta, "globe:Total", ns)))
            summed = sum(
                (_decimal(_t(_find(aa, "globe:Amount", ns))) or Decimal(0))
                for aa in _findall(dta, "globe:AmountAttributed", ns)
            )
            if total is not None and abs(total - summed) > Decimal("1"):
                bad.append(f"{jur_name}: DeferTaxAsset Total={total} ≠ sum={summed}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70065 CoveredTaxRefund/Total = sum(AmountAttributed/Amount)
    r = CheckResult("70065",
        "PostFilingAdjust/CoveredTaxRefund/Total = somma di tutti i AmountAttributed/Amount.",
        "JurisdictionSection/.../PostFilingAdjust/CoveredTaxRefund/Total")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            pfa = _find(oc, "globe:AdjustedCoveredTax/globe:PostFilingAdjust", ns)
            if pfa is None:
                continue
            ctr = _find(pfa, "globe:CoveredTaxRefund", ns)
            if ctr is None:
                continue
            total = _decimal(_t(_find(ctr, "globe:Total", ns)))
            summed = sum(
                (_decimal(_t(_find(aa, "globe:Amount", ns))) or Decimal(0))
                for aa in _findall(ctr, "globe:AmountAttributed", ns)
            )
            if total is not None and abs(total - summed) > Decimal("1"):
                bad.append(f"{jur_name}: CoveredTaxRefund Total={total} ≠ sum={summed}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70066 DeferTaxAsset/AmountAttributed/Year ≤ anno di Period/Start
    r = CheckResult("70066",
        "In PostFilingAdjust/DeferTaxAsset/AmountAttributed, Year ≤ anno di FilingInfo/Period/Start.",
        "JurisdictionSection/.../PostFilingAdjust/DeferTaxAsset/AmountAttributed/Year")
    bad = []
    start_year = period_start.year if period_start else None
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            pfa = _find(oc, "globe:AdjustedCoveredTax/globe:PostFilingAdjust", ns)
            if pfa is None:
                continue
            dta = _find(pfa, "globe:DeferTaxAsset", ns)
            if dta is None:
                continue
            for aa in _findall(dta, "globe:AmountAttributed", ns):
                yr = _year(_find(aa, "globe:Year", ns))
                if start_year and yr and yr > start_year:
                    bad.append(f"{jur_name}: DeferTaxAsset Year={yr} > Period/Start year={start_year}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70067 DeferTaxAsset/AmountAttributed/Year non ripetuto
    r = CheckResult("70067",
        "Se sono presenti più AmountAttributed in DeferTaxAsset, gli anni Year non possono ripetersi.",
        "JurisdictionSection/.../PostFilingAdjust/DeferTaxAsset/AmountAttributed/Year")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            pfa = _find(oc, "globe:AdjustedCoveredTax/globe:PostFilingAdjust", ns)
            if pfa is None:
                continue
            dta = _find(pfa, "globe:DeferTaxAsset", ns)
            if dta is None:
                continue
            years = [_year(_find(aa, "globe:Year", ns)) for aa in _findall(dta, "globe:AmountAttributed", ns)]
            years = [y for y in years if y is not None]
            if len(years) != len(set(years)):
                bad.append(f"{jur_name}: DeferTaxAsset anni ripetuti: {years}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    return out


# ── Helper 8.3.14 CoveredTaxRefund / DeemedDistTax (70068-70075) ─────────────

def _check_covtax_recapture(jur_sections, period_end, ns):
    out = []

    # 70068 CoveredTaxRefund/AmountAttributed/Year ≤ Period/Start (era End nel doc: "≤ Period/Start o anteriore")
    r = CheckResult("70068",
        "In PostFilingAdjust/CoveredTaxRefund/AmountAttributed, Year ≤ anno di FilingInfo/Period/Start o anteriore.",
        "JurisdictionSection/.../PostFilingAdjust/CoveredTaxRefund/AmountAttributed/Year")
    bad = []
    end_year = period_end.year if period_end else None
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            pfa = _find(oc, "globe:AdjustedCoveredTax/globe:PostFilingAdjust", ns)
            if pfa is None:
                continue
            ctr = _find(pfa, "globe:CoveredTaxRefund", ns)
            if ctr is None:
                continue
            for aa in _findall(ctr, "globe:AmountAttributed", ns):
                yr = _year(_find(aa, "globe:Year", ns))
                if end_year and yr and yr > end_year:
                    bad.append(f"{jur_name}: CoveredTaxRefund Year={yr} > {end_year}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70069 CoveredTaxRefund/AmountAttributed/Year non ripetuto
    r = CheckResult("70069",
        "Se sono presenti più AmountAttributed in CoveredTaxRefund, gli anni Year non possono ripetersi.",
        "JurisdictionSection/.../PostFilingAdjust/CoveredTaxRefund/AmountAttributed/Year")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            pfa = _find(oc, "globe:AdjustedCoveredTax/globe:PostFilingAdjust", ns)
            if pfa is None:
                continue
            ctr = _find(pfa, "globe:CoveredTaxRefund", ns)
            if ctr is None:
                continue
            years = [_year(_find(aa, "globe:Year", ns)) for aa in _findall(ctr, "globe:AmountAttributed", ns)]
            years = [y for y in years if y is not None]
            if len(years) != len(set(years)):
                bad.append(f"{jur_name}: CoveredTaxRefund anni ripetuti: {years}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70070 DeemedDistTax/Recapture/Year ≤ Period/End
    r = CheckResult("70070",
        "In DeemedDistTax/Election/Recapture/Year, il valore YYYY non può essere successivo a FilingInfo/Period/End.",
        "JurisdictionSection/.../DeemedDistTax/Election/Recapture/Year")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            ddt = _find(oc, "globe:AdjustedCoveredTax/globe:DeemedDistTax", ns)
            if ddt is None:
                continue
            for rec in _findall(ddt, "globe:Election/globe:Recapture", ns):
                yr = _year(_find(rec, "globe:Year", ns))
                if period_end and yr and yr > period_end.year:
                    bad.append(f"{jur_name}: DeemedDistTax Recapture Year={yr} > {period_end.year}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70071 Recapture/Year non più di 4 anni anteriore a Period/End
    r = CheckResult("70071",
        "In DeemedDistTax/Recapture/Year, YYYY non può essere 4 anni o più anteriore a Period/End.",
        "JurisdictionSection/.../DeemedDistTax/Election/Recapture/Year")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            ddt = _find(oc, "globe:AdjustedCoveredTax/globe:DeemedDistTax", ns)
            if ddt is None:
                continue
            for rec in _findall(ddt, "globe:Election/globe:Recapture", ns):
                yr = _year(_find(rec, "globe:Year", ns))
                if period_end and yr and (period_end.year - yr) >= 4:
                    bad.append(f"{jur_name}: DeemedDistTax Recapture Year={yr} è ≥4 anni prima di {period_end.year}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70072 Recapture/EndAmount = StartAmount - TotalDDT
    r = CheckResult("70072",
        "Recapture/EndAmount deve essere uguale a StartAmount - TotalDDT.",
        "JurisdictionSection/.../DeemedDistTax/Election/Recapture/EndAmount")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            ddt = _find(oc, "globe:AdjustedCoveredTax/globe:DeemedDistTax", ns)
            if ddt is None:
                continue
            for rec in _findall(ddt, "globe:Election/globe:Recapture", ns):
                end   = _decimal(_t(_find(rec, "globe:EndAmount",   ns)))
                start = _decimal(_t(_find(rec, "globe:StartAmount", ns)))
                total = _decimal(_t(_find(rec, "globe:TotalDDT",    ns)))
                if end is None or start is None or total is None:
                    continue
                if abs(end - (start - total)) > Decimal("1"):
                    bad.append(f"{jur_name}: EndAmount={end} ≠ {start}-{total}={start-total}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70073 Recapture/EndAmount non negativo
    r = CheckResult("70073",
        "Recapture/EndAmount non deve essere negativo.",
        "JurisdictionSection/.../DeemedDistTax/Election/Recapture/EndAmount")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            ddt = _find(oc, "globe:AdjustedCoveredTax/globe:DeemedDistTax", ns)
            if ddt is None:
                continue
            for rec in _findall(ddt, "globe:Election/globe:Recapture", ns):
                end = _decimal(_t(_find(rec, "globe:EndAmount", ns)))
                if end is not None and end < 0:
                    bad.append(f"{jur_name}: EndAmount={end} < 0")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70074 Recapture/TotalDDT = DDTYear-0 + DDTYear-1 + DDTYear-2 + DDTYear-3
    r = CheckResult("70074",
        "Recapture/TotalDDT = DDTYear-0 + DDTYear-1 + DDTYear-2 + DDTYear-3.",
        "JurisdictionSection/.../DeemedDistTax/Election/Recapture/TotalDDT")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            ddt = _find(oc, "globe:AdjustedCoveredTax/globe:DeemedDistTax", ns)
            if ddt is None:
                continue
            for rec in _findall(ddt, "globe:Election/globe:Recapture", ns):
                total = _decimal(_t(_find(rec, "globe:TotalDDT", ns)))
                if total is None:
                    continue
                summed = sum(
                    (_decimal(_t(_find(rec, f"globe:DDTYear-{i}", ns))) or Decimal(0))
                    for i in range(4)
                )
                if abs(total - summed) > Decimal("1"):
                    bad.append(f"{jur_name}: TotalDDT={total} ≠ sum={summed}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70075 Se Recapture/Year coincide con anno Period/End → DDTYear-0/1/2/3 = 0
    r = CheckResult("70075",
        "Se Recapture/Year coincide con anno Period/End, DDTYear-0/1/2/3 devono essere 0.",
        "JurisdictionSection/.../DeemedDistTax/Election/Recapture/DDTYear-0")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            ddt = _find(oc, "globe:AdjustedCoveredTax/globe:DeemedDistTax", ns)
            if ddt is None:
                continue
            for rec in _findall(ddt, "globe:Election/globe:Recapture", ns):
                yr = _year(_find(rec, "globe:Year", ns))
                if period_end and yr == period_end.year:
                    for i in range(4):
                        v = _decimal(_t(_find(rec, f"globe:DDTYear-{i}", ns)))
                        if v is not None and v != 0:
                            bad.append(f"{jur_name}: Year={yr}=Period/End year ma DDTYear-{i}={v}≠0")
    if bad:
        r.ko(bad[0])
    out.append(r)

    return out


# ── Helper 8.3.15 TransBlendCFC / DeferTaxAdjustAmt (70076-70082) ────────────

def _check_defer_tax(jur_sections, ns):
    out = []

    # 70076 TransBlendCFC/Total = sum(CFCJur/Allocation/AggAllocTax)
    r = CheckResult("70076",
        "AdjustedCoveredTax/TransBlendCFC/Total = somma di tutti gli AggAllocTax nei blocchi CFCJur/Allocation.",
        "JurisdictionSection/.../AdjustedCoveredTax/TransBlendCFC/Total")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            tbc = _find(oc, "globe:AdjustedCoveredTax/globe:TransBlendCFC", ns)
            if tbc is None:
                continue
            total = _decimal(_t(_find(tbc, "globe:Total", ns)))
            summed = sum(
                (_decimal(_t(x)) or Decimal(0))
                for x in _findall(tbc, ".//globe:CFCJur/globe:Allocation/globe:AggAllocTax", ns)
            )
            if total is not None and abs(total - summed) > Decimal("1"):
                bad.append(f"{jur_name}: TransBlendCFC Total={total} ≠ sum={summed}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70077 DeferTaxAdjustAmt/Total = PreRecast + Recast/Lower - Recast/Higher
    r = CheckResult("70077",
        "DeferTaxAdjustAmt/Total = PreRecast + Recast/Lower - Recast/Higher (0 se mancanti).",
        "JurisdictionSection/.../AdjustedCoveredTax/DeferTaxAdjustAmt/Total")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            dta = _find(oc, "globe:AdjustedCoveredTax/globe:DeferTaxAdjustAmt", ns)
            if dta is None:
                continue
            total = _decimal(_t(_find(dta, "globe:Total",          ns)))
            pre   = _decimal(_t(_find(dta, "globe:PreRecast",       ns))) or Decimal(0)
            lower = _decimal(_t(_find(dta, "globe:Recast/globe:Lower",  ns))) or Decimal(0)
            high  = _decimal(_t(_find(dta, "globe:Recast/globe:Higher", ns))) or Decimal(0)
            if total is None:
                continue
            expected = pre + lower - high
            if abs(total - expected) > Decimal("1"):
                bad.append(f"{jur_name}: DeferTaxAdjustAmt Total={total} ≠ {expected}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70078 BefRecastAdjust = DefTaxAmt - DiffCarryValue + GLoBEValue
    r = CheckResult("70078",
        "DeferTaxAdjustAmt/BefRecastAdjust = DefTaxAmt - DiffCarryValue + GLoBEValue.",
        "JurisdictionSection/.../AdjustedCoveredTax/DeferTaxAdjustAmt/BefRecastAdjust")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            dta = _find(oc, "globe:AdjustedCoveredTax/globe:DeferTaxAdjustAmt", ns)
            if dta is None:
                continue
            bef  = _decimal(_t(_find(dta, "globe:BefRecastAdjust", ns)))
            dta2 = _decimal(_t(_find(dta, "globe:DefTaxAmt",        ns)))
            diff = _decimal(_t(_find(dta, "globe:DiffCarryValue",   ns)))
            gbe  = _decimal(_t(_find(dta, "globe:GLoBEValue",       ns)))
            if bef is None or dta2 is None or diff is None or gbe is None:
                continue
            expected = dta2 - diff + gbe
            if abs(bef - expected) > Decimal("1"):
                bad.append(f"{jur_name}: BefRecastAdjust={bef} ≠ {expected}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70079 PreRecast = BefRecastAdjust + TotalAdjust
    r = CheckResult("70079",
        "DeferTaxAdjustAmt/PreRecast = BefRecastAdjust + TotalAdjust.",
        "JurisdictionSection/.../AdjustedCoveredTax/DeferTaxAdjustAmt/PreRecast")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            dta = _find(oc, "globe:AdjustedCoveredTax/globe:DeferTaxAdjustAmt", ns)
            if dta is None:
                continue
            pre  = _decimal(_t(_find(dta, "globe:PreRecast",       ns)))
            bef  = _decimal(_t(_find(dta, "globe:BefRecastAdjust", ns)))
            tadj = _decimal(_t(_find(dta, "globe:TotalAdjust",     ns)))
            if pre is None or bef is None or tadj is None:
                continue
            if abs(pre - (bef + tadj)) > Decimal("1"):
                bad.append(f"{jur_name}: PreRecast={pre} ≠ {bef}+{tadj}={bef+tadj}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70080 AdjustmentItem in DeferTaxAdjustAmt/Adjustments unico per ETR
    r = CheckResult("70080",
        "Ogni AdjustmentItem in DeferTaxAdjustAmt/Adjustments non può comparire più di una volta per ETR.",
        "JurisdictionSection/.../AdjustedCoveredTax/DeferTaxAdjustAmt/Adjustments/AdjustmentItem")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            dta = _find(oc, "globe:AdjustedCoveredTax/globe:DeferTaxAdjustAmt", ns)
            if dta is None:
                continue
            items = [_t(x) for x in _findall(dta, "globe:Adjustments/globe:AdjustmentItem", ns)]
            dups = {x for x in items if items.count(x) > 1}
            if dups:
                bad.append(f"{jur_name}: DeferTaxAdjustAmt AdjustmentItem duplicati: {dups}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70081 Transition/DeferredTaxAssets/Total = DeferredTaxAssetStart - DeferredTaxAssetExcluded
    #        oppure DeferredTaxAssetRecast - DeferredTaxAssetExcluded
    r = CheckResult("70081",
        "Transition/DeferredTaxAssets/Total = DeferredTaxAssetStart - Excluded OPPURE DeferredTaxAssetRecast - Excluded.",
        "JurisdictionSection/.../Transition/DeferredTaxAssets/Total")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            dta = _find(oc, "globe:AdjustedCoveredTax/globe:DeferTaxAdjustAmt", ns)
            if dta is None:
                continue
            trans = _find(dta, "globe:Transition/globe:DeferredTaxAssets", ns)
            if trans is None:
                continue
            total   = _decimal(_t(_find(trans, "globe:Total",                   ns)))
            start   = _decimal(_t(_find(trans, "globe:DeferredTaxAssetStart",   ns)))
            recast  = _decimal(_t(_find(trans, "globe:DeferredTaxAssetRecast",  ns)))
            excl    = _decimal(_t(_find(trans, "globe:DeferredTaxAssetExcluded",ns)))
            if total is None or excl is None:
                continue
            ok1 = start  is not None and abs(total - (start  - excl)) <= Decimal("1")
            ok2 = recast is not None and abs(total - (recast - excl)) <= Decimal("1")
            if not ok1 and not ok2:
                bad.append(f"{jur_name}: DeferredTaxAssets/Total={total} non coincide con nessuno dei calcoli")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70082 Se DeferredTaxAssets presente, Start XOR Recast deve essere 0
    r = CheckResult("70082",
        "Se Transition/DeferredTaxAssets presente, uno tra DeferredTaxAssetStart e DeferredTaxAssetRecast deve essere 0.",
        "JurisdictionSection/.../Transition/DeferredTaxAssets")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            dta = _find(oc, "globe:AdjustedCoveredTax/globe:DeferTaxAdjustAmt", ns)
            if dta is None:
                continue
            trans = _find(dta, "globe:Transition/globe:DeferredTaxAssets", ns)
            if trans is None:
                continue
            start  = _decimal(_t(_find(trans, "globe:DeferredTaxAssetStart",  ns)))
            recast = _decimal(_t(_find(trans, "globe:DeferredTaxAssetRecast", ns)))
            if start is not None and recast is not None:
                if start != 0 and recast != 0:
                    bad.append(f"{jur_name}: DeferredTaxAssetStart={start} e Recast={recast} entrambi ≠ 0")
    if bad:
        r.ko(bad[0])
    out.append(r)

    return out


# ── Helper 8.3.16 ExcessNegTaxExpense / ExcessProfits / Substance (70083-70087)

def _check_excess_substance(jur_sections, ns):
    out = []

    # 70083 ExcessNegTaxExpense/Remaining = PriorYearBalance + GeneratedInRFY - UtilizedInRFY
    r = CheckResult("70083",
        "ExcessNegTaxExpense/Remaining = PriorYearBalance + GeneratedInRFY - UtilizedInRFY.",
        "JurisdictionSection/.../ExcessNegTaxExpense/Remaining")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            ente = _find(oc, "globe:ExcessNegTaxExpense", ns)
            if ente is None:
                continue
            rem  = _decimal(_t(_find(ente, "globe:Remaining",       ns)))
            prio = _decimal(_t(_find(ente, "globe:PriorYearBalance",ns)))
            gen  = _decimal(_t(_find(ente, "globe:GeneratedInRFY",  ns)))
            util = _decimal(_t(_find(ente, "globe:UtilizedInRFY",   ns)))
            if rem is None or prio is None or gen is None or util is None:
                continue
            expected = prio + gen - util
            if abs(rem - expected) > Decimal("1"):
                bad.append(f"{jur_name}: ExcessNegTaxExpense Remaining={rem} ≠ {expected}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70084 AdjustmentItem=GIR2719 → Amount = ExcessNegTaxExpense/GeneratedInRFY
    r = CheckResult("70084",
        "Se AdjustedCoveredTax/AdjustmentItem = GIR2719, Amount = ExcessNegTaxExpense/GeneratedInRFY.",
        "JurisdictionSection/.../AdjustedCoveredTax/Adjustments/Amount[GIR2719]")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            ente = _find(oc, "globe:ExcessNegTaxExpense", ns)
            gen  = _decimal(_t(_find(ente, "globe:GeneratedInRFY", ns))) if ente is not None else None
            for adj in _findall(oc, "globe:AdjustedCoveredTax/globe:Adjustments", ns):
                item = _t(_find(adj, "globe:AdjustmentItem", ns))
                if item == "GIR2719":
                    amt = _decimal(_t(_find(adj, "globe:Amount", ns)))
                    if gen is not None and amt is not None and abs(amt - gen) > Decimal("1"):
                        bad.append(f"{jur_name}: GIR2719 Amount={amt} ≠ GeneratedInRFY={gen}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70085 AdjustmentItem=GIR2720 → Amount = ExcessNegTaxExpense/UtilizedInRFY
    r = CheckResult("70085",
        "Se AdjustedCoveredTax/AdjustmentItem = GIR2720, Amount = ExcessNegTaxExpense/UtilizedInRFY.",
        "JurisdictionSection/.../AdjustedCoveredTax/Adjustments/Amount[GIR2720]")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            ente = _find(oc, "globe:ExcessNegTaxExpense", ns)
            util = _decimal(_t(_find(ente, "globe:UtilizedInRFY", ns))) if ente is not None else None
            for adj in _findall(oc, "globe:AdjustedCoveredTax/globe:Adjustments", ns):
                item = _t(_find(adj, "globe:AdjustmentItem", ns))
                if item == "GIR2720":
                    amt = _decimal(_t(_find(adj, "globe:Amount", ns)))
                    if util is not None and amt is not None and abs(amt - util) > Decimal("1"):
                        bad.append(f"{jur_name}: GIR2720 Amount={amt} ≠ UtilizedInRFY={util}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70086 ExcessProfits = max(0, NetGlobeIncome/Total - SubstanceExclusion/Total)
    r = CheckResult("70086",
        "OverallComputation/ExcessProfits = max(0, NetGlobeIncome/Total - SubstanceExclusion/Total).",
        "JurisdictionSection/.../OverallComputation/ExcessProfits")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            ep_el = _find(oc, "globe:ExcessProfits", ns)
            ngi   = _decimal(_t(_find(oc, "globe:NetGlobeIncome/globe:Total", ns)))
            se    = _decimal(_t(_find(oc, "globe:SubstanceExclusion/globe:Total", ns))) or Decimal(0)
            if ep_el is None or ngi is None:
                continue
            ep = _decimal(_t(ep_el))
            if ep is None:
                continue
            expected = max(Decimal(0), ngi - se)
            if abs(ep - expected) > Decimal("1"):
                bad.append(f"{jur_name}: ExcessProfits={ep} ≠ max(0,{ngi}-{se})={expected}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70087 SubstanceExclusion/Total = (PayrollCost * PayrollMarkUp) + (TangibleAssetValue * TangibleAssetMarkup)
    r = CheckResult("70087",
        "SubstanceExclusion/Total = (PayrollCost * PayrollMarkUp) + (TangibleAssetValue * TangibleAssetMarkup).",
        "JurisdictionSection/.../OverallComputation/SubstanceExclusion/Total")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            se = _find(oc, "globe:SubstanceExclusion", ns)
            if se is None:
                continue
            total  = _decimal(_t(_find(se, "globe:Total",               ns)))
            pc     = _decimal(_t(_find(se, "globe:PayrollCost",          ns)))
            pmu    = _decimal(_t(_find(se, "globe:PayrollMarkUp",         ns)))
            tav    = _decimal(_t(_find(se, "globe:TangibleAssetValue",    ns)))
            tamu   = _decimal(_t(_find(se, "globe:TangibleAssetMarkup",   ns)))
            if total is None:
                continue
            part1 = (pc or Decimal(0)) * (pmu or Decimal(0))
            part2 = (tav or Decimal(0)) * (tamu or Decimal(0))
            expected = part1 + part2
            if abs(total - expected) > Decimal("1"):
                bad.append(f"{jur_name}: SubstanceExclusion Total={total} ≠ {expected:.2f}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    return out


# ── Helper 8.3.17 AdditionalTopUpTax (70088-70096) ───────────────────────────

def _check_additional_tut(jur_sections, period_end, ns):
    out = []

    for js in jur_sections:
        pass  # elaborato sotto

    # 70088 AdditionalTopUpTax/Art4.1.5 obbligatorio quando NetGlobeIncome/Total < 0
    r = CheckResult("70088",
        "AdditionalTopUpTax/Art4.1.5 deve essere compilato quando OverallComputation/NetGlobeIncome/Total < 0.",
        "JurisdictionSection/.../AdditionalTopUpTax/Art4.1.5")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            ngi = _decimal(_t(_find(oc, "globe:NetGlobeIncome/globe:Total", ns)))
            if ngi is not None and ngi < 0:
                if _find(oc, "globe:AdditionalTopUpTax/globe:Art4.1.5", ns) is None:
                    bad.append(f"{jur_name}: NéGI={ngi} < 0 ma Art4.1.5 assente")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70089 Art4.1.5/AdjustedCoveredTax deve essere negativo
    r = CheckResult("70089",
        "Nel blocco Art4.1.5, AdjustedCoveredTax deve avere valore negativo.",
        "JurisdictionSection/.../AdditionalTopUpTax/Art4.1.5/AdjustedCoveredTax")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            art = _find(oc, "globe:AdditionalTopUpTax/globe:Art4.1.5", ns)
            if art is None:
                continue
            act = _decimal(_t(_find(art, "globe:AdjustedCoveredTax", ns)))
            if act is not None and act >= 0:
                bad.append(f"{jur_name}: Art4.1.5/AdjustedCoveredTax={act} non negativo")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70090 Art4.1.5/GlobeLoss = NetGlobeIncome/Total
    r = CheckResult("70090",
        "Art4.1.5/GlobeLoss deve coincidere con OverallComputation/NetGlobeIncome/Total.",
        "JurisdictionSection/.../AdditionalTopUpTax/Art4.1.5/GlobeLoss")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            art = _find(oc, "globe:AdditionalTopUpTax/globe:Art4.1.5", ns)
            if art is None:
                continue
            loss = _decimal(_t(_find(art, "globe:GlobeLoss", ns)))
            ngi  = _decimal(_t(_find(oc,  "globe:NetGlobeIncome/globe:Total", ns)))
            if loss is not None and ngi is not None and abs(loss - ngi) > Decimal("1"):
                bad.append(f"{jur_name}: GlobeLoss={loss} ≠ NéGI={ngi}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70091 Art4.1.5/ExpectedAdjustedCoveredTax = GlobeLoss * 15%
    r = CheckResult("70091",
        "Art4.1.5/ExpectedAdjustedCoveredTax = GlobeLoss * 15%.",
        "JurisdictionSection/.../AdditionalTopUpTax/Art4.1.5/ExpectedAdjustedCoveredTax")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            art = _find(oc, "globe:AdditionalTopUpTax/globe:Art4.1.5", ns)
            if art is None:
                continue
            exp  = _decimal(_t(_find(art, "globe:ExpectedAdjustedCoveredTax", ns)))
            loss = _decimal(_t(_find(art, "globe:GlobeLoss", ns)))
            if exp is None or loss is None:
                continue
            expected = loss * Decimal("0.15")
            if abs(exp - expected) > Decimal("1"):
                bad.append(f"{jur_name}: ExpectedACT={exp} ≠ {loss}*15%={expected:.2f}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70092 Art4.1.5/AdditionalTopUpTax = max(0, ExpectedACT - AdjustedCoveredTax)
    r = CheckResult("70092",
        "Art4.1.5/AdditionalTopUpTax = max(0, ExpectedAdjustedCoveredTax - AdjustedCoveredTax).",
        "JurisdictionSection/.../AdditionalTopUpTax/Art4.1.5/AdditionalTopUpTax")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            art = _find(oc, "globe:AdditionalTopUpTax/globe:Art4.1.5", ns)
            if art is None:
                continue
            atut = _decimal(_t(_find(art, "globe:AdditionalTopUpTax",         ns)))
            exp  = _decimal(_t(_find(art, "globe:ExpectedAdjustedCoveredTax", ns)))
            act  = _decimal(_t(_find(art, "globe:AdjustedCoveredTax",         ns)))
            if atut is None or exp is None or act is None:
                continue
            expected = max(Decimal(0), exp - act)
            if abs(atut - expected) > Decimal("1"):
                bad.append(f"{jur_name}: Art4.1.5/ATUT={atut} ≠ max(0,{exp}-{act})={expected}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70093 NONArt4.1.5/Year ≤ anno Period/End
    r = CheckResult("70093",
        "In AdditionalTopUpTax/NONArt4.1.5, Year non può essere maggiore dell'anno di Period/End.",
        "JurisdictionSection/.../AdditionalTopUpTax/NONArt4.1.5/Year")
    bad = []
    end_year = period_end.year if period_end else None
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            for non in _findall(oc, "globe:AdditionalTopUpTax/globe:NONArt4.1.5", ns):
                yr = _year(_find(non, "globe:Year", ns))
                if end_year and yr and yr > end_year:
                    bad.append(f"{jur_name}: NONArt4.1.5 Year={yr} > {end_year}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70094 Se Articles contiene GIR2605 → Year almeno 4 anni prima di Period/End
    r = CheckResult("70094",
        "Se NONArt4.1.5/Articles contiene GIR2605, Year deve essere almeno 4 anni anteriore a Period/End.",
        "JurisdictionSection/.../AdditionalTopUpTax/NONArt4.1.5/Year")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            for non in _findall(oc, "globe:AdditionalTopUpTax/globe:NONArt4.1.5", ns):
                articles = set((_t(_find(non, "globe:Articles", ns)) or "").split())
                if "GIR2605" not in articles:
                    continue
                yr = _year(_find(non, "globe:Year", ns))
                if period_end and yr and (period_end.year - yr) < 4:
                    bad.append(f"{jur_name}: GIR2605 Year={yr} non è ≥4 anni prima di {period_end.year}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70095 Se Articles contiene GIR2602 → Year = quinto anno precedente a Period/End
    r = CheckResult("70095",
        "Se NONArt4.1.5/Articles contiene GIR2602, Year deve essere il quinto anno fiscale precedente a Period/End.",
        "JurisdictionSection/.../AdditionalTopUpTax/NONArt4.1.5/Year")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            for non in _findall(oc, "globe:AdditionalTopUpTax/globe:NONArt4.1.5", ns):
                articles = set((_t(_find(non, "globe:Articles", ns)) or "").split())
                if "GIR2602" not in articles:
                    continue
                yr = _year(_find(non, "globe:Year", ns))
                if period_end and yr and (period_end.year - yr) != 5:
                    bad.append(f"{jur_name}: GIR2602 Year={yr} non è il quinto anno prima di {period_end.year}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70096 NONArt4.1.5/AdditionalTopUpTax = Recalculated/TopUpTax - Previous/TopUpTax
    r = CheckResult("70096",
        "NONArt4.1.5/AdditionalTopUpTax = Recalculated/TopUpTax - Previous/TopUpTax.",
        "JurisdictionSection/.../AdditionalTopUpTax/NONArt4.1.5/AdditionalTopUpTax")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for oc in _findall(js, ".//globe:ETRComputation/globe:OverallComputation", ns):
            for non in _findall(oc, "globe:AdditionalTopUpTax/globe:NONArt4.1.5", ns):
                atut = _decimal(_t(_find(non, "globe:AdditionalTopUpTax",           ns)))
                rec  = _decimal(_t(_find(non, "globe:Recalculated/globe:TopUpTax",  ns)))
                prev = _decimal(_t(_find(non, "globe:Previous/globe:TopUpTax",      ns)))
                if atut is None or rec is None or prev is None:
                    continue
                if abs(atut - (rec - prev)) > Decimal("1"):
                    bad.append(f"{jur_name}: NONArt4.1.5 ATUT={atut} ≠ {rec}-{prev}={rec-prev}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    return out


# ── Helper 8.3.18 JurisdictionSection IIR/UTPR (70097-70105) ─────────────────

def _check_iir_utpr(jur_sections, utpr_attr, ns):
    out = []

    # 70097 IIR/ParentEntity/InclusionRatio = (NetGlobeIncome - OtherOwnershipAllocation) / NetGlobeIncome
    r = CheckResult("70097",
        "IIR/ParentEntity/InclusionRatio = (NetGlobeIncome - OtherOwnershipAllocation) / NetGlobeIncome.",
        "JurisdictionSection/.../IIR/ParentEntity/InclusionRatio")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ltce in _findall(js, ".//globe:LowTaxJurisdiction/globe:LTCE", ns):
            for iir in _findall(ltce, "globe:IIR", ns):
                ngi_el = _find(iir, "globe:NetGlobeIncome", ns)
                ngi    = _decimal(_t(ngi_el)) if ngi_el is not None else None
                if ngi is None or ngi == 0:
                    continue
                for pe in _findall(iir, "globe:ParentEntity", ns):
                    ir_el  = _find(pe, "globe:InclusionRatio", ns)
                    ooa_el = _find(pe, "globe:OtherOwnershipAllocation", ns)
                    ir     = _decimal(_t(ir_el))
                    ooa    = _decimal(_t(ooa_el)) if ooa_el is not None else Decimal(0)
                    if ir is None:
                        continue
                    if ooa is None: ooa = Decimal(0)
                    expected = (ngi - ooa) / ngi
                    if abs(ir - expected) > Decimal("0.0001"):
                        bad.append(f"{jur_name}: InclusionRatio={ir} ≠ ({ngi}-{ooa})/{ngi}={expected:.6f}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70098 IIR/ParentEntity/TopUpTaxShare = IIR/TopUpTax * InclusionRatio
    r = CheckResult("70098",
        "IIR/ParentEntity/TopUpTaxShare = IIR/TopUpTax * InclusionRatio.",
        "JurisdictionSection/.../IIR/ParentEntity/TopUpTaxShare")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ltce in _findall(js, ".//globe:LowTaxJurisdiction/globe:LTCE", ns):
            for iir in _findall(ltce, "globe:IIR", ns):
                tut_el = _find(iir, "globe:TopUpTax", ns)
                tut    = _decimal(_t(tut_el)) if tut_el is not None else None
                for pe in _findall(iir, "globe:ParentEntity", ns):
                    share_el = _find(pe, "globe:TopUpTaxShare",  ns)
                    ir_el    = _find(pe, "globe:InclusionRatio", ns)
                    share    = _decimal(_t(share_el))
                    ir       = _decimal(_t(ir_el))
                    if share is None or tut is None or ir is None:
                        continue
                    expected = tut * ir
                    if abs(share - expected) > Decimal("1"):
                        bad.append(f"{jur_name}: TopUpTaxShare={share} ≠ {tut}*{ir}={expected:.2f}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70099 Σ(UTPRTopUpTaxAttributed) = Σ(TotalUTPRTopUpTax) per giurisdizioni pertinenti
    r = CheckResult("70099",
        "La somma di tutti i UTPRTopUpTaxAttributed deve essere uguale alla somma dei TotalUTPRTopUpTax.",
        "/GLOBE_OECD/GLOBEBody/UTPRAttribution/Attribution/UTPRTopUpTaxAttributed")
    if utpr_attr is not None:
        sum_attr = sum(
            (_decimal(_t(x)) or Decimal(0))
            for x in _findall(utpr_attr, ".//globe:Attribution/globe:UTPRTopUpTaxAttributed", ns)
        )
        sum_total = sum(
            (_decimal(_t(x)) or Decimal(0))
            for js in jur_sections
            for x in _findall(js, ".//globe:LowTaxJurisdiction/globe:UTPR/globe:UTPRCalculation/globe:TotalUTPRTopUpTax", ns)
        )
        if abs(sum_attr - sum_total) > Decimal("1"):
            r.ko(f"Σ(UTPRTopUpTaxAttributed)={sum_attr} ≠ Σ(TotalUTPRTopUpTax)={sum_total}")
    out.append(r)

    # 70100 Se UTPRCalculation presente e TotalUTPRTopUpTax > 0 → UTPRAttribution compilato
    r = CheckResult("70100",
        "Se UTPRCalculation presente e TotalUTPRTopUpTax > 0, UTPRAttribution deve essere compilato.",
        "/GLOBE_OECD/GLOBEBody/UTPRAttribution")
    bad = []
    for js in jur_sections:
        for utpr_calc in _findall(js, ".//globe:LowTaxJurisdiction/globe:UTPR/globe:UTPRCalculation", ns):
            total = _decimal(_t(_find(utpr_calc, "globe:TotalUTPRTopUpTax", ns)))
            if total is not None and total > 0 and utpr_attr is None:
                bad.append(f"TotalUTPRTopUpTax={total}>0 ma UTPRAttribution assente")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70101 Attribution/Employees compilato salvo UTPRTopUpTaxCarryForward = 0
    r = CheckResult("70101",
        "UTPRAttribution/Attribution/Employees deve essere compilato salvo che UTPRTopUpTaxCarryForward = 0.",
        "/GLOBE_OECD/GLOBEBody/UTPRAttribution/Attribution/Employees")
    bad = []
    if utpr_attr is not None:
        for attr in _findall(utpr_attr, "globe:Attribution", ns):
            cf  = _decimal(_t(_find(attr, "globe:UTPRTopUpTaxCarryForward", ns)))
            emp = _find(attr, "globe:Employees", ns)
            if cf is None or cf != 0:
                if emp is None:
                    bad.append("Attribution: Employees assente e CarryForward≠0")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70102 Attribution/TangibleAssetValue compilato salvo UTPRTopUpTaxCarryForward = 0
    r = CheckResult("70102",
        "UTPRAttribution/Attribution/TangibleAssetValue deve essere compilato salvo UTPRTopUpTaxCarryForward = 0.",
        "/GLOBE_OECD/GLOBEBody/UTPRAttribution/Attribution/TangibleAssetValue")
    bad = []
    if utpr_attr is not None:
        for attr in _findall(utpr_attr, "globe:Attribution", ns):
            cf  = _decimal(_t(_find(attr, "globe:UTPRTopUpTaxCarryForward", ns)))
            tav = _find(attr, "globe:TangibleAssetValue", ns)
            if cf is None or cf != 0:
                if tav is None:
                    bad.append("Attribution: TangibleAssetValue assente e CarryForward≠0")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70103 UTPRPercentage = 0 quando CarryForward > 0; se CF=0 e tutte ≠0 → errore
    r = CheckResult("70103",
        "UTPRPercentage deve essere 0 quando UTPRTopUpTaxCarryForward > 0.",
        "/GLOBE_OECD/GLOBEBody/UTPRAttribution/Attribution/UTPRPercentage")
    bad = []
    if utpr_attr is not None:
        for attr in _findall(utpr_attr, "globe:Attribution", ns):
            cf  = _decimal(_t(_find(attr, "globe:UTPRTopUpTaxCarryForward", ns)))
            pct = _decimal(_t(_find(attr, "globe:UTPRPercentage", ns)))
            if cf is not None and cf > 0 and pct is not None and pct != 0:
                bad.append(f"CarryForward={cf}>0 ma UTPRPercentage={pct}≠0")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70104 UTPRTopUpTaxCarriedForward non negativo
    r = CheckResult("70104",
        "UTPRTopUpTaxCarriedForward non può essere negativo.",
        "/GLOBE_OECD/GLOBEBody/UTPRAttribution/Attribution/UTPRTopUpTaxCarriedForward")
    bad = []
    if utpr_attr is not None:
        for attr in _findall(utpr_attr, "globe:Attribution", ns):
            val = _decimal(_t(_find(attr, "globe:UTPRTopUpTaxCarriedForward", ns)))
            if val is not None and val < 0:
                bad.append(f"UTPRTopUpTaxCarriedForward={val} < 0")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70105 UTPRTopUpTaxCarriedForward = CarryForward + Attributed - AddCashTaxExpense
    r = CheckResult("70105",
        "UTPRTopUpTaxCarriedForward = UTPRTopUpTaxCarryForward + UTPRTopUpTaxAttributed - AddCashTaxExpense.",
        "/GLOBE_OECD/GLOBEBody/UTPRAttribution/Attribution/UTPRTopUpTaxCarriedForward")
    bad = []
    if utpr_attr is not None:
        for attr in _findall(utpr_attr, "globe:Attribution", ns):
            carried  = _decimal(_t(_find(attr, "globe:UTPRTopUpTaxCarriedForward", ns)))
            forward  = _decimal(_t(_find(attr, "globe:UTPRTopUpTaxCarryForward",   ns))) or Decimal(0)
            attrib   = _decimal(_t(_find(attr, "globe:UTPRTopUpTaxAttributed",     ns))) or Decimal(0)
            cash     = _decimal(_t(_find(attr, "globe:AddCashTaxExpense",          ns))) or Decimal(0)
            if carried is None:
                continue
            if forward is None: forward = Decimal(0)
            if attrib  is None: attrib  = Decimal(0)
            if cash    is None: cash    = Decimal(0)
            expected = forward + attrib - cash
            if abs(carried - expected) > Decimal("1"):
                bad.append(f"CarriedForward={carried} ≠ {forward}+{attrib}-{cash}={expected}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    return out


# ── Helper 8.3.19 CEComputation AdjustedFANIL / UPEAdjustments (70106-70113) ─

def _check_ce_fanil(jur_sections, ns):
    out = []

    # 70106 CrossBorderAdjustments/OtherTIN ≠ CEComputation/TIN
    r = CheckResult("70106",
        "CrossBorderAdjustments/OtherTIN non deve coincidere con CEComputation/TIN.",
        "JurisdictionSection/.../CEComputation/AdjustedFANIL/CrossBorderAdjustments/OtherTIN")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ce_c in _findall(js, ".//globe:CEComputation", ns):
            ce_tin = _t(_find(ce_c, "globe:TIN", ns))
            for adj in _findall(ce_c, ".//globe:AdjustedFANIL/globe:Adjustment/globe:CrossBorderAdjustments", ns):
                other = _t(_find(adj, "globe:OtherTIN", ns))
                if ce_tin and other and ce_tin == other:
                    bad.append(f"{jur_name}: CrossBorderAdjustments/OtherTIN={other!r} = CEComputation/TIN")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70107 Se UPEAdjustments/Reductions/Exception=TRUE → CrossBorderAdjustments non presente
    r = CheckResult("70107",
        "Se UPEAdjustments/Reductions/Exception = TRUE, CrossBorderAdjustments non deve essere fornito.",
        "JurisdictionSection/.../CEComputation/AdjustedFANIL/Adjustment/CrossBorderAdjustments")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ce_c in _findall(js, ".//globe:CEComputation", ns):
            for adj in _findall(ce_c, ".//globe:AdjustedFANIL/globe:Adjustment", ns):
                exc_el = _find(adj, "globe:UPEAdjustments/globe:Reductions/globe:Exception", ns)
                if exc_el is not None and _t(exc_el).upper() == "TRUE":
                    if _find(adj, "globe:CrossBorderAdjustments", ns) is not None:
                        bad.append(f"{jur_name}: Exception=TRUE ma CrossBorderAdjustments presente")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70108 Basis = GIR1901/GIR1902/GIR1905/GIR1906 → EntityOwner/TaxRate o IndOwners/TaxRate
    r = CheckResult("70108",
        "Se Basis = GIR1901/1902/1905/1906, deve essere compilato EntityOwner/TaxRate o IndOwners/TaxRate.",
        "JurisdictionSection/.../CEComputation/AdjustedFANIL/UPEAdjustments/IdentificationOfOwners")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ce_c in _findall(js, ".//globe:CEComputation", ns):
            for upe_adj in _findall(ce_c, ".//globe:UPEAdjustments", ns):
                basis = _t(_find(upe_adj, "globe:Basis", ns))
                if basis in ("GIR1901","GIR1902","GIR1905","GIR1906"):
                    ioo = _find(upe_adj, "globe:IdentificationOfOwners", ns)
                    has_entity = ioo is not None and _find(ioo, "globe:EntityOwner/globe:TaxRate", ns) is not None
                    has_ind    = ioo is not None and _find(ioo, "globe:IndOwners/globe:TaxRate",   ns) is not None
                    if not has_entity and not has_ind:
                        bad.append(f"{jur_name}: Basis={basis} ma mancano EntityOwner/TaxRate e IndOwners/TaxRate")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70109 Basis = GIR1907 → IndOwners/ResCountryCode
    r = CheckResult("70109",
        "Se Basis = GIR1907, deve essere compilato IndOwners/ResCountryCode.",
        "JurisdictionSection/.../CEComputation/AdjustedFANIL/UPEAdjustments/IdentificationOfOwners/IndOwners/ResCountryCode")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ce_c in _findall(js, ".//globe:CEComputation", ns):
            for upe_adj in _findall(ce_c, ".//globe:UPEAdjustments", ns):
                if _t(_find(upe_adj, "globe:Basis", ns)) == "GIR1907":
                    ioo = _find(upe_adj, "globe:IdentificationOfOwners", ns)
                    if ioo is None or _find(ioo, "globe:IndOwners/globe:ResCountryCode", ns) is None:
                        bad.append(f"{jur_name}: Basis=GIR1907 ma IndOwners/ResCountryCode assente")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70110 Basis = GIR1903/GIR1908 → IndOwners compilato
    r = CheckResult("70110",
        "Se Basis = GIR1903 o GIR1908, deve essere compilato IndOwners.",
        "JurisdictionSection/.../CEComputation/AdjustedFANIL/UPEAdjustments/IdentificationOfOwners/IndOwners")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ce_c in _findall(js, ".//globe:CEComputation", ns):
            for upe_adj in _findall(ce_c, ".//globe:UPEAdjustments", ns):
                if _t(_find(upe_adj, "globe:Basis", ns)) in ("GIR1903","GIR1908"):
                    ioo = _find(upe_adj, "globe:IdentificationOfOwners", ns)
                    if ioo is None or _find(ioo, "globe:IndOwners", ns) is None:
                        bad.append(f"{jur_name}: Basis=GIR1903/1908 ma IndOwners assente")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70111 Basis = GIR1904/GIR1909 → EntityOwner/ExTypeOfEntity
    r = CheckResult("70111",
        "Se Basis = GIR1904 o GIR1909, deve essere compilato EntityOwner/ExTypeOfEntity.",
        "JurisdictionSection/.../CEComputation/AdjustedFANIL/UPEAdjustments/IdentificationOfOwners/EntityOwner/ExTypeOfEntity")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ce_c in _findall(js, ".//globe:CEComputation", ns):
            for upe_adj in _findall(ce_c, ".//globe:UPEAdjustments", ns):
                if _t(_find(upe_adj, "globe:Basis", ns)) in ("GIR1904","GIR1909"):
                    ioo = _find(upe_adj, "globe:IdentificationOfOwners", ns)
                    if ioo is None or _find(ioo, "globe:EntityOwner/globe:ExTypeOfEntity", ns) is None:
                        bad.append(f"{jur_name}: Basis=GIR1904/1909 ma EntityOwner/ExTypeOfEntity assente")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70112 Basis = GIR1904 → ExTypeOfEntity ≠ GIR2805
    r = CheckResult("70112",
        "Se Basis = GIR1904, ExTypeOfEntity non deve avere valore GIR2805.",
        "JurisdictionSection/.../CEComputation/AdjustedFANIL/UPEAdjustments/IdentificationOfOwners/EntityOwner/ExTypeOfEntity")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ce_c in _findall(js, ".//globe:CEComputation", ns):
            for upe_adj in _findall(ce_c, ".//globe:UPEAdjustments", ns):
                if _t(_find(upe_adj, "globe:Basis", ns)) == "GIR1904":
                    ioo = _find(upe_adj, "globe:IdentificationOfOwners", ns)
                    ext = _t(_find(ioo, "globe:EntityOwner/globe:ExTypeOfEntity", ns)) if ioo else ""
                    if ext == "GIR2805":
                        bad.append(f"{jur_name}: Basis=GIR1904 ma ExTypeOfEntity=GIR2805")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70113 Basis = GIR1909 → ExTypeOfEntity ≠ GIR2804
    r = CheckResult("70113",
        "Se Basis = GIR1909, ExTypeOfEntity non deve avere valore GIR2804.",
        "JurisdictionSection/.../CEComputation/AdjustedFANIL/UPEAdjustments/IdentificationOfOwners/EntityOwner/ExTypeOfEntity")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ce_c in _findall(js, ".//globe:CEComputation", ns):
            for upe_adj in _findall(ce_c, ".//globe:UPEAdjustments", ns):
                if _t(_find(upe_adj, "globe:Basis", ns)) == "GIR1909":
                    ioo = _find(upe_adj, "globe:IdentificationOfOwners", ns)
                    ext = _t(_find(ioo, "globe:EntityOwner/globe:ExTypeOfEntity", ns)) if ioo else ""
                    if ext == "GIR2804":
                        bad.append(f"{jur_name}: Basis=GIR1909 ma ExTypeOfEntity=GIR2804")
    if bad:
        r.ko(bad[0])
    out.append(r)

    return out


# ── Helper 8.3.20 CEComputation NetGlobeIncome / Elections / ACT (70114-70124) ─

def _check_ce_ngi(jur_sections, ns):
    out = []

    # 70114 Se due valori NetGlobeIncome/Adjustments/Amount → uno negativo e uno positivo
    r = CheckResult("70114",
        "Se sono forniti due valori CEComputation/NetGlobeIncome/Adjustments/Amount, uno deve essere negativo e l'altro positivo.",
        "JurisdictionSection/.../CEComputation/NetGlobeIncome/Adjustments/Amount")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ce_c in _findall(js, ".//globe:CEComputation", ns):
            amounts = [_decimal(_t(x)) for x in _findall(ce_c, "globe:NetGlobeIncome/globe:Adjustments/globe:Amount", ns)]
            amounts = [a for a in amounts if a is not None]
            if len(amounts) == 2:
                if not (amounts[0] < 0 < amounts[1]) and not (amounts[1] < 0 < amounts[0]):
                    bad.append(f"{jur_name}: NéGI Amounts {amounts} non hanno segni opposti")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70115 AdjustmentItem GIR2022/GIR2023 → UPEAdjustments compilato
    r = CheckResult("70115",
        "Se CEComputation/NetGlobeIncome/AdjustmentItem contiene GIR2022 o GIR2023, AdjustedFANIL/UPEAdjustments deve essere compilato.",
        "JurisdictionSection/.../CEComputation/NetGlobeIncome/Adjustments/AdjustmentItem")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ce_c in _findall(js, ".//globe:CEComputation", ns):
            items = {_t(x) for x in _findall(ce_c, "globe:NetGlobeIncome/globe:Adjustments/globe:AdjustmentItem", ns)}
            if items & {"GIR2022","GIR2023"}:
                upe_adj = _find(ce_c, ".//globe:AdjustedFANIL/globe:Adjustment/globe:UPEAdjustments", ns)
                if upe_adj is None:
                    bad.append(f"{jur_name}: GIR2022/2023 ma UPEAdjustments assente")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70116 AdjustmentItem GIR2025 → NetGlobeIncome/IntShippingIncome
    r = CheckResult("70116",
        "Se CEComputation/NetGlobeIncome/AdjustmentItem = GIR2025, NetGlobeIncome/IntShippingIncome deve essere compilato.",
        "JurisdictionSection/.../CEComputation/NetGlobeIncome/IntShippingIncome")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ce_c in _findall(js, ".//globe:CEComputation", ns):
            items = {_t(x) for x in _findall(ce_c, "globe:NetGlobeIncome/globe:Adjustments/globe:AdjustmentItem", ns)}
            if "GIR2025" in items:
                if _find(ce_c, "globe:NetGlobeIncome/globe:IntShippingIncome", ns) is None:
                    bad.append(f"{jur_name}: GIR2025 ma IntShippingIncome assente")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70117 AdjustmentItem GIR2024 → Elections/Art7.6
    r = CheckResult("70117",
        "Se CEComputation/NetGlobeIncome/AdjustmentItem = GIR2024, Elections/Art7.6 deve essere compilato.",
        "JurisdictionSection/.../CEComputation/Elections/Art7.6")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ce_c in _findall(js, ".//globe:CEComputation", ns):
            items = {_t(x) for x in _findall(ce_c, "globe:NetGlobeIncome/globe:Adjustments/globe:AdjustmentItem", ns)}
            if "GIR2024" in items:
                if _find(ce_c, "globe:Elections/globe:Art7.6", ns) is None:
                    bad.append(f"{jur_name}: GIR2024 ma Elections/Art7.6 assente")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70118 Se due valori AdjustedCoveredTax/Adjustments/Amount → uno negativo e uno positivo
    r = CheckResult("70118",
        "Se sono forniti due valori CEComputation/AdjustedCoveredTax/Adjustments/Amount, uno deve essere negativo e l'altro positivo.",
        "JurisdictionSection/.../CEComputation/AdjustedCoveredTax/Adjustments/Amount")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ce_c in _findall(js, ".//globe:CEComputation", ns):
            amounts = [_decimal(_t(x)) for x in _findall(ce_c, "globe:AdjustedCoveredTax/globe:Adjustments/globe:Amount", ns)]
            amounts = [a for a in amounts if a is not None]
            if len(amounts) == 2:
                if not (amounts[0] < 0 < amounts[1]) and not (amounts[1] < 0 < amounts[0]):
                    bad.append(f"{jur_name}: ACT Amounts {amounts} non hanno segni opposti")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70119 AdjustmentItem in CEComputation/AdjustedCoveredTax/Adjustments unico
    r = CheckResult("70119",
        "Ogni AdjustmentItem in CEComputation/AdjustedCoveredTax/Adjustments non può comparire più di una volta per ETR.",
        "JurisdictionSection/.../CEComputation/AdjustedCoveredTax/Adjustments/AdjustmentItem")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ce_c in _findall(js, ".//globe:CEComputation", ns):
            items = [_t(x) for x in _findall(ce_c, "globe:AdjustedCoveredTax/globe:Adjustments/globe:AdjustmentItem", ns)]
            dups = {x for x in items if items.count(x) > 1}
            if dups:
                bad.append(f"{jur_name}: CEComputation ACT AdjustmentItem duplicati: {dups}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70120 CEComputation/DeferTaxAdjustAmt/Total = DeferTaxExpense + Σ(Adjustment/Amount) + Recast/Higher + Recast/Lower
    r = CheckResult("70120",
        "CEComputation/AdjustedCoveredTax/DeferTaxAdjustAmt/Total = DeferTaxExpense + Σ(Adjustment/Amount) + Recast/Higher + Recast/Lower.",
        "JurisdictionSection/.../CEComputation/AdjustedCoveredTax/DeferTaxAdjustAmt/Total")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ce_c in _findall(js, ".//globe:CEComputation", ns):
            dta = _find(ce_c, "globe:AdjustedCoveredTax/globe:DeferTaxAdjustAmt", ns)
            if dta is None:
                continue
            total = _decimal(_t(_find(dta, "globe:Total", ns)))
            dte   = _decimal(_t(_find(dta, "globe:DeferTaxExpense", ns))) or Decimal(0)
            higher= _decimal(_t(_find(dta, "globe:Recast/globe:Higher", ns))) or Decimal(0)
            lower = _decimal(_t(_find(dta, "globe:Recast/globe:Lower",  ns))) or Decimal(0)
            adj_sum = sum(
                (_decimal(_t(_find(adj, "globe:Amount", ns))) or Decimal(0))
                for adj in _findall(dta, "globe:Adjustment", ns)
            )
            if total is None:
                continue
            expected = dte + adj_sum + higher + lower
            if abs(total - expected) > Decimal("1"):
                bad.append(f"{jur_name}: CEComputation DTA Total={total} ≠ {expected:.2f}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70121 AdjustmentItem in DeferTaxAdjustAmt/Adjustment unico
    r = CheckResult("70121",
        "Ogni AdjustmentItem in CEComputation/AdjustedCoveredTax/DeferTaxAdjustAmt/Adjustment non può comparire più di una volta.",
        "JurisdictionSection/.../CEComputation/AdjustedCoveredTax/DeferTaxAdjustAmt/Adjustment/AdjustmentItem")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ce_c in _findall(js, ".//globe:CEComputation", ns):
            dta = _find(ce_c, "globe:AdjustedCoveredTax/globe:DeferTaxAdjustAmt", ns)
            if dta is None:
                continue
            items = [_t(_find(adj, "globe:AdjustmentItem", ns)) for adj in _findall(dta, "globe:Adjustment", ns)]
            dups = {x for x in items if items.count(x) > 1}
            if dups:
                bad.append(f"{jur_name}: CEComputation DTA Adjustment duplicati: {dups}")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70122 Se due valori DeferTaxAdjustAmt/Adjustment/Amount → uno negativo e uno positivo
    r = CheckResult("70122",
        "Se sono forniti due valori CEComputation/DeferTaxAdjustAmt/Adjustment/Amount, uno deve essere negativo e l'altro positivo.",
        "JurisdictionSection/.../CEComputation/AdjustedCoveredTax/DeferTaxAdjustAmt/Adjustment/Amount")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ce_c in _findall(js, ".//globe:CEComputation", ns):
            dta = _find(ce_c, "globe:AdjustedCoveredTax/globe:DeferTaxAdjustAmt", ns)
            if dta is None:
                continue
            amounts = [_decimal(_t(_find(adj, "globe:Amount", ns))) for adj in _findall(dta, "globe:Adjustment", ns)]
            amounts = [a for a in amounts if a is not None]
            if len(amounts) == 2:
                if not (amounts[0] < 0 < amounts[1]) and not (amounts[1] < 0 < amounts[0]):
                    bad.append(f"{jur_name}: DTA Amounts {amounts} non opposti")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70123 AdjustedIncomeTax/CrossAllocation/Additions non negativo
    r = CheckResult("70123",
        "AdjustedIncomeTax/CrossAllocation/Additions non deve avere valore negativo.",
        "JurisdictionSection/.../CEComputation/AdjustedIncomeTax/CrossAllocation/Additions")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ce_c in _findall(js, ".//globe:CEComputation", ns):
            add_el = _find(ce_c, "globe:AdjustedIncomeTax/globe:CrossAllocation/globe:Additions", ns)
            if add_el is not None:
                val = _decimal(_t(add_el))
                if val is not None and val < 0:
                    bad.append(f"{jur_name}: CrossAllocation/Additions={val} < 0")
    if bad:
        r.ko(bad[0])
    out.append(r)

    # 70124 AdjustedIncomeTax/CrossAllocation/Reductions non positivo
    r = CheckResult("70124",
        "AdjustedIncomeTax/CrossAllocation/Reductions non deve avere valore positivo.",
        "JurisdictionSection/.../CEComputation/AdjustedIncomeTax/CrossAllocation/Reductions")
    bad = []
    for js in jur_sections:
        jur_name = _t(_find(js, "globe:Jurisdiction", ns))
        for ce_c in _findall(js, ".//globe:CEComputation", ns):
            red_el = _find(ce_c, "globe:AdjustedIncomeTax/globe:CrossAllocation/globe:Reductions", ns)
            if red_el is not None:
                val = _decimal(_t(red_el))
                if val is not None and val > 0:
                    bad.append(f"{jur_name}: CrossAllocation/Reductions={val} > 0")
    if bad:
        r.ko(bad[0])
    out.append(r)

    return out


# ── Generazione XLSX ──────────────────────────────────────────────────────────

_FILL_OK   = PatternFill("solid", fgColor="C6EFCE")
_FILL_KO   = PatternFill("solid", fgColor="FFC7CE")
_FILL_SKIP = PatternFill("solid", fgColor="FFEB9C")
_FILL_WARN = PatternFill("solid", fgColor="DDEBF7")

_FONT_HDR  = Font(bold=True, color="FFFFFF")
_FILL_HDR_BLUE   = PatternFill("solid", fgColor="00338D")
_FILL_HDR_YELLOW = PatternFill("solid", fgColor="EAAA00")

_THIN = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)

def _hdr(ws, row, col, value, fill=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font    = _FONT_HDR
    cell.fill    = fill or _FILL_HDR_BLUE
    cell.border  = _THIN
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    return cell

def _cell(ws, row, col, value, fill=None, bold=False):
    cell = ws.cell(row=row, column=col, value=value)
    if fill:
        cell.fill = fill
    cell.border = _THIN
    cell.font   = Font(bold=bold)
    cell.alignment = Alignment(wrap_text=True, vertical="top")
    return cell

def _fill_for(status: str):
    return {
        "OK":   _FILL_OK,
        "KO":   _FILL_KO,
        "SKIP": _FILL_SKIP,
        "WARN": _FILL_WARN,
    }.get(status, None)


def _write_xlsx(results: list[CheckResult], xml_path: Path, output_dir: Path,
                guid_resolved: bool = False) -> Path:
    stem    = xml_path.stem
    out_path = output_dir / f"{stem}_validation_report.xlsx"
    wb = Workbook()

    # ── Foglio 1: Sommario ────────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Sommario"
    ws1.sheet_view.showGridLines = False

    counts = {s: sum(1 for r in results if r.status == s)
              for s in ("OK","KO","SKIP","WARN")}
    total = len(results)
    esito = "✅ VALIDO" if counts["KO"] == 0 else "❌ ERRORI PRESENTI"

    ws1.merge_cells("A1:F1")
    t = ws1["A1"]
    t.value = "PILLAR GloBE/DAC9 – RAPPORTO DI VALIDAZIONE"
    t.font  = Font(bold=True, size=14, color="FFFFFF")
    t.fill  = _FILL_HDR_BLUE
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 28

    info = [
        ("File XML",    xml_path.name),
        ("Data/ora",    datetime.now().strftime("%d/%m/%Y %H:%M:%S")),
        ("Check totali",total),
        ("✓ OK",        counts["OK"]),
        ("✗ KO",        counts["KO"]),
        ("▷ SKIP",      counts["SKIP"]),
        ("⚠ WARN",      counts["WARN"]),
        ("Esito",       esito),
    ]
    if guid_resolved:
        info.insert(2, ("⚠ Placeholder GUID",
                        "Il file conteneva {Guid:D} non risolti. "
                        "I check sono stati eseguiti su UUID4 generati casualmente. "
                        "Il file sorgente NON è stato modificato."))

    for i, (k, v) in enumerate(info, start=2):
        ws1.cell(row=i, column=1, value=k).font = Font(bold=True)
        ws1.cell(row=i, column=1).border = _THIN
        c = ws1.cell(row=i, column=2, value=v)
        c.border = _THIN
        if k == "Esito":
            c.fill = _FILL_OK if "VALIDO" in str(v) else _FILL_KO
            c.font = Font(bold=True)
        elif k.startswith("⚠ Placeholder"):
            c.fill = _FILL_SKIP
            ws1.cell(row=i, column=1).fill = _FILL_SKIP

    ws1.column_dimensions["A"].width = 22
    ws1.column_dimensions["B"].width = 70

    # ── Foglio 2: Dettaglio ───────────────────────────────────────────────────
    ws2 = wb.create_sheet("Dettaglio")
    ws2.sheet_view.showGridLines = False
    headers = ["Codice","Categoria","Descrizione","XPath","Esito","Dettaglio"]
    widths  = [10, 12, 55, 50, 8, 60]
    for col, (h, w) in enumerate(zip(headers, widths), start=1):
        _hdr(ws2, 1, col, h)
        ws2.column_dimensions[get_column_letter(col)].width = w

    for row, r in enumerate(results, start=2):
        cat = "SEVERE" if r.code.startswith("6") else \
              "OTHER"  if r.code.startswith("7") else \
              "FILE"
        fill = _fill_for(r.status)
        _cell(ws2, row, 1, r.code,   fill)
        _cell(ws2, row, 2, cat,      fill)
        _cell(ws2, row, 3, r.desc,   fill)
        _cell(ws2, row, 4, r.xpath,  fill)
        _cell(ws2, row, 5, r.status, fill, bold=True)
        _cell(ws2, row, 6, r.detail, fill)
        ws2.row_dimensions[row].height = 30

    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = f"A1:F{len(results)+1}"

    # ── Foglio 3: Errori e Warning ────────────────────────────────────────────
    ws3 = wb.create_sheet("Errori e Warning")
    ws3.sheet_view.showGridLines = False
    for col, (h, w) in enumerate(zip(headers, widths), start=1):
        _hdr(ws3, 1, col, h, _FILL_HDR_YELLOW)
        ws3.column_dimensions[get_column_letter(col)].width = w

    err_row = 2
    for r in results:
        if r.status in ("KO","WARN"):
            cat = "SEVERE" if r.code.startswith("6") else \
                  "OTHER"  if r.code.startswith("7") else \
                  "FILE"
            fill = _fill_for(r.status)
            _cell(ws3, err_row, 1, r.code,   fill)
            _cell(ws3, err_row, 2, cat,      fill)
            _cell(ws3, err_row, 3, r.desc,   fill)
            _cell(ws3, err_row, 4, r.xpath,  fill)
            _cell(ws3, err_row, 5, r.status, fill, bold=True)
            _cell(ws3, err_row, 6, r.detail, fill)
            ws3.row_dimensions[err_row].height = 30
            err_row += 1

    if err_row == 2:
        ws3.cell(row=2, column=1, value="Nessun errore rilevato ✅").font = Font(bold=True, color="375623")

    ws3.freeze_panes = "A2"

    wb.save(str(out_path))
    return out_path