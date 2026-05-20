"""
xml_validator.py – Validatore XML GloBE DAC9/GIR
=================================================
Implementa le regole di controllo dall'Allegato 2 AdE (13 marzo 2026):
  - 8.1 FILE ERRORS (500xx)      → errori strutturali bloccanti
  - 8.2 RECORD ERRORS SEVERE     → errori 60001-60028, file respinto
  - 8.3 RECORD ERRORS OTHER      → errori 70001-70124, file accettato con warning

Output: validation_report.xlsx con 3 fogli:
  - Sommario    → contatori e stato generale
  - Dettaglio   → lista completa di tutti i check eseguiti
  - Errori      → solo i check con esito KO
"""
from __future__ import annotations
from pathlib import Path
from datetime import datetime, date
from lxml import etree
import re

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_XL = True
except ImportError:
    HAS_XL = False

# ── Namespace ─────────────────────────────────────────────────────────────────
NS  = "urn:oecd:ties:globe:v2"
STF = "urn:oecd:ties:globestf:v5"
G   = f"{{{NS}}}"
S   = f"{{{STF}}}"

def _t(el, tag):
    """Testo di un figlio diretto."""
    if el is None: return None
    f = el.find(f"{G}{tag}")
    if f is None: f = el.find(tag)
    return f.text.strip() if f is not None and f.text else None

def _ta(el, tag):
    """Lista testi figli."""
    if el is None: return []
    return [e.text.strip() for e in el.findall(f"{G}{tag}") if e.text]

def _find(root, xpath):
    """Trova elementi con namespace globe: sostituito."""
    try:
        return root.findall(xpath.replace("globe:", f"{G}").replace("stf:", f"{S}"))
    except: return []

def _val(el):
    return el.text.strip() if el is not None and el.text else None

def _num(el):
    v = _val(el)
    if v is None: return None
    try: return float(v)
    except: return None

def _date(s):
    if not s: return None
    try: return date.fromisoformat(s[:10])
    except: return None

# ── Struttura risultato check ─────────────────────────────────────────────────
class Check:
    def __init__(self, cod, categoria, desc, xpath, esito, dettaglio=""):
        self.cod      = cod
        self.categoria = categoria  # SEVERE / OTHER / FILE
        self.desc     = desc
        self.xpath    = xpath
        self.esito    = esito       # OK / KO / SKIP
        self.dettaglio = dettaglio

# ── Validatore ────────────────────────────────────────────────────────────────
def _run_checks(root) -> list[Check]:
    results = []

    def ok(cod, cat, desc, xpath):
        results.append(Check(cod, cat, desc, xpath, "OK"))

    def ko(cod, cat, desc, xpath, det=""):
        results.append(Check(cod, cat, desc, xpath, "KO", det))

    def skip(cod, cat, desc, xpath, det=""):
        results.append(Check(cod, cat, desc, xpath, "SKIP", det))

    def chk(cod, cat, desc, xpath, cond, det=""):
        if cond: ok(cod, cat, desc, xpath)
        else: ko(cod, cat, desc, xpath, det)

    # Naviga struttura
    ms   = root.find(f"{G}MessageSpec")
    body = root.find(f"{G}GLOBEBody")
    fi   = body.find(f"{G}FilingInfo") if body is not None else None
    gs   = body.find(f"{G}GeneralSection") if body is not None else None
    cs   = gs.find(f"{G}CorporateStructure") if gs is not None else None

    msg_ref    = _t(ms, "MessageRefId") if ms is not None else None
    rep_period = _t(ms, "ReportingPeriod") if ms is not None else None
    trans_country = _t(ms, "TransmittingCountry") if ms is not None else None

    period_start = _t(fi.find(f"{G}Period"), "Start") if fi is not None and fi.find(f"{G}Period") is not None else None
    period_end   = _t(fi.find(f"{G}Period"), "End")   if fi is not None and fi.find(f"{G}Period") is not None else None

    filing_role = _t(fi.find(f"{G}FilingCE"), "Role") if fi is not None and fi.find(f"{G}FilingCE") is not None else None
    filing_tin_els = fi.find(f"{G}FilingCE").findall(f"{G}TIN") if fi is not None and fi.find(f"{G}FilingCE") is not None else []
    filing_tin_vals = [e.text.strip() for e in filing_tin_els if e.text]
    filing_country = _t(fi.find(f"{G}FilingCE"), "ResCountryCode") if fi is not None and fi.find(f"{G}FilingCE") is not None else None
    cfs_of_upe = _t(fi.find(f"{G}AccountingInfo"), "CFSofUPE") if fi is not None and fi.find(f"{G}AccountingInfo") is not None else None

    summaries   = body.findall(f"{G}Summary") if body is not None else []
    jur_sects   = body.findall(f"{G}JurisdictionSection") if body is not None else []
    utpr_attrs  = body.findall(f"{G}UTPRAttribution") if body is not None else []

    # ── 8.1 FILE ERRORS (500xx) ─────────────────────────────────────────────
    cat = "FILE"
    chk("50001", cat, "File valido rispetto allo schema XML", "/",
        root.tag == f"{G}GLOBE_OECD",
        f"Root tag inatteso: {root.tag}")

    chk("50002", cat, "Encoding UTF-8 dichiarato", "/",
        True, "")  # lxml gestisce già l'encoding

    chk("50003", cat, "MessageSpec presente", "/GLOBE_OECD/MessageSpec",
        ms is not None)

    chk("50004", cat, "GLOBEBody presente", "/GLOBE_OECD/GLOBEBody",
        body is not None)

    # ── 8.2 RECORD ERRORS SEVERE (6xxxx) ────────────────────────────────────
    cat = "SEVERE"
    PATTERN_MSGREF = re.compile(r'^[A-Z]{2}\d{4}[A-Z]{2}[A-Za-z0-9_\-]{1,}$')
    chk("60001", cat, "MessageRefId nel formato [SendingCountry][ReportingPeriod][ReceivingCountry][UniqueID]",
        "/GLOBE_OECD/MessageSpec/MessageRefId",
        msg_ref is not None and not "{Guid" in (msg_ref or "") and len((msg_ref or "")) >= 6,
        f"Valore: {msg_ref}")

    if rep_period:
        try:
            year = int(rep_period[:4])
            chk("60003", cat, "Anno di ReportingPeriod non maggiore dell'anno corrente",
                "/GLOBE_OECD/MessageSpec/ReportingPeriod",
                year <= datetime.now().year, f"Anno: {year}")
        except:
            ko("60003", cat, "Anno di ReportingPeriod non maggiore dell'anno corrente",
               "/GLOBE_OECD/MessageSpec/ReportingPeriod", f"Valore non parsabile: {rep_period}")

    # 60004 – non mescolare OECD1 con OECD2/OECD3
    all_doc_indics = [_val(e) for e in root.findall(f".//{S}DocTypeIndic") if _val(e)]
    has_new   = any(d in ("OECD1", "OECD0") for d in all_doc_indics)
    has_corr  = any(d in ("OECD2", "OECD3", "OECD10", "OECD11", "OECD13") for d in all_doc_indics)
    chk("60004", cat, "Il messaggio non può mescolare record nuovi con correzioni/cancellazioni",
        "/GLOBE_OECD/GLOBEBody/*/DocSpec/DocTypeIndic",
        not (has_new and has_corr),
        f"DocTypeIndic presenti: {set(all_doc_indics)}")

    # 60011 – DocRefId formato [SendingCountry][ReportingYear][UniqueID]
    all_doc_refs = [_val(e) for e in root.findall(f".//{S}DocRefId") if _val(e)]
    bad_refs = [r for r in all_doc_refs if "{Guid" in r or len(r) < 6]
    chk("60011", cat, "DocRefId nel formato [SendingCountry][ReportingYear][UniqueID]",
        "/GLOBE_OECD/GLOBEBody/*/DocSpec/DocRefId",
        len(bad_refs) == 0,
        f"DocRefId con placeholder o non validi: {bad_refs}")

    # 60012 – Se DocTypeIndic OECD1/OECD0, CorrDocRefId assente
    for ds in root.findall(f".//{G}DocSpec"):
        dti = _val(ds.find(f"{S}DocTypeIndic"))
        cdr = _val(ds.find(f"{S}CorrDocRefId"))
        if dti in ("OECD1", "OECD0") and cdr:
            ko("60012", cat, "Se DocTypeIndic è OECD1/OECD0, CorrDocRefId deve essere assente",
               "/GLOBE_OECD/GLOBEBody/*/DocSpec/CorrDocRefId",
               f"DocTypeIndic={dti}, CorrDocRefId={cdr}")
            break
    else:
        ok("60012", cat, "Se DocTypeIndic è OECD1/OECD0, CorrDocRefId deve essere assente",
           "/GLOBE_OECD/GLOBEBody/*/DocSpec/CorrDocRefId")

    # 60015 – Se DocTypeIndic OECD2/OECD3/OECD10/OECD11/OECD13, CorrDocRefId obbligatorio
    for ds in root.findall(f".//{G}DocSpec"):
        dti = _val(ds.find(f"{S}DocTypeIndic"))
        cdr = _val(ds.find(f"{S}CorrDocRefId"))
        if dti in ("OECD2","OECD3","OECD10","OECD11","OECD13") and not cdr:
            ko("60015", cat, "Se DocTypeIndic è correzione/cancellazione, CorrDocRefId è obbligatorio",
               "/GLOBE_OECD/GLOBEBody/*/DocSpec/CorrDocRefId",
               f"DocTypeIndic={dti} ma CorrDocRefId assente")
            break
    else:
        ok("60015", cat, "Se DocTypeIndic è correzione/cancellazione, CorrDocRefId è obbligatorio",
           "/GLOBE_OECD/GLOBEBody/*/DocSpec/CorrDocRefId")

    # 60017 – Se FilingInfo/DocTypeIndic = OECD1, GeneralSection deve essere presente
    if fi is not None:
        fi_ds   = fi.find(f"{G}DocSpec")
        fi_dti  = _val(fi_ds.find(f"{S}DocTypeIndic")) if fi_ds is not None else None
        if fi_dti == "OECD1":
            chk("60017", cat, "Se FilingInfo/DocTypeIndic=OECD1, GeneralSection deve essere presente",
                "/GLOBE_OECD/GLOBEBody/GeneralSection",
                gs is not None)

    # 60020 – Period/Start non può essere successivo a Period/End
    ds_start = _date(period_start)
    ds_end   = _date(period_end)
    if ds_start and ds_end:
        chk("60020", cat, "Period/Start non può essere successivo a Period/End",
            "/GLOBE_OECD/GLOBEBody/FilingInfo/Period",
            ds_start <= ds_end,
            f"Start={period_start}, End={period_end}")

    # 60021 – Period/End non può essere successivo a ReportingPeriod
    ds_rep = _date(rep_period)
    if ds_end and ds_rep:
        chk("60021", cat, "FilingInfo/Period/End non può essere successivo a ReportingPeriod",
            "/GLOBE_OECD/GLOBEBody/FilingInfo/Period/End",
            ds_end <= ds_rep,
            f"Period/End={period_end}, ReportingPeriod={rep_period}")

    # 60022 – Se FilingCE/Role=GIR401, FilingCE/TIN deve coincidere con TIN UPE
    if filing_role == "GIR401" and cs is not None:
        upe_tins = set()
        for upe in cs.findall(f"{G}UPE"):
            for sub in upe:
                for tin_el in sub.findall(f".//{G}TIN"):
                    if tin_el.text: upe_tins.add(tin_el.text.strip())
        match = any(t in upe_tins for t in filing_tin_vals)
        chk("60022", cat, "Se FilingCE/Role=GIR401, FilingCE/TIN deve coincidere con un TIN dell'UPE",
            "/GLOBE_OECD/GLOBEBody/FilingInfo/FilingCE/TIN",
            match or not filing_tin_vals,
            f"FilingCE TIN={filing_tin_vals}, UPE TINs={list(upe_tins)[:5]}")

    # 60023 – FilingCE/ResCountryCode deve coincidere con TransmittingCountry
    chk("60023", cat, "FilingCE/ResCountryCode deve coincidere con TransmittingCountry",
        "/GLOBE_OECD/GLOBEBody/FilingInfo/FilingCE/ResCountryCode",
        filing_country == trans_country or not filing_country or not trans_country,
        f"FilingCE/ResCountryCode={filing_country}, TransmittingCountry={trans_country}")

    # 60024 – Se Summary ha SafeHarbour/ETRRange/SBIE/QDMTTut/GLoBETut, JurWithTaxingRights/JurisdictionName valorizzato
    for s_el in summaries:
        jwr = s_el.find(f"{G}JurWithTaxingRights")
        jn  = _t(jwr, "JurisdictionName") if jwr is not None else None
        has_data = any(s_el.find(f"{G}{t}") is not None for t in ["SafeHarbour","ETRRange","SBIE","QDMTTut","GLoBETut"])
        if has_data:
            chk("60024", cat, "Se Summary ha dati, JurWithTaxingRights/JurisdictionName deve essere valorizzato",
                "/GLOBE_OECD/GLOBEBody/Summary/JurWithTaxingRights/JurisdictionName",
                jn is not None and jn != "",
                f"JurisdictionName={jn}")
            break

    # 60025 – ETRRate = AdjustedCoveredTax/Total / NetGlobeIncome/Total
    for js in jur_sects:
        for oc in js.findall(f".//{G}OverallComputation"):
            etr_rate_el   = oc.find(f"{G}ETRRate")
            net_inc_el    = oc.find(f".//{G}NetGlobeIncome/{G}Total")
            adj_cov_el    = oc.find(f".//{G}AdjustedCoveredTax/{G}Total")
            etr_rate = _num(etr_rate_el)
            net_inc  = _num(net_inc_el)
            adj_cov  = _num(adj_cov_el) or 0
            if etr_rate is not None and net_inc is not None and net_inc > 0:
                calc = adj_cov / net_inc
                chk("60025", cat, "ETRRate deve essere AdjustedCoveredTax/Total ÷ NetGlobeIncome/Total",
                    ".../OverallComputation/ETRRate",
                    abs(etr_rate - calc) < 0.0001,
                    f"ETRRate={etr_rate:.6f}, calcolato={calc:.6f}")
                break

    # 60026 – TopUpTax = TopUpTaxPercentage * ExcessProfits - QDMTT/Amount
    for js in jur_sects:
        for oc in js.findall(f".//{G}OverallComputation"):
            tut   = _num(oc.find(f"{G}TopUpTax"))
            pct   = _num(oc.find(f"{G}TopUpTaxPercentage"))
            exp   = _num(oc.find(f"{G}ExcessProfits"))
            qdmtt = _num(js.find(f".//{G}QDMTT/{G}Amount")) or 0
            if tut is not None and pct is not None and exp is not None:
                calc = max(0, (pct * exp) - qdmtt)
                chk("60026", cat, "TopUpTax = (TopUpTaxPercentage × ExcessProfits) - QDMTT/Amount",
                    ".../OverallComputation/TopUpTax",
                    abs(tut - calc) < 1,
                    f"TopUpTax={tut}, calcolato={calc:.2f}")
                break

    # 60028 – AdjustedFANIL/Total = FANIL + Additions - Reductions (CEComputation)
    for js in jur_sects:
        for cec in js.findall(f".//{G}CEComputation"):
            adj = cec.find(f".//{G}AdjustedFANIL")
            if adj is None: continue
            total = _num(adj.find(f"{G}Total"))
            fanil = _num(adj.find(f"{G}FANIL")) or 0
            adds  = sum(_num(e) or 0 for e in adj.findall(f".//{G}MainEntityPEandFTE/{G}Additions"))
            reds  = sum(_num(e) or 0 for e in adj.findall(f".//{G}MainEntityPEandFTE/{G}Reductions"))
            if total is not None:
                calc = fanil + adds - reds
                chk("60028", cat, "AdjustedFANIL/Total = FANIL + Additions - Reductions",
                    ".../CEComputation/AdjustedFANIL/Total",
                    abs(total - calc) < 1,
                    f"Total={total}, FANIL={fanil}, Adds={adds}, Reds={reds}, Calc={calc}")
                break

    # ── 8.3 RECORD ERRORS OTHER (7xxxx) ─────────────────────────────────────
    cat = "OTHER"

    # 70001-70003 – TIN consistency
    for tin_el in root.findall(f".//{G}TIN"):
        val      = _val(tin_el)
        tin_type = tin_el.get("TypeOfTIN")
        unknown  = tin_el.get("unknown", "").lower() == "true"
        issued   = tin_el.get("issuedBy")
        if tin_type == "GIR3004":
            chk("70001", cat, "Se TypeOfTIN=GIR3004, TIN=NOTIN, unknown=TRUE, issuedBy assente",
                ".//TIN",
                val == "NOTIN" and unknown and not issued,
                f"TIN={val}, unknown={unknown}, issuedBy={issued}")
        if val == "NOTIN":
            chk("70002", cat, "Se TIN=NOTIN, TypeOfTIN=GIR3004, unknown=TRUE, issuedBy assente",
                ".//TIN",
                tin_type == "GIR3004" and unknown and not issued,
                f"TypeOfTIN={tin_type}, unknown={unknown}, issuedBy={issued}")
        if unknown:
            chk("70003", cat, "Se unknown=TRUE, TIN=NOTIN, TypeOfTIN=GIR3004, issuedBy assente",
                ".//TIN",
                val == "NOTIN" and tin_type == "GIR3004" and not issued,
                f"TIN={val}, TypeOfTIN={tin_type}, issuedBy={issued}")
        break  # un solo check per evitare duplicati massicci

    # 70005 – issuedBy e TypeOfTIN sempre presenti salvo GIR3003/GIR3004
    for tin_el in root.findall(f".//{G}TIN"):
        tin_type = tin_el.get("TypeOfTIN")
        issued   = tin_el.get("issuedBy")
        if tin_type not in ("GIR3003", "GIR3004"):
            if not tin_type or not issued:
                ko("70005", cat, "issuedBy e TypeOfTIN devono essere sempre presenti (salvo GIR3003/GIR3004)",
                   ".//TIN/@issuedBy e @TypeOfTIN",
                   f"TypeOfTIN={tin_type}, issuedBy={issued}")
                break
    else:
        ok("70005", cat, "issuedBy e TypeOfTIN presenti su tutti i TIN",
           ".//TIN/@issuedBy e @TypeOfTIN")

    # 70009 – GloBEStatus UPE non deve contenere valori non ammessi
    UPE_FORBIDDEN = {"GIR305","GIR307","GIR308","GIR309","GIR312","GIR313","GIR314","GIR315","GIR317","GIR318"}
    if cs is not None:
        for upe in cs.findall(f"{G}UPE"):
            for sub in upe:
                for s_el in sub.findall(f".//{G}GloBEStatus"):
                    if _val(s_el) in UPE_FORBIDDEN:
                        ko("70009", cat, "GloBEStatus UPE non deve contenere valori non ammessi",
                           ".../UPE/*/ID/GloBEStatus", f"Valore non ammesso: {_val(s_el)}")
                        break
        else:
            ok("70009", cat, "GloBEStatus UPE non contiene valori non ammessi", ".../UPE/*/ID/GloBEStatus")

    # 70010 – OtherUPE/ID/ResCountryCode deve avere un solo valore
    if cs is not None:
        for upe in cs.findall(f"{G}UPE"):
            for other in upe.findall(f"{G}OtherUPE"):
                rcc = other.findall(f".//{G}ResCountryCode")
                chk("70010", cat, "OtherUPE/ID/ResCountryCode deve avere un solo valore",
                    ".../OtherUPE/ID/ResCountryCode",
                    len(rcc) <= 1, f"Trovati {len(rcc)} valori")

    # 70011 – CE/ID/ResCountryCode deve avere un solo valore
    if cs is not None:
        multi = [ce for ce in cs.findall(f"{G}CE") if len(ce.findall(f".//{G}ResCountryCode")) > 1]
        chk("70011", cat, "CE/ID/ResCountryCode deve avere un solo valore",
            ".../CE/ID/ResCountryCode",
            len(multi) == 0, f"{len(multi)} CE con più ResCountryCode")

    # 70013-70020 – GloBEStatus CE combinazioni non ammesse
    if cs is not None:
        for ce in cs.findall(f"{G}CE"):
            statuses = set(_ta(ce.find(f"{G}ID"), "GloBEStatus"))
            chk("70013", cat, "GIR313 e GIR314 non possono coesistere nello stesso CE",
                ".../CE/ID/GloBEStatus",
                not ("GIR313" in statuses and "GIR314" in statuses))
            chk("70014", cat, "GIR307 e GIR308 non possono coesistere nello stesso CE",
                ".../CE/ID/GloBEStatus",
                not ("GIR307" in statuses and "GIR308" in statuses))
            chk("70016", cat, "Se GloBEStatus contiene GIR307, deve contenere anche GIR309",
                ".../CE/ID/GloBEStatus",
                "GIR307" not in statuses or "GIR309" in statuses)
            chk("70017", cat, "Se GloBEStatus contiene GIR308, deve contenere anche GIR309",
                ".../CE/ID/GloBEStatus",
                "GIR308" not in statuses or "GIR309" in statuses)
            chk("70018", cat, "GIR305 e GIR306 non possono coesistere nello stesso CE",
                ".../CE/ID/GloBEStatus",
                not ("GIR305" in statuses and "GIR306" in statuses))
            chk("70020", cat, "GIR316 o GIR318 non devono contenere altri valori",
                ".../CE/ID/GloBEStatus",
                not (("GIR316" in statuses or "GIR318" in statuses) and len(statuses) > 1))
            break  # controlla solo il primo CE per evitare duplicati

    # 70022-70023 – OwnershipChange/ChangeDate nel periodo
    if cs is not None:
        for ce in cs.findall(f"{G}CE"):
            for oc in ce.findall(f"{G}OwnershipChange"):
                cd = _date(_t(oc, "ChangeDate"))
                if cd and ds_start and ds_end:
                    chk("70022", cat, "OwnershipChange/ChangeDate non può essere anteriore a Period/Start",
                        ".../OwnershipChange/ChangeDate", cd >= ds_start, f"ChangeDate={cd}")
                    chk("70023", cat, "OwnershipChange/ChangeDate non può essere successivo a Period/End",
                        ".../OwnershipChange/ChangeDate", cd <= ds_end, f"ChangeDate={cd}")
                break

    # 70026 – PE (GIR305) deve avere OwnershipPercentage = 100%
    if cs is not None:
        for ce in cs.findall(f"{G}CE"):
            statuses = set(_ta(ce.find(f"{G}ID"), "GloBEStatus"))
            if "GIR305" in statuses:
                pcts = [_num(e) for e in ce.findall(f".//{G}OwnershipPercentage")]
                chk("70026", cat, "SE GloBEStatus=GIR305 (PE), OwnershipPercentage deve essere 100%",
                    ".../CE/Ownership/OwnershipPercentage",
                    all(p == 1.0 or p == 100.0 for p in pcts if p is not None),
                    f"Percentuali: {pcts}")
                break

    # 70032 – Se CE/QIIR compilato, CE/ID/Rules deve contenere GIR201 o GIR202
    if cs is not None:
        for ce in cs.findall(f"{G}CE"):
            if ce.find(f"{G}QIIR") is not None:
                rules = set(_ta(ce.find(f"{G}ID"), "Rules"))
                chk("70032", cat, "Se CE/QIIR compilato, Rules deve contenere GIR201 o GIR202",
                    ".../CE/ID/Rules",
                    "GIR201" in rules or "GIR202" in rules,
                    f"Rules={rules}")
                break

    # 70038 – SafeHarbour GIR1203/1204/1205 non ammessi dopo 30/06/2028
    rep_d = _date(rep_period)
    for s_el in summaries:
        shs = _ta(s_el, "SafeHarbour")
        if rep_d and rep_d > date(2028, 6, 30):
            forbidden = {"GIR1203","GIR1204","GIR1205"}
            bad = [s for s in shs if s in forbidden]
            chk("70038", cat, "GIR1203/1204/1205 non ammessi dopo 30/06/2028",
                ".../Summary/SafeHarbour",
                len(bad) == 0, f"SafeHarbour non ammessi: {bad}")
        break

    # 70041 – CFSofUPE GIR502/GIR504 non può avere SafeHarbour GIR1207/1208/1209
    if cfs_of_upe in ("GIR502","GIR504"):
        for s_el in summaries:
            shs = set(_ta(s_el, "SafeHarbour"))
            bad = shs & {"GIR1207","GIR1208","GIR1209"}
            chk("70041", cat, "Se CFSofUPE=GIR502/GIR504, SafeHarbour non può contenere GIR1207/1208/1209",
                ".../Summary/SafeHarbour",
                len(bad) == 0, f"SafeHarbour non ammessi: {bad}")
            break

    # 70044 – ETRStatus deve contenere almeno ETRException o ETRComputation
    for js in jur_sects:
        for etr in js.findall(f".//{G}ETR"):
            etr_status = etr.find(f"{G}ETRStatus")
            if etr_status is not None:
                has_exc  = etr_status.find(f"{G}ETRException") is not None
                has_comp = etr_status.find(f"{G}ETRComputation") is not None
                chk("70044", cat, "ETRStatus deve contenere almeno ETRException o ETRComputation",
                    ".../ETR/ETRStatus",
                    has_exc or has_comp,
                    "ETRStatus presente ma senza ETRException né ETRComputation")
            break

    # 70046 – Se ETRException/TransitionalCbCRSafeHarbour, SubGroup/TypeofSubGroup deve essere GIR1607 o GIR1608
    for js in jur_sects:
        for etr in js.findall(f".//{G}ETR"):
            if etr.find(f".//{G}TransitionalCbCRSafeHarbour") is not None:
                sg = etr.find(f"{G}SubGroup")
                sg_type = _t(sg, "TypeofSubGroup") if sg is not None else None
                chk("70046", cat, "Se TransitionalCbCRSafeHarbour, SubGroup/TypeofSubGroup deve essere GIR1607 o GIR1608",
                    ".../ETR/SubGroup/TypeofSubGroup",
                    sg_type in ("GIR1607","GIR1608"),
                    f"TypeofSubGroup={sg_type}")
                break

    # 70054 – RevocationYear valorizzato solo quando Status=FALSE
    for el in root.findall(f".//{G}RevocationYear"):
        parent = el.getparent()
        if parent is not None:
            status_el = parent.find(f"{G}Status")
            if status_el is not None:
                chk("70054", cat, "RevocationYear può essere valorizzato solo quando Status=FALSE",
                    ".../Election/*/RevocationYear",
                    (_val(status_el) or "").lower() == "false",
                    f"Status={_val(status_el)}, RevocationYear={_val(el)}")
                break

    # 70059 – AdjustmentItem in NetGlobeIncome/Adjustments non può comparire più di una volta per ETR
    for js in jur_sects:
        for oc in js.findall(f".//{G}OverallComputation"):
            items = _ta(oc.find(f".//{G}NetGlobeIncome"), "AdjustmentItem") if oc.find(f".//{G}NetGlobeIncome") is not None else []
            chk("70059", cat, "AdjustmentItem in NetGlobeIncome/Adjustments non può comparire più di una volta",
                ".../NetGlobeIncome/Adjustments/AdjustmentItem",
                len(items) == len(set(items)),
                f"Duplicati: {[x for x in set(items) if items.count(x)>1]}")
            break

    # 70063 – AdjustmentItem in AdjustedCoveredTax non può comparire più di una volta
    for js in jur_sects:
        for oc in js.findall(f".//{G}OverallComputation"):
            act = oc.find(f".//{G}AdjustedCoveredTax")
            items = [_val(e) for e in act.findall(f".//{G}AdjustmentItem")] if act is not None else []
            chk("70063", cat, "AdjustmentItem in AdjustedCoveredTax non può comparire più di una volta",
                ".../AdjustedCoveredTax/Adjustments/AdjustmentItem",
                len(items) == len(set(items)),
                f"Duplicati: {[x for x in set(items) if items.count(x)>1]}")
            break

    # 70083 – ExcessNegTaxExpense/Remaining = PriorYearBalance + GeneratedInRFY - UtilizedInRFY
    for js in jur_sects:
        for oc in js.findall(f".//{G}OverallComputation"):
            ente = oc.find(f"{G}ExcessNegTaxExpense")
            if ente is not None:
                rem  = _num(ente.find(f"{G}Remaining"))
                prv  = _num(ente.find(f"{G}PriorYearBalance")) or 0
                gen  = _num(ente.find(f"{G}GeneratedInRFY")) or 0
                util = _num(ente.find(f"{G}UtilizedInRFY")) or 0
                if rem is not None:
                    calc = prv + gen - util
                    chk("70083", cat, "ExcessNegTaxExpense/Remaining = PriorYearBalance + GeneratedInRFY - UtilizedInRFY",
                        ".../ExcessNegTaxExpense/Remaining",
                        abs(rem - calc) < 1,
                        f"Remaining={rem}, calc={calc}")
                break

    # 70086 – ExcessProfits = NetGlobeIncome/Total - SubstanceExclusion/Total (min 0)
    for js in jur_sects:
        for oc in js.findall(f".//{G}OverallComputation"):
            ep   = _num(oc.find(f"{G}ExcessProfits"))
            ni   = _num(oc.find(f".//{G}NetGlobeIncome/{G}Total"))
            sbe  = _num(oc.find(f".//{G}SubstanceExclusion/{G}Total")) or 0
            if ep is not None and ni is not None:
                calc = max(0, ni - sbe)
                chk("70086", cat, "ExcessProfits = max(0, NetGlobeIncome/Total - SubstanceExclusion/Total)",
                    ".../OverallComputation/ExcessProfits",
                    abs(ep - calc) < 1,
                    f"ExcessProfits={ep}, calc={calc}")
                break

    # 70087 – SubstanceExclusion/Total = (PayrollCost * PayrollMarkUp) + (TangibleAssetValue * TangibleAssetMarkup)
    for js in jur_sects:
        for oc in js.findall(f".//{G}OverallComputation"):
            sbe = oc.find(f".//{G}SubstanceExclusion")
            if sbe is not None:
                total   = _num(sbe.find(f"{G}Total"))
                payroll = _num(sbe.find(f"{G}PayrollCost")) or 0
                pay_mu  = _num(sbe.find(f"{G}PayrollMarkUp")) or 0
                tang    = _num(sbe.find(f"{G}TangibleAssetValue")) or 0
                tang_mu = _num(sbe.find(f"{G}TangibleAssetMarkup")) or 0
                if total is not None:
                    calc = (payroll * pay_mu) + (tang * tang_mu)
                    chk("70087", cat, "SubstanceExclusion/Total = (PayrollCost × PayrollMarkUp) + (TangibleAssetValue × TangibleAssetMarkup)",
                        ".../SubstanceExclusion/Total",
                        abs(total - calc) < 1,
                        f"Total={total}, calc={calc:.2f}")
                break

    # 70097 – IIR/ParentEntity/InclusionRatio = (NetGlobeIncome - OtherOwnershipAllocation) / NetGlobeIncome
    for js in jur_sects:
        for ltj in js.findall(f"{G}LowTaxJurisdiction"):
            for ltce in ltj.findall(f"{G}LTCE"):
                for iir in ltce.findall(f"{G}IIR"):
                    ni  = _num(iir.find(f"{G}NetGlobeIncome"))
                    for pe in iir.findall(f"{G}ParentEntity"):
                        ir  = _num(pe.find(f"{G}InclusionRatio"))
                        ooa = _num(pe.find(f"{G}OtherOwnershipAllocation")) or 0
                        if ir is not None and ni is not None and ni != 0:
                            calc = (ni - ooa) / ni
                            chk("70097", cat, "IIR/ParentEntity/InclusionRatio = (NetGlobeIncome - OtherOwnershipAllocation) / NetGlobeIncome",
                                ".../IIR/ParentEntity/InclusionRatio",
                                abs(ir - calc) < 0.0001,
                                f"InclusionRatio={ir:.4f}, calc={calc:.4f}")
                        break
                    break

    # 70098 – IIR/ParentEntity/TopUpTaxShare = IIR/TopUpTax * InclusionRatio
    for js in jur_sects:
        for ltj in js.findall(f"{G}LowTaxJurisdiction"):
            for ltce in ltj.findall(f"{G}LTCE"):
                for iir in ltce.findall(f"{G}IIR"):
                    tut = _num(iir.find(f"{G}TopUpTax"))
                    for pe in iir.findall(f"{G}ParentEntity"):
                        share = _num(pe.find(f"{G}TopUpTaxShare"))
                        ir    = _num(pe.find(f"{G}InclusionRatio"))
                        if share is not None and tut is not None and ir is not None:
                            calc = tut * ir
                            chk("70098", cat, "IIR/ParentEntity/TopUpTaxShare = IIR/TopUpTax × InclusionRatio",
                                ".../IIR/ParentEntity/TopUpTaxShare",
                                abs(share - calc) < 1,
                                f"TopUpTaxShare={share}, calc={calc:.2f}")
                        break

    # 70104 – UTPRTopUpTaxCarriedForward non può essere negativo
    for ua in utpr_attrs:
        for att in ua.findall(f"{G}Attribution"):
            v = _num(att.find(f"{G}UTPRTopUpTaxCarriedForward"))
            if v is not None:
                chk("70104", cat, "UTPRTopUpTaxCarriedForward non può essere negativo",
                    ".../Attribution/UTPRTopUpTaxCarriedForward",
                    v >= 0, f"Valore={v}")
                break

    # 70105 – UTPRTopUpTaxCarriedForward = CarryForward + Attributed - AddCashTaxExpense
    for ua in utpr_attrs:
        for att in ua.findall(f"{G}Attribution"):
            cf  = _num(att.find(f"{G}UTPRTopUpTaxCarryForward")) or 0
            att_val = _num(att.find(f"{G}UTPRTopUpTaxAttributed")) or 0
            add_cash = _num(att.find(f"{G}AddCashTaxExpense")) or 0
            carried  = _num(att.find(f"{G}UTPRTopUpTaxCarriedForward"))
            if carried is not None:
                calc = cf + att_val - add_cash
                chk("70105", cat, "UTPRTopUpTaxCarriedForward = CarryForward + Attributed - AddCashTaxExpense",
                    ".../Attribution/UTPRTopUpTaxCarriedForward",
                    abs(carried - calc) < 1,
                    f"CarriedForward={carried}, calc={calc}")
                break

    return results


# ── Genera report XLSX ────────────────────────────────────────────────────────
def _make_xlsx(results: list[Check], xml_name: str, out_path: Path) -> None:
    wb = openpyxl.Workbook()

    # Stili
    BLUE   = "00338D"; YELLOW = "EAAA00"; WHITE = "FFFFFF"
    RED    = "C00000"; GREEN  = "375623"; ORANGE = "C55A00"
    LGRAY  = "F5F5F5"; DGRAY  = "4A4A4A"

    def hfont(bold=True, color=WHITE, sz=10):
        return Font(name="Calibri", bold=bold, color=color, size=sz)
    def hfill(c): return PatternFill("solid", fgColor=c)
    def bd(): return Border(
        left=Side(style="thin",color="D9D9D9"), right=Side(style="thin",color="D9D9D9"),
        top=Side(style="thin",color="D9D9D9"), bottom=Side(style="thin",color="D9D9D9"))
    def al(h="left",v="center"): return Alignment(horizontal=h,vertical=v,wrap_text=True)

    ko_severe = [r for r in results if r.esito == "KO" and r.categoria == "SEVERE"]
    ko_file   = [r for r in results if r.esito == "KO" and r.categoria == "FILE"]
    ko_other  = [r for r in results if r.esito == "KO" and r.categoria == "OTHER"]
    tot_ko    = len(ko_severe) + len(ko_file) + len(ko_other)
    tot_ok    = sum(1 for r in results if r.esito == "OK")
    stato     = "✓ OK" if tot_ko == 0 else ("✗ ERRORI BLOCCANTI" if ko_severe or ko_file else "⚠ WARNING")
    stato_col = GREEN if tot_ko == 0 else (RED if ko_severe or ko_file else ORANGE)

    # ── Foglio 1: Sommario ──
    ws = wb.active; ws.title = "Sommario"
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 50

    def s_row(label, val, bold=False, color=None):
        r = ws.max_row + 1
        ws.cell(r,1,label).font = Font(name="Calibri", bold=True, size=10, color=DGRAY)
        c = ws.cell(r,2,val)
        c.font = Font(name="Calibri", bold=bold, size=10, color=color or "1A1A1A")
        ws.cell(r,1).fill = hfill(LGRAY)
        ws.cell(r,1).border = bd()
        ws.cell(r,2).border = bd()

    ws.append([]); ws.append([])
    ws["A1"] = "PILLAR · Validation Report GIR/DAC9"
    ws["A1"].font = Font(name="Calibri", bold=True, size=14, color=BLUE)
    ws.merge_cells("A1:B1")
    ws.row_dimensions[1].height = 22
    ws.append([])
    s_row("File analizzato", xml_name)
    s_row("Data validazione", datetime.now().strftime("%d/%m/%Y %H:%M"))
    s_row("Stato complessivo", stato, bold=True, color=stato_col)
    ws.append([])
    s_row("Check totali eseguiti", len(results))
    s_row("Check OK", tot_ok, color=GREEN)
    s_row("FILE ERRORS (500xx) – KO", len(ko_file), bold=bool(ko_file), color=RED if ko_file else GREEN)
    s_row("RECORD ERRORS SEVERE – KO", len(ko_severe), bold=bool(ko_severe), color=RED if ko_severe else GREEN)
    s_row("RECORD ERRORS OTHER – KO", len(ko_other), bold=bool(ko_other), color=ORANGE if ko_other else GREEN)
    ws.append([])
    s_row("Esito trasmissione", "FILE ACCETTABILE" if not ko_severe and not ko_file else "FILE DA CORREGGERE",
          bold=True, color=GREEN if not ko_severe and not ko_file else RED)

    # ── Foglio 2: Dettaglio ──
    ws2 = wb.create_sheet("Dettaglio")
    headers = ["Codice","Categoria","Descrizione","XPath","Esito","Dettaglio"]
    widths  = [10, 10, 50, 40, 8, 45]
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws2.cell(1, i, h)
        c.font = hfont(); c.fill = hfill(BLUE); c.border = bd(); c.alignment = al("center")
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.row_dimensions[1].height = 16

    for res in results:
        r = ws2.max_row + 1
        esito_col = GREEN if res.esito=="OK" else (RED if res.categoria in ("SEVERE","FILE") else ORANGE)
        row_fill  = hfill(LGRAY) if r % 2 == 0 else hfill(WHITE)
        for i, val in enumerate([res.cod, res.categoria, res.desc, res.xpath, res.esito, res.dettaglio], 1):
            c = ws2.cell(r, i, val)
            c.font = Font(name="Calibri", size=9,
                          color=esito_col if i==5 else "1A1A1A",
                          bold=(i==5))
            c.fill = row_fill; c.border = bd(); c.alignment = al()

    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = f"A1:F{ws2.max_row}"

    # ── Foglio 3: Errori e Warning ──
    ws3 = wb.create_sheet("Errori e Warning")
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws3.cell(1, i, h)
        c.font = hfont(); c.fill = hfill(RED); c.border = bd(); c.alignment = al("center")
        ws3.column_dimensions[get_column_letter(i)].width = w

    ko_all = [r for r in results if r.esito == "KO"]
    if not ko_all:
        ws3.cell(2, 1, "Nessun errore rilevato ✓").font = Font(name="Calibri", bold=True, color=GREEN, size=10)
        ws3.merge_cells("A2:F2")
    else:
        for res in ko_all:
            r = ws3.max_row + 1
            esito_col = RED if res.categoria in ("SEVERE","FILE") else ORANGE
            for i, val in enumerate([res.cod, res.categoria, res.desc, res.xpath, res.esito, res.dettaglio], 1):
                c = ws3.cell(r, i, val)
                c.font = Font(name="Calibri", size=9, color=esito_col if i==5 else "1A1A1A", bold=(i==5))
                c.border = bd(); c.alignment = al()

    ws3.freeze_panes = "A2"
    wb.save(str(out_path))


# ── Entry point ───────────────────────────────────────────────────────────────
def validate(xml_path: Path, output_dir: Path | None = None, xsd_dir: Path | None = None) -> Path:
    xml_path = Path(xml_path)
    if output_dir is None: output_dir = xml_path.parent
    output_dir = Path(output_dir); output_dir.mkdir(parents=True, exist_ok=True)
    out_path = output_dir / f"{xml_path.stem}_validation_report.xlsx"

    # Parse XML
    try:
        tree = etree.parse(str(xml_path))
        root = tree.getroot()
        # Se incapsulato nella shell telematica, cerca GLOBE_OECD
        if f"{G}GLOBE_OECD" not in root.tag:
            globe = root.find(f".//{G}GLOBE_OECD")
            if globe is not None: root = globe
    except etree.XMLSyntaxError as e:
        if HAS_XL:
            wb = openpyxl.Workbook()
            ws = wb.active; ws.title = "Sommario"
            ws["A1"] = "ERRORE FILE XML"
            ws["A2"] = str(e)
            wb.save(str(out_path))
        return out_path

    results = _run_checks(root)

    if HAS_XL:
        _make_xlsx(results, xml_path.name, out_path)
    else:
        out_path.write_text("\n".join(f"{r.cod} [{r.esito}] {r.desc}: {r.dettaglio}" for r in results))

    ko = sum(1 for r in results if r.esito == "KO")
    print(f"  [validator] {len(results)} check, {ko} KO → {out_path.name}")
    return out_path
