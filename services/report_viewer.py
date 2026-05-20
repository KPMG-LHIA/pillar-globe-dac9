"""
report_viewer.py – Report HTML navigabile da XML GloBE OECD v2
==============================================================
Lookup codici aggiornati da GLOBEXML_v1.0.xsd ufficiale AdE.
Sezioni estratte: MessageSpec, FilingInfo, CorporateStructure,
Summary, JurisdictionSection (ETR, SBIE, QDMTT, LowTaxJurisdiction,
Election, UTPRAttribution), ExcludedEntity.
"""
from __future__ import annotations
import json
from pathlib import Path
from lxml import etree

_CODES = json.loads((Path(__file__).parent / "globe_codes.json").read_text(encoding="utf-8"))

NS = {
    "globe": "urn:oecd:ties:globe:v2",
    "stf":   "urn:oecd:ties:globestf:v5",
    "iso":   "urn:oecd:ties:isoglobetypes:v1",
    "tm":    "www.agenziaentrate.gov.it:specificheTecniche:telent:v1",
    "girtel":"urn:www.agenziaentrate.gov.it:specificheTecniche:girtel",
}

def _g(el, path):
    if el is None: return None
    f = el.find(path, NS)
    if f is None and "stf:" in path:
        # Fallback: cerca senza namespace (alcuni file omettono il prefisso stf:)
        bare = path.split(":")[-1]
        f = el.find(bare)
    return f.text.strip() if f is not None and f.text else None

def _ga(el, path):
    if el is None: return []
    return [e.text.strip() for e in el.findall(path, NS) if e.text]

def _fmt(v, cur=""):
    if v is None: return "—"
    try:
        n = float(v)
        s = f"{int(n):,}".replace(",", ".") if n == int(n) else f"{n:,.4f}".replace(",","X").replace(".",",").replace("X",".")
        return f"{s} {cur}".strip() if cur else s
    except: return str(v)

def _pct(v):
    if v is None: return "—"
    try: return f"{float(v)*100:.2f}%"
    except: return str(v)

def _bool(v):
    if v is None: return "—"
    return "✓ Sì" if str(v).lower() == "true" else "✗ No"

# ── Lookup da globe_codes.json ───────────────────────────────────────────────
MSG_TYPE_INDIC = _CODES["MSG_TYPE_INDIC"]
ROLE           = _CODES["ROLE"]
CFS            = _CODES["CFS"]
GLOBE_STATUS   = _CODES["GLOBE_STATUS"]
RULES          = _CODES["RULES"]
OWN_TYPE       = _CODES["OWN_TYPE"]
EXCL_ENTITY    = _CODES["EXCL_ENTITY"]
SAFE_HARBOUR   = _CODES["SAFE_HARBOUR"]
ETR_RANGE      = _CODES["ETR_RANGE"]
QDMTT_UT       = _CODES["QDMTT_UT"]
GLOBE_UT       = _CODES["GLOBE_UT"]
ADJ_ITEM       = _CODES["ADJ_ITEM"]
CURR_ADJ       = _CODES["CURR_ADJ"]
FINAL_ADJ      = _CODES["FINAL_ADJ"]
TYPE_INDIC     = _CODES["TYPE_INDIC"]
SUBGROUP_TYPE  = _CODES["SUBGROUP_TYPE"]
ETR_BADGE      = _CODES["ETR_BADGE"]

def lk(d, k): return d.get(k, k) if k else "—"

# ── ETR badge colore ──────────────────────────────────────────────────────────

def _etr_badge(etr_code):
    badge = ETR_BADGE 
    label = lk(ETR_RANGE, etr_code)
    if etr_code in badge["danger"]:  return f'<span class="badge-red">{label}</span>'
    if etr_code in badge["warning"]: return f'<span class="badge-yellow">{label}</span>'
    if etr_code in badge["ok"]:      return f'<span class="badge-green">{label}</span>'
    return f'<span class="badge-gray">{label}</span>'

# ── Parser ────────────────────────────────────────────────────────────────────
def parse_globe(xml_path):
    tree = etree.parse(str(xml_path))
    root = tree.getroot()
    globe = root if "GLOBE_OECD" in root.tag else root.find(".//globe:GLOBE_OECD", NS)
    if globe is None: raise ValueError("globe:GLOBE_OECD non trovato")

    body = globe.find("globe:GLOBEBody", NS)
    ms   = globe.find("globe:MessageSpec", NS)
    fi   = body.find("globe:FilingInfo", NS) if body is not None else None
    gs   = body.find("globe:GeneralSection", NS) if body is not None else None
    fce  = fi.find("globe:FilingCE", NS) if fi is not None else None
    acc  = fi.find("globe:AccountingInfo", NS) if fi is not None else None
    per  = fi.find("globe:Period", NS) if fi is not None else None
    ds   = fi.find("globe:DocSpec", NS) if fi is not None else None

    # MessageSpec
    msg_spec = {
        "ref_id":      _g(ms, "globe:MessageRefId"),
        "ref_id_warn":  "{Guid" in (_g(ms, "globe:MessageRefId") or ""),
        "type_indic":  lk(MSG_TYPE_INDIC, _g(ms, "globe:MessageTypeIndic")),
        "period":      _g(ms, "globe:ReportingPeriod"),
        "timestamp":   _g(ms, "globe:Timestamp"),
        "contact":     _g(ms, "globe:Contact"),
        "warning":     _g(ms, "globe:Warning"),
        "transmitting":_g(ms, "globe:TransmittingCountry"),
        "receiving":   _g(ms, "globe:ReceivingCountry"),
    }

    # FilingInfo
    filing = {
        "name":     _g(fce, "globe:Name"),
        "country":  _g(fce, "globe:ResCountryCode"),
        "tin":      _g(fce, "globe:TIN"),
        "role":     lk(ROLE, _g(fce, "globe:Role")),
        "cfs":      lk(CFS, _g(acc, "globe:CFSofUPE")),
        "fas":      _g(acc, "globe:FAS"),
        "currency": _g(acc, "globe:Currency"),
        "start":    _g(per, "globe:Start"),
        "end":      _g(per, "globe:End"),
        "mne_name": _g(fi, "globe:NameMNE"),
        "add_info": _g(fi, "globe:AdditionalInfo"),
        "doc_type": lk(TYPE_INDIC, _g(ds, "stf:DocTypeIndic")),
        "doc_ref":  _g(ds, "stf:DocRefId"),
        "doc_ref_warn": "{Guid" in (_g(ds, "stf:DocRefId") or ""),
    }

    # CorporateStructure
    cs = gs.find("globe:CorporateStructure", NS) if gs is not None else None
    entities, excluded_entities = [], []
    if cs is not None:
        for upe_el in cs.findall("globe:UPE", NS):
            for sub in upe_el:
                idd = sub.find("globe:ID", NS)
                if idd is not None:
                    entities.append({
                        "type": "UPE",
                        "name": _g(idd, "globe:Name"),
                        "country": ", ".join(_ga(idd, "globe:ResCountryCode")),
                        "tin": _g(idd, "globe:TIN"),
                        "status": ", ".join([lk(GLOBE_STATUS, s) for s in _ga(idd, "globe:GlobeStatus")]),
                        "rules": [lk(RULES, r) for r in _ga(idd, "globe:Rules")],
                        "ownership": [],
                    })
        for ce_el in cs.findall("globe:CE", NS):
            idd = ce_el.find("globe:ID", NS)
            owns = [{
                "type": lk(OWN_TYPE, _g(ow, "globe:OwnershipType")),
                "tin":  _g(ow, "globe:TIN"),
                "pct":  _pct(_g(ow, "globe:OwnershipPercentage")),
            } for ow in ce_el.findall("globe:Ownership", NS)]
            if idd is not None:
                entities.append({
                    "type": "CE",
                    "name": _g(idd, "globe:Name"),
                    "country": ", ".join(_ga(idd, "globe:ResCountryCode")),
                    "tin": _g(idd, "globe:TIN"),
                    "status": ", ".join([lk(GLOBE_STATUS, s) for s in _ga(idd, "globe:GlobeStatus")]),
                    "rules": [lk(RULES, r) for r in _ga(idd, "globe:Rules")],
                    "ownership": owns,
                })
        for ex in cs.findall("globe:ExcludedEntity", NS):
            excluded_entities.append({
                "name":   _g(ex, "globe:Name"),
                "type":   lk(EXCL_ENTITY, _g(ex, "globe:Type")),
                "change": _bool(_g(ex, "globe:Change")),
            })

    # Summary
    summaries = []
    for s in (body.findall("globe:Summary", NS) if body is not None else []):
        jn  = s.find("globe:Jurisdiction", NS)
        ds2 = s.find("globe:DocSpec", NS)
        sbie = s.find("globe:SBIE", NS)
        jur_name = (_g(jn, "globe:JurisdictionName") if jn is not None else None) or _g(s, "globe:Jurisdiction")
        etr_code = _g(s, "globe:ETRRange")
        summaries.append({
            "jurisdiction":  jur_name,
            "etr_code":      etr_code,
            "safe_harbours": [lk(SAFE_HARBOUR, sh) for sh in _ga(s, "globe:SafeHarbour")],
            "qdmtt_ut":      lk(QDMTT_UT, _g(s, "globe:QDMTTut")),
            "globe_ut":      lk(GLOBE_UT,  _g(s, "globe:GLoBETut")),
            "sbie_na":       _bool(_g(sbie, "globe:NotApplicable")) if sbie is not None else "—",
            "sbie_notut":    _bool(_g(sbie, "globe:NoTut")) if sbie is not None else "—",
            "doc_type":      lk(TYPE_INDIC, _g(ds2, "stf:DocTypeIndic")),
        })

    # JurisdictionSections
    jurisdictions = []
    for js in (body.findall("globe:JurisdictionSection", NS) if body is not None else []):
        cur = _g(js, "globe:LocalCurrency") or ""
        ds3 = js.find("globe:DocSpec", NS)
        oc  = js.find(".//globe:OverallComputation", NS)
        exc_el = js.find(".//globe:ETRException", NS)
        sg_el  = js.find(".//globe:SubGroup", NS)
        # ETRException: estrai tipo e dati da tutti gli ETR nella JurisdictionSection
        etr_exception = None
        etr_list = js.findall(".//{%s}ETR" % NS["globe"])
        if etr_list:
            exc_entries = []
            for etr in etr_list:
                sg = etr.find("{%s}SubGroup" % NS["globe"])
                exc_node = etr.find(".//{%s}ETRException" % NS["globe"])
                if exc_node is None: continue
                sg_info = None
                if sg is not None:
                    sg_tin = sg.find("{%s}TIN" % NS["globe"])
                    sg_type = sg.find("{%s}TypeofSubGroup" % NS["globe"])
                    tin_val = sg_tin.text.strip() if sg_tin is not None and sg_tin.text else "NOTIN"
                    type_val = sg_type.text.strip() if sg_type is not None and sg_type.text else ""
                    from_lk = SUBGROUP_TYPE.get(type_val, type_val)
                    sg_info = f"{from_lk} · TIN: {tin_val}"
                for child in exc_node:
                    exc_type = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                    rev  = child.find("{%s}Revenue" % NS["globe"])
                    prof = child.find("{%s}Profit" % NS["globe"])
                    it   = child.find("{%s}IncomeTax" % NS["globe"])
                    exc_entries.append({
                        "type":       exc_type,
                        "subgroup":   sg_info,
                        "revenue":    _fmt(rev.text if rev is not None else None),
                        "profit":     _fmt(prof.text if prof is not None else None),
                        "income_tax": _fmt(it.text if it is not None else None),
                    })
                    break
            etr_exception = exc_entries if exc_entries else None
        subgroup_type = None
        if sg_el is not None:
            sg_tin  = _g(sg_el, "globe:TIN")
            sg_type = _g(sg_el, "globe:TypeofSubGroup")
            subgroup_type = f"{sg_type} (TIN: {sg_tin})" if sg_tin else sg_type
        sbe = js.find(".//globe:SubstanceExclusion", NS)
        qde = js.find(".//globe:QDMTT", NS)
        ene = js.find(".//globe:ExcessNegTaxExpense", NS)
        etr_el = js.find(".//globe:ETR", NS)
        ltj    = js.find("globe:LowTaxJurisdiction", NS)

        # Adjustments decodificati
        def _parse_adjs(parent_path, code_dict):
            results = []
            if oc is None: return results
            parent = oc.find(parent_path, NS)
            if parent is None: return results
            for adj in parent.findall("globe:Adjustments", NS):
                amounts = [_fmt(a.text, cur) for a in adj.findall("globe:Amount", NS)]
                code    = _g(adj, "globe:AdjustmentItem")
                results.append({"amounts": amounts, "label": lk(code_dict, code)})
            return results

        # CE Computations
        ce_comps = []
        for cec in js.findall(".//globe:CEComputation", NS):
            ce_comps.append({
                "tin":         _g(cec, "globe:TIN"),
                "adj_fanil":   _fmt(_g(cec, ".//globe:AdjustedFANIL/globe:Total"), cur),
                "fanil":       _fmt(_g(cec, ".//globe:AdjustedFANIL/globe:FANIL"), cur),
                "net_income":  _fmt(_g(cec, ".//globe:NetGlobeIncome/globe:Total"), cur),
                "adj_cov_tax": _fmt(_g(cec, ".//globe:AdjustedCoveredTax/globe:Total"), cur),
                "defer_total": _fmt(_g(cec, ".//globe:DeferTaxAdjustAmt/globe:Total"), cur),
            })

        # Election
        elections = []
        el_el = js.find(".//globe:Election", NS)
        if el_el is not None:
            for child in el_el:
                tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                val = child.text.strip() if child.text else ""
                if val in ("true", "false"):
                    elections.append(f"{tag}: {_bool(val)}")
                elif val:
                    elections.append(f"{tag}: {val}")

        # LowTaxJurisdiction
        ltj_data = None
        if ltj is not None:
            ltce_list = []
            for ltce in ltj.findall("globe:LTCE", NS):
                iir_list = []
                for iir in ltce.findall("globe:IIR", NS):
                    parents = []
                    for pe in iir.findall("globe:ParentEntity", NS):
                        parents.append({
                            "tin":       _g(pe, "globe:TIN"),
                            "country":   _g(pe, "globe:ResCountryCode"),
                            "inc_ratio": _pct(_g(pe, "globe:InclusionRatio")),
                            "tut_share": _fmt(_g(pe, "globe:TopUpTaxShare"), cur),
                            "top_up":    _fmt(_g(pe, "globe:TopUpTax"), cur),
                        })
                    iir_list.append({
                        "net_income": _fmt(_g(iir, "globe:NetGlobeIncome"), cur),
                        "top_up":     _fmt(_g(iir, "globe:TopUpTax"), cur),
                        "parents":    parents,
                    })
                ltce_list.append({"tin": _g(ltce, "globe:TIN"), "iir": iir_list})
            ltj_data = {
                "total_tut": _fmt(_g(ltj, "globe:TopUpTaxAmount"), cur),
                "ltce":      ltce_list,
            }

        jurisdictions.append({
            "jurisdiction": _g(js, "globe:Jurisdiction") or "",
            "currency":     cur,
            "doc_type":     lk(TYPE_INDIC, _g(ds3, "stf:DocTypeIndic")),
            "jur_taxing":   _ga(js.find("globe:JurWithTaxingRights", NS), "globe:JurisdictionName") if js.find("globe:JurWithTaxingRights", NS) is not None else [],
            "ce_computations": ce_comps,
            "overall": {
                "fanil":          _fmt(_g(oc, "globe:FANIL"), cur),
                "adj_fanil":      _fmt(_g(oc, "globe:AdjustedFANIL"), cur),
                "net_income":     _fmt(_g(oc, ".//globe:NetGlobeIncome/globe:Total"), cur),
                "net_income_adjs":_parse_adjs(".//globe:NetGlobeIncome", ADJ_ITEM),
                "income_tax":     _fmt(_g(oc, "globe:IncomeTaxExpense"), cur),
                "etr_rate":       _pct(_g(oc, "globe:ETRRate")),
                "top_up_pct":     _pct(_g(oc, "globe:TopUpTaxPercentage")),
                "adj_cov_tax":    _fmt(_g(oc, ".//globe:AdjustedCoveredTax/globe:Total"), cur),
                "aggr_curr":      _fmt(_g(oc, ".//globe:AdjustedCoveredTax/globe:AggregrateCurrentTax"), cur),
                "adj_cov_adjs":   _parse_adjs(".//globe:AdjustedCoveredTax", FINAL_ADJ),
                "excess_profits": _fmt(_g(oc, "globe:ExcessProfits"), cur),
                "top_up_tax":     _fmt(_g(oc, "globe:TopUpTax"), cur),
            } if oc is not None else None,
            "substance": {
                "total":       _fmt(_g(sbe, "globe:Total"), cur),
                "payroll":     _fmt(_g(sbe, "globe:PayrollCost"), cur),
                "payroll_mu":  _pct(_g(sbe, "globe:PayrollMarkUp")),
                "tangible":    _fmt(_g(sbe, "globe:TangibleAssetValue"), cur),
                "tangible_mu": _pct(_g(sbe, "globe:TangibleAssetMarkup")),
            } if sbe is not None else None,
            "qdmtt": {
                "amount":   _fmt(_g(qde, "globe:Amount"), cur),
                "fas":      _g(qde, "globe:FAS"),
                "sbie":     _bool(_g(qde, "globe:SBIEAvailable")),
                "de_min":   _bool(_g(qde, "globe:DeMinAvailable")),
                "currency": _g(qde, "globe:Currency"),
            } if qde is not None else None,
            "excess_neg": {
                "prior":     _fmt(_g(ene, "globe:PriorYearBalance"), cur),
                "generated": _fmt(_g(ene, "globe:GeneratedInRFY"), cur),
                "utilized":  _fmt(_g(ene, "globe:UtilizedInRFY"), cur),
                "remaining": _fmt(_g(ene, "globe:Remaining"), cur),
            } if ene is not None else None,
            "elections":  elections,
            "etr_exception": etr_exception,
            "subgroup_type": subgroup_type,
            "ltj":        ltj_data,
        })

    # UTPRAttribution
    utpr_sections = []
    for ua in (body.findall("globe:UTPRAttribution", NS) if body is not None else []):
        ds_u = ua.find("globe:DocSpec", NS)
        attrs = []
        for att in ua.findall("globe:Attribution", NS):
            attrs.append({
                "country":     _g(att, "globe:ResCountryCode"),
                "carry_fwd":   _fmt(_g(att, "globe:UTPRTopUpTaxCarryForward")),
                "employees":   _fmt(_g(att, "globe:Employees")),
                "tangible":    _fmt(_g(att, "globe:TangibleAssetValue")),
                "pct":         _pct(_g(att, "globe:UTPRPercentage")),
                "attributed":  _fmt(_g(att, "globe:UTPRTopUpTaxAttributed")),
                "add_cash":    _fmt(_g(att, "globe:AddCashTaxExpense")),
                "carried_fwd": _fmt(_g(att, "globe:UTPRTopUpTaxCarriedForward")),
            })
        utpr_sections.append({
            "doc_type":   lk(TYPE_INDIC, _g(ds_u, "stf:DocTypeIndic")),
            "attributions": attrs,
        })

    return {
        "msg_spec":    msg_spec,
        "filing":      filing,
        "entities":    entities,
        "excluded":    excluded_entities,
        "summaries":   summaries,
        "jurisdictions": jurisdictions,
        "utpr":        utpr_sections,
    }

# ── HTML builder ──────────────────────────────────────────────────────────────
def _html(data, xml_name):
    d = data; ms = d["msg_spec"]; fi = d["filing"]

    def row(l, v): return f'<tr><td class="lbl">{l}</td><td>{v or "—"}</td></tr>'
    def tbl(rows, heads=None):
        h = "" if not heads else "<thead><tr>" + "".join(f"<th>{x}</th>" for x in heads) + "</tr></thead>"
        return f"<table>{h}<tbody>{rows}</tbody></table>"
    def sec(title, content, sid=""):
        a = f' id="{sid}"' if sid else ""
        return f'<section{a}><h2>{title}</h2>{content}</section>'
    def card(title, content):
        return f'<div class="icard"><h3>{title}</h3>{content}</div>'

    # ── Sommario ──
    msg_type_class = ""
    if "GIR101" in ms["type_indic"]: msg_type_class = "badge-blue"
    elif "GIR102" in ms["type_indic"]: msg_type_class = "badge-yellow"
    elif "GIR103" in ms["type_indic"]: msg_type_class = "badge-gray"

    sr = (row("File", xml_name) +
          row("Periodo rendicontazione", ms["period"]) +
          row("Periodo (Start → End)", f"{fi['start']} → {fi['end']}") +
          row("Gruppo MNE", fi["mne_name"]) +
          row("Entità dichiarante (FilingCE)", fi["name"]) +
          row("Paese FilingCE", fi["country"]) +
          row("TIN FilingCE", fi["tin"]) +
          row("Ruolo", fi["role"]) +
          row("Valuta consolidato", fi["currency"]) +
          row("Standard contabile (FAS)", fi["fas"]) +
          row("CFSofUPE", fi["cfs"]) +
          (row("Informazioni aggiuntive", fi["add_info"]) if fi.get("add_info") else ""))
    mr = (row("MessageTypeIndic", f'<span class="{msg_type_class}">{ms["type_indic"]}</span>') +
          row("MessageRefId", ('⚠ <span class="badge-yellow">Placeholder non risolto: ' + ms["ref_id"] + '</span>' if ms.get("ref_id_warn") else ms["ref_id"])) +
          row("Timestamp", ms["timestamp"]) +
          row("Paese trasmittente", ms["transmitting"]) +
          row("Paese ricevente", ms["receiving"]) +
          (row("Contatto", ms["contact"]) if ms.get("contact") and ms["contact"].replace(",","").replace("Email:","").replace("email:","").strip() else "") +
          row("DocTypeIndic", fi["doc_type"]) +
          row("DocRefId", ('⚠ <span class="badge-yellow">Placeholder non risolto: ' + fi["doc_ref"] + '</span>' if fi.get("doc_ref_warn") else fi["doc_ref"])) +
          (row("⚠ Warning", f'<span class="badge-yellow">{ms["warning"]}</span>') if ms.get("warning") else ""))
    html_som = f'<div class="g2"><div>{card("Gruppo e Periodo", tbl(sr))}</div><div>{card("Messaggio", tbl(mr))}</div></div>'

    # ── Corporate Structure ──
    er = ""
    for e in d["entities"]:
        b = '<span class="bu">UPE</span>' if e["type"] == "UPE" else '<span class="bc">CE</span>'
        owns = "<br>".join(f"<small>{o['type']} · {o['tin']} · {o['pct']}</small>" for o in e["ownership"]) or "—"
        er += f"<tr><td>{b}</td><td><strong>{e['name'] or '—'}</strong></td><td>{e['country'] or '—'}</td><td><code>{e['tin'] or '—'}</code></td><td>{e['status']}</td><td><small>{'<br>'.join(e['rules']) or '—'}</small></td><td>{owns}</td></tr>"
    html_corp = tbl(er, ["Tipo", "Nome", "Paese", "TIN", "GlobeStatus", "Regole", "Ownership"])

    if d["excluded"]:
        ex_rows = "".join(f"<tr><td>{e['name']}</td><td>{e['type']}</td><td>{e['change']}</td></tr>" for e in d["excluded"])
        html_corp += f"<h3 style='margin-top:16px'>Excluded Entities</h3>" + tbl(ex_rows, ["Nome", "Tipo", "Variazione"])

    # ── Summary ──
    sumr = ""
    for s in d["summaries"]:
        sh = ", ".join(s["safe_harbours"]) or "—"
        sumr += (f"<tr><td><strong>{s['jurisdiction'] or '—'}</strong></td>"
                 f"<td>{_etr_badge(s['etr_code'])}</td>"
                 f"<td><small>{sh}</small></td>"
                 f"<td>{s['qdmtt_ut']}</td>"
                 f"<td>{s['globe_ut']}</td>"
                 f"<td>{s['sbie_na']} / {s['sbie_notut']}</td>"
                 f"<td>{s['doc_type']}</td></tr>")
    html_sum = tbl(sumr, ["Giurisdizione", "ETR Range", "Safe Harbour", "QDMTTut", "GloBETut", "SBIE (NA/NoTuT)", "DocType"])

    # ── Jurisdiction Sections ──
    jcards = ""
    for j in d["jurisdictions"]:
        ov = j["overall"]; cur = j["currency"]

        def adj_rows(adjs):
            if not adjs: return ""
            r = ""
            for a in adjs:
                amts = " / ".join(a["amounts"])
                r += f"<tr><td colspan='2' class='adj-row'><small>↳ {a['label']}: <strong>{amts}</strong></small></td></tr>"
            return r

        # ETRException (Safe Harbour / No ETRComputation)
        exc_h = ""
        if j.get("etr_exception"):
            EXC_LABELS = _CODES["EXC_LABELS"]
            exc_rows = ""
            for ex in j["etr_exception"]:
                ex_label = EXC_LABELS.get(ex["type"], ex["type"])
                sg_txt = f'<small>{ex["subgroup"]}</small>' if ex.get("subgroup") else "—"
                exc_rows += (
                    f'<tr><td colspan="2" style="background:#FFF8E6;padding:6px 11px;font-size:0.72rem;font-weight:600;color:#B07800">{ex_label}</td></tr>' +
                    row("SubGroup", sg_txt) +
                    (row("Revenue CbCR", ex["revenue"]) if ex.get("revenue") and ex["revenue"] != "—" else "") +
                    (row("Profit CbCR", f'<strong>{ex["profit"]}</strong>') if ex.get("profit") and ex["profit"] != "—" else "") +
                    (row("IncomeTax CbCR", ex["income_tax"]) if ex.get("income_tax") and ex["income_tax"] != "—" else "")
                )
            if exc_rows:
                exc_h = card("ETR Status – Safe Harbour", tbl(exc_rows))

        ov_rows = ""
        if ov:
            ov_rows = (
                row("FANIL", ov["fanil"]) +
                row("AdjustedFANIL", ov["adj_fanil"]) +
                row("NetGlobeIncome", f'<strong>{ov["net_income"]}</strong>') +
                adj_rows(ov.get("net_income_adjs", [])) +
                row("IncomeTaxExpense", ov["income_tax"]) +
                row("AdjustedCoveredTax", f'<strong>{ov["adj_cov_tax"]}</strong>') +
                adj_rows(ov.get("adj_cov_adjs", [])) +
                row("Aggregate Current Tax", ov["aggr_curr"]) +
                row("ETR Rate", f'<strong class="etr-val">{ov["etr_rate"]}</strong>') +
                row("TopUpTax %", ov["top_up_pct"]) +
                row("ExcessProfits", ov["excess_profits"]) +
                row("TopUpTax", ov["top_up_tax"])
            )

        sub_h = ""
        if j["substance"]:
            s = j["substance"]
            sub_h = card("Substance-Based Income Exclusion (SBIE)", tbl(
                row("Totale SBIE", f'<strong>{s["total"]}</strong>') +
                row("PayrollCost", s["payroll"]) +
                row("PayrollMarkUp", s["payroll_mu"]) +
                row("TangibleAssetValue", s["tangible"]) +
                row("TangibleAssetMarkup", s["tangible_mu"])
            ))

        qd_h = ""
        if j["qdmtt"]:
            q = j["qdmtt"]
            qd_h = card("QDMTT", tbl(
                row("Amount", f'<strong>{q["amount"]}</strong>') +
                row("FAS", q["fas"]) + row("Currency", q["currency"]) +
                row("SBIE Available", q["sbie"]) + row("De Minimis Available", q["de_min"])
            ))

        en_h = ""
        if j["excess_neg"]:
            e = j["excess_neg"]
            en_h = card("Excess Negative Tax Expense Carry-Forward", tbl(
                row("Prior Year Balance", e["prior"]) +
                row("Generated in RFY", e["generated"]) +
                row("Utilized in RFY", e["utilized"]) +
                row("Remaining", f'<strong>{e["remaining"]}</strong>')
            ))

        ce_h = ""
        if j["ce_computations"]:
            cer = "".join(f"<tr><td><code>{c['tin']}</code></td><td>{c['fanil']}</td><td>{c['adj_fanil']}</td><td>{c['net_income']}</td><td>{c['adj_cov_tax']}</td><td>{c['defer_total']}</td></tr>" for c in j["ce_computations"])
            ce_h = card("CE Computations", tbl(cer, ["TIN", "FANIL", "AdjFANIL", "NetGlobeIncome", "AdjCoveredTax", "DeferTaxAdjust"]))

        el_h = ""
        if j["elections"]:
            el_h = card("Elezioni attive", "<ul class='elist'>" + "".join(f"<li>{e}</li>" for e in j["elections"]) + "</ul>")

        ltj_h = ""
        if j["ltj"]:
            ltj = j["ltj"]
            ltj_rows = row("Total TopUpTax", f'<strong>{ltj["total_tut"]}</strong>')
            for ltce in ltj["ltce"]:
                for iir in ltce["iir"]:
                    ltj_rows += f'<tr><td colspan="2" style="padding-top:6px"><strong>CE: {ltce["tin"]}</strong> — NetGlobeIncome: {iir["net_income"]} | TopUpTax: {iir["top_up"]}</td></tr>'
                    for pe in iir["parents"]:
                        ltj_rows += f'<tr><td class="lbl" style="padding-left:24px">→ Parent {pe["tin"]} ({pe["country"]})</td><td>Ratio: {pe["inc_ratio"]} | TuT Share: {pe["tut_share"]} | TopUpTax: {pe["top_up"]}</td></tr>'
            ltj_h = card("Low Tax Jurisdiction – IIR Allocation", tbl(ltj_rows))

        jur_flag = j['jurisdiction']
        etr_display = ov["etr_rate"] if ov else "—"
        etr_color = "#EAAA00"
        if ov:
            try:
                rate = float(ov["etr_rate"].replace("%","").replace(",","."))
                if rate < 15: etr_color = "#E05C5C"
                elif rate >= 15: etr_color = "#00B2A9"
            except: pass

        jcards += f'''
        <div class="jcard" id="jur-{jur_flag}">
          <div class="jhead">
            <span class="jflag">{jur_flag}</span>
            <span class="jcur">{cur}</span>
            <span class="jdt">{j["doc_type"]}</span>
            {'<span class="jjur">Tassazione: ' + ", ".join(j["jur_taxing"]) + '</span>' if j["jur_taxing"] else ""}
            <span class="jetr" style="color:{etr_color}">{etr_display}</span>
          </div>
          <div class="jbody">
            <div class="g2">
              <div>{exc_h}{card("Overall Computation", tbl(ov_rows)) if ov_rows else ""}{sub_h}</div>
              <div>{qd_h}{en_h}{ce_h}{el_h}{ltj_h}</div>
            </div>
          </div>
        </div>'''

    # ── UTPRAttribution ──
    utpr_html = ""
    for ua in d["utpr"]:
        if not ua["attributions"]: continue
        rows = "".join(f"<tr><td><strong>{a['country']}</strong></td><td>{a['carry_fwd']}</td><td>{a['employees']}</td><td>{a['tangible']}</td><td>{a['pct']}</td><td>{a['attributed']}</td><td>{a['add_cash']}</td><td>{a['carried_fwd']}</td></tr>" for a in ua["attributions"])
        utpr_html += tbl(rows, ["Paese", "TuT C/Fwd", "Employees", "TangibleAsset", "UTPR%", "TuT Attributed", "AddCashTax", "TuT C/Fwd Out"])
    if utpr_html:
        utpr_html = sec("UTPR Attribution", utpr_html, "utpr")

    # ── Nav ──
    nav_jur = "".join(f'<a href="#jur-{j["jurisdiction"]}" class="nj">{j["jurisdiction"]}</a>' for j in d["jurisdictions"])
    utpr_nav = '<a href="#utpr">UTPR Attribution</a>' if d["utpr"] else ""
    nav = f'''<nav id="sb">
      <div class="nl">PILLAR · GloBE</div>
      <div class="nf">{xml_name}</div>
      <div class="np">{fi["start"]} → {fi["end"]}</div>
      <div class="ns">Sezioni</div>
      <a href="#som">Sommario</a>
      <a href="#corp">Corporate Structure</a>
      <a href="#sum">Summary</a>
      <a href="#jurs">Jurisdiction Sections</a>
      {utpr_nav}
      <div class="ns">Giurisdizioni</div>
      {nav_jur}
      <div class="nm">{fi["mne_name"] or ""}</div>
    </nav>'''

    return f'''<!DOCTYPE html><html lang="it"><head><meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>PILLAR Report · {fi["mne_name"] or xml_name}</title>
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
:root{{--blue:#00338D;--teal:#00B2A9;--yellow:#EAAA00;--white:#fff;--gray:#F5F5F5;--border:#D9D9D9;--muted:#6B7280;--dark:#1A1A1A;--sb:250px;--sans:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;--mono:"SF Mono","Fira Mono",monospace}}
html{{scroll-behavior:smooth}}
body{{font-family:var(--sans);background:var(--gray);color:var(--dark);display:flex;min-height:100vh;font-size:14px}}
#sb{{width:var(--sb);min-width:var(--sb);background:var(--blue);color:#fff;padding:20px 0;position:fixed;top:0;left:0;height:100vh;overflow-y:auto;display:flex;flex-direction:column;gap:1px}}
.nl{{padding:0 18px 3px;font-size:0.9rem;font-weight:700;letter-spacing:0.06em;color:var(--yellow)}}
.nf{{padding:0 18px 2px;font-size:0.65rem;color:rgba(255,255,255,0.45);word-break:break-all}}
.np{{padding:0 18px 14px;font-size:0.65rem;color:rgba(255,255,255,0.35)}}
.ns{{padding:10px 18px 3px;font-size:0.6rem;font-weight:700;letter-spacing:0.12em;text-transform:uppercase;color:rgba(255,255,255,0.35)}}
.nm{{margin-top:auto;padding:14px 18px 0;font-size:0.65rem;color:rgba(255,255,255,0.25);border-top:1px solid rgba(255,255,255,0.08)}}
#sb a{{display:block;padding:6px 18px;color:rgba(255,255,255,0.7);text-decoration:none;font-size:0.8rem;border-left:3px solid transparent;transition:all 0.12s}}
#sb a:hover{{color:#fff;background:rgba(255,255,255,0.07);border-left-color:var(--yellow)}}
.nj{{font-family:var(--mono);font-size:0.75rem!important}}
#main{{margin-left:var(--sb);flex:1;padding:28px 36px;max-width:calc(100vw - var(--sb))}}
.rh{{background:var(--blue);color:#fff;border-radius:5px;padding:20px 24px;margin-bottom:24px;display:flex;align-items:flex-start;justify-content:space-between;position:relative;overflow:hidden}}
.rh::after{{content:"";position:absolute;bottom:0;left:0;right:0;height:3px;background:var(--yellow)}}
.rht{{font-size:1.3rem;font-weight:300;margin-bottom:3px}}.rht strong{{font-weight:700}}
.rhs{{font-size:0.78rem;color:rgba(255,255,255,0.6)}}
.rhb{{background:var(--yellow);color:var(--blue);font-size:0.68rem;font-weight:700;padding:3px 10px;border-radius:2px}}
section{{margin-bottom:28px}}
h2{{font-size:0.7rem;font-weight:700;color:var(--blue);text-transform:uppercase;letter-spacing:0.1em;margin-bottom:12px;padding-bottom:7px;border-bottom:2px solid var(--yellow)}}
h3{{font-size:0.7rem;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:0.08em;margin:12px 0 7px}}
.icard{{background:var(--white);border:1px solid var(--border);border-radius:4px;padding:14px 16px;margin-bottom:12px}}
table{{width:100%;border-collapse:collapse;font-size:0.8rem;background:var(--white);border:1px solid var(--border);border-radius:4px;overflow:hidden;margin-bottom:10px}}
thead tr{{background:var(--blue);color:#fff}}
th{{padding:8px 11px;text-align:left;font-size:0.68rem;font-weight:600;letter-spacing:0.04em}}
td{{padding:7px 11px;border-bottom:1px solid var(--border);vertical-align:top}}
tr:last-child td{{border-bottom:none}}
tr:hover td{{background:#FAFAFA}}
td.lbl{{color:var(--muted);font-size:0.75rem;width:200px;white-space:nowrap}}
td.adj-row{{background:#FAFCFF;padding-top:2px;padding-bottom:2px}}
code{{font-family:var(--mono);font-size:0.78rem;background:#F0F4FF;padding:1px 5px;border-radius:3px;color:#003}}
.g2{{display:grid;grid-template-columns:1fr 1fr;gap:16px}}
.bu{{background:var(--blue);color:#fff;font-size:0.62rem;font-weight:700;padding:2px 7px;border-radius:2px}}
.bc{{background:var(--teal);color:#fff;font-size:0.62rem;font-weight:700;padding:2px 7px;border-radius:2px}}
.etr-val{{color:var(--blue);font-size:1rem;font-weight:700}}
.badge-red{{background:#FCE8E8;color:#B00;font-size:0.72rem;font-weight:600;padding:2px 8px;border-radius:3px}}
.badge-yellow{{background:#FFF8E6;color:#B07800;font-size:0.72rem;font-weight:600;padding:2px 8px;border-radius:3px}}
.badge-green{{background:#E6F6F5;color:#006B65;font-size:0.72rem;font-weight:600;padding:2px 8px;border-radius:3px}}
.badge-gray{{background:#F0F0F0;color:#555;font-size:0.72rem;font-weight:600;padding:2px 8px;border-radius:3px}}
.badge-blue{{background:#E6EEF9;color:var(--blue);font-size:0.78rem;font-weight:600;padding:2px 10px;border-radius:3px}}
.jcard{{background:var(--white);border:1px solid var(--border);border-radius:5px;margin-bottom:18px;overflow:hidden}}
.jhead{{background:var(--blue);padding:12px 18px;display:flex;align-items:center;gap:14px;flex-wrap:wrap}}
.jflag{{font-family:var(--mono);font-size:1rem;font-weight:700;color:#fff;letter-spacing:0.1em}}
.jcur{{font-size:0.68rem;color:rgba(255,255,255,0.55);background:rgba(255,255,255,0.1);padding:2px 8px;border-radius:2px}}
.jdt{{font-size:0.68rem;color:rgba(255,255,255,0.45)}}
.jjur{{font-size:0.68rem;color:rgba(255,255,255,0.5)}}
.jetr{{margin-left:auto;font-size:1.05rem;font-weight:700}}
.jbody{{padding:16px}}
.elist{{padding-left:20px;font-size:0.8rem;color:var(--muted)}}
.elist li{{margin-bottom:3px}}
@media print{{#sb{{display:none}}#main{{margin-left:0}}.jcard{{page-break-inside:avoid}}}}
</style></head><body>
{nav}
<div id="main">
  <div class="rh">
    <div>
      <div class="rht">Report GloBE · <strong>{fi["mne_name"] or xml_name}</strong></div>
      <div class="rhs">{fi["start"]} → {fi["end"]} · {fi["currency"]} · {fi["fas"]}</div>
    </div>
    <div class="rhb">GIR26 · Entratel</div>
  </div>
  {sec("Sommario", html_som, "som")}
  {sec("Corporate Structure", html_corp, "corp")}
  {sec("Summary per Giurisdizione", html_sum, "sum")}
  <section id="jurs"><h2>Jurisdiction Sections</h2>{jcards}</section>
  {utpr_html}
</div>
</body></html>'''

# ── Entry point ───────────────────────────────────────────────────────────────
def generate_report(xml_path: Path, output_dir: Path | None = None) -> Path:
    xml_path = Path(xml_path)
    if output_dir is None: output_dir = xml_path.parent
    output_dir = Path(output_dir); output_dir.mkdir(parents=True, exist_ok=True)
    data = parse_globe(xml_path)
    html = _html(data, xml_path.name)
    out = output_dir / f"{xml_path.stem}_report.html"
    out.write_text(html, encoding="utf-8")
    print(f"  [report] Generato: {out.name} ({out.stat().st_size:,} byte)")
    return out
