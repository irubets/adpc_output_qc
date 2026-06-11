"""
adpc_output_qc_report.py
Generates qc_output_report.md and qc_output_report.html
from qc_output_results.json.

Usage:
    python adpc_output_qc_report.py --json qc_output_results.json
    python adpc_output_qc_report.py --json qc_output_results.json --out_dir ./reports
"""

import argparse
import json
import os
from datetime import datetime
from pathlib import Path


# ── Helpers ───────────────────────────────────────────────────────────────────

def load(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _n(x):
    try:
        return f"{int(x):,}"
    except (TypeError, ValueError):
        return str(x) if x is not None else "—"


def _flag(n, warn=1, error=None):
    """✅ if 0, ⚠️ if >= warn, 🔴 if >= error."""
    try:
        v = int(n)
        if error is not None and v >= error:
            return "🔴"
        if v >= warn:
            return "⚠️"
        return "✅"
    except (TypeError, ValueError):
        return "—"


def _list_str(items, max_show=10):
    if not items:
        return "none"
    shown = [str(x) for x in items[:max_show]]
    suffix = f" … (+{len(items) - max_show} more)" if len(items) > max_show else ""
    return ", ".join(shown) + suffix


def _tbl(rows, headers=None, max_rows=None):
    """Render list-of-dicts as a Markdown pipe table."""
    if not rows:
        return "*No records.*"
    if headers is None:
        headers = list(rows[0].keys())
    if max_rows:
        suffix_rows = rows[max_rows:]
        rows = rows[:max_rows]
    else:
        suffix_rows = []
    hdr  = "| " + " | ".join(str(h) for h in headers) + " |"
    sep  = "| " + " | ".join(["---"] * len(headers)) + " |"
    body = []
    for r in rows:
        if isinstance(r, dict):
            vals = [str(r.get(h, "")) for h in headers]
        else:
            vals = [str(v) for v in r]
        body.append("| " + " | ".join(vals) + " |")
    result = "\n".join([hdr, sep] + body)
    if suffix_rows:
        result += f"\n\n*… and {len(suffix_rows)} more rows. See XLSX for full listing.*"
    return result


# ── Report builder ────────────────────────────────────────────────────────────

def generate_md(data: dict) -> str:
    meta      = data.get("meta", {})
    sas_logic = data.get("sas_flag_logic", {})
    meta_d    = data.get("meta_detail",       {})
    subj      = data.get("subject_counts",    {})
    trt_map   = data.get("treatment_mapping", {})
    route_tbl = data.get("route_table",       {})
    crossref  = data.get("crit_nca_crossref", {})
    pops      = data.get("populations",       {})
    dtype_fl  = data.get("dtype_flags",       {})
    mrrlt_dup = data.get("mrrlt_duplicates",  {})
    err_rec   = data.get("erroneous_records", {})
    miss_dose = data.get("missing_dose",      {})
    prepost   = data.get("prepost_flags",     {})
    vomit     = data.get("vomiting",          {})
    prof_type = data.get("profile_types",     {})
    rich_list = data.get("richrfl_listrfl",   {})
    specs     = data.get("specs_comparison",  {})
    aprofile  = data.get("aprofile_alignment",{})
    # Gap checks
    ph3_excl  = data.get("phase3_exclusion",      {})
    ax_excl   = data.get("astx030_exclusion",     {})
    avalu_c   = data.get("avalu_consistency",     {})
    sort_key  = data.get("sort_key_uniqueness",   {})
    bsacat_d  = data.get("bsacat_derivation",     {})
    aperiodc_d= data.get("aperiodc_derivation",   {})
    aseq_d    = data.get("aseq_derivation",       {})
    cohortcd_d= data.get("cohortcd_derivation",   {})
    drugcat_d = data.get("drugcat_derivation",    {})
    trtint_d  = data.get("trtint_derivation",     {})
    nca40_c   = data.get("nca40xrs_consistency",  {})
    crit_ov   = data.get("crit_subject_overrides",{})
    patches   = data.get("manual_patches",        {})
    wide_sum  = data.get("wide_summaries",        {})
    aval_inc  = data.get("aval_increase_flags",   {})
    sas_log   = data.get("sas_log_analysis",      {})

    run_dt    = meta.get("generated", datetime.now().isoformat(timespec="seconds"))
    data_path = meta.get("data_path", "")
    sas_path  = meta.get("sas_path",  "")
    specs_path = meta.get("specs_path", "")

    L = []
    def w(*parts): L.append("".join(str(p) for p in parts))
    def nl(): L.append("")

    # ── Title ─────────────────────────────────────────────────────────────────
    w("# ADPC / ADNCA Output Data QC Report")
    nl()
    w("**Generated:** ", run_dt, "  ")
    w("**Dataset:** `", data_path, "`  ")
    w("**SAS program:** `", sas_path or "(not provided)", "`  ")
    w("**Specifications:** `", specs_path or "(not provided)", "`  ")
    w("**Phase filter:** ", meta.get("phase_filter", "ALL"), "  ")
    w(f"**Records:** {_n(meta.get('n_records'))}  "
      f"**Columns:** {_n(meta.get('n_columns'))}")
    nl()

    # ── User override misses — QC finding ────────────────────────────────────
    overrides = meta.get("user_override_misses", [])
    if overrides:
        w("> ⚠️ **Variable mapping QC finding:** The following variable names were provided "
          "by the user but were **not found in the dataset**. This may indicate the variable "
          "is missing from the dataset or was named differently.")
        nl()
        rows = [{"Canonical": o["canonical"], "User_Provided": o["provided"],
                 "Status": "❌ Not found in dataset"} for o in overrides]
        w(_tbl(rows)); nl()
    nl()
    w("---"); nl()

    # ══════════════════════════════════════════════════════════════════════════
    # Section 0: Flag Derivation Logic
    # ══════════════════════════════════════════════════════════════════════════
    w("## 0 — Flag Derivation Logic")
    nl()
    w("> **Important:** This section shows exactly which CRIT and NCAxxXRS variables ")
    w("> drive `PKSUMXFL` and `NCA01FL` in this program. Not all CRITs and not all ")
    w("> NCAxxXRS variables participate — some are informational only or reserved for ")
    w("> future subset analyses. If you apply this tool to a different project with ")
    w("> different logic, compare this table to your `whichc()` and `cmiss()` calls.")
    nl()

    parse_method = sas_logic.get("parse_method", "unknown")
    if parse_method == "whichc_cmiss":
        w(f"✅ **Parse method:** `whichc()` / `cmiss()` patterns found and parsed from SAS program.")
    else:
        w(f"⚠️ **Parse method:** `{parse_method}` — default lists used. "
          "Verify against your SAS program.")
    nl()

    # Parse warnings
    for pw in sas_logic.get("parse_warnings", []):
        w(f"> ⚠️ {pw}"); nl()

    # Derivation table
    logic_rows = [
        {
            "Derived Flag": "PKSUMXFL = 'Y'",
            "SAS Function": "whichc('Y', ...)",
            "Variables":    ", ".join(sas_logic.get("pksumxfl_crits", [])),
            "Explanation":  "Record is EXCLUDED from summary statistics if ANY of the listed CRITs has a value of Y.",
        },
        {
            "Derived Flag": "NCAXFL = 'Y'",
            "SAS Function": "cmiss(...) < N",
            "Variables":    ", ".join(sas_logic.get("ncaxfl_ncaxrs", [])),
            "Explanation":  "Record is flagged as having at least one NCA exclusion reason if ANY of the listed NCAxxXRS variables is non-empty (i.e., not ALL of them are missing).",
        },
        {
            "Derived Flag": "NCA01FL = 'Y'",
            "SAS Function": "compound",
            "Variables":    sas_logic.get("nca01fl_logic", "RICHRFL='Y' AND NCAXFL=' '"),
            "Explanation":  "Record is INCLUDED in NCA analysis if it belongs to a rich profile (RICHRFL='Y') AND none of the NCAxxXRS exclusion variables above are populated (NCAXFL=' ').",
        },
    ]
    w(_tbl(logic_rows)); nl()

    # Raw SAS statements (verbatim from program)
    if sas_logic.get("raw_pksumxfl_stmt"):
        w("**Raw SAS — PKSUMXFL derivation:**"); nl()
        w("```sas"); w(sas_logic["raw_pksumxfl_stmt"]); w("```"); nl()
    if sas_logic.get("raw_ncaxfl_stmt"):
        w("**Raw SAS — NCAXFL derivation:**"); nl()
        w("```sas"); w(sas_logic["raw_ncaxfl_stmt"]); w("```"); nl()

    # Informational-only flags
    all_crit  = set(meta_d.get("crit_fl_vars", []))
    all_ncaxrs = set(meta_d.get("nca_xrs_vars", []))
    driving_crits  = set(sas_logic.get("pksumxfl_crits", []))
    driving_ncaxrs = set(sas_logic.get("ncaxfl_ncaxrs", []))
    info_crits  = sorted(all_crit  - driving_crits)
    info_ncaxrs = sorted(all_ncaxrs - driving_ncaxrs)

    if info_crits:
        w(f"**CRITxFL in dataset but NOT driving PKSUMXFL** *(informational only)*:  ")
        w("`" + "`, `".join(info_crits) + "`"); nl()
    if info_ncaxrs:
        w(f"**NCAxxXRS in dataset but NOT driving NCAXFL** *(informational only)*:  ")
        w("`" + "`, `".join(info_ncaxrs) + "`"); nl()

    nl(); w("---"); nl()

    # ══════════════════════════════════════════════════════════════════════════
    # Section A: Dataset Overview
    # ══════════════════════════════════════════════════════════════════════════
    w("## A — Dataset Overview"); nl()

    not_mapped = meta_d.get("not_mapped", [])
    if not_mapped:
        w(f"⚠️ **Variables not mapped** (expected but not found or not provided): "
          f"`{'`, `'.join(not_mapped)}`"); nl()

    w(f"Total records: **{_n(meta.get('n_records'))}** | "
      f"Total columns: **{_n(meta.get('n_columns'))}**"); nl()

    # Subject counts
    w("### Subject and record counts"); nl()
    total_s = subj.get("total_subjects")
    if total_s is not None:
        w(f"Total subjects: **{_n(total_s)}**"); nl()

    if subj.get("by_phase"):
        w("**By PHASE:**"); nl(); w(_tbl(subj["by_phase"])); nl()
    if subj.get("by_actarm"):
        w("**By ACTARM:**"); nl(); w(_tbl(subj["by_actarm"])); nl()
    if subj.get("by_pctest"):
        w("**By PCTEST / PARAM:**"); nl(); w(_tbl(subj["by_pctest"])); nl()

    nl(); w("---"); nl()

    # ══════════════════════════════════════════════════════════════════════════
    # Section B: Treatment Mapping
    # ══════════════════════════════════════════════════════════════════════════
    w("## B — Treatment Mapping Cross-Reference"); nl()
    w("> Unique combinations of TREAT / SATRT / ATRT / AVISIT / VISITCD / "
      "PHASE / ACTARM / ACTARMCD. Use this table to verify that every visit × arm "
      "combination has been assigned the correct treatment labels."); nl()

    if trt_map.get("status"):
        w(f"_{trt_map['status']}_"); nl()
    else:
        w(f"**{_n(trt_map.get('n_combinations', 0))} unique combinations** "
          f"(variables found: {', '.join(trt_map.get('variables_found', []))})"); nl()
        w(_tbl(trt_map.get("table", []), max_rows=50)); nl()

    nl(); w("---"); nl()

    # ══════════════════════════════════════════════════════════════════════════
    # Section C: Route Table
    # ══════════════════════════════════════════════════════════════════════════
    w("## C — Route Table (EXROUTE / ECROUTE)"); nl()
    if route_tbl.get("status"):
        w(f"_{route_tbl['status']}_"); nl()
    else:
        w(_tbl(route_tbl.get("table", []))); nl()

    nl(); w("---"); nl()

    # ══════════════════════════════════════════════════════════════════════════
    # Section D: CRIT / NCAxxXRS Cross-Reference
    # ══════════════════════════════════════════════════════════════════════════
    w("## D — CRIT and NCAxxXRS Cross-Reference"); nl()
    w("> All CRIT and NCAxxXRS variables found in the dataset. "
      "`DRIVES_PKSUMXFL` and `DRIVES_NCAXFL` indicate whether a variable "
      "participates in deriving NCA01FL or PKSUMXFL (parsed from the SAS program). "
      "Variables marked 'no' are informational only for this program."); nl()

    cr = crossref.get("crossref", [])
    if cr:
        w(_tbl(cr, max_rows=60)); nl()
    else:
        w("*No CRIT or NCAxxXRS variables found in dataset.*"); nl()

    nl(); w("---"); nl()

    # ══════════════════════════════════════════════════════════════════════════
    # Section E: Population Definitions and Checks
    # ══════════════════════════════════════════════════════════════════════════
    w("## E — Population Definitions and Integrity Checks"); nl()

    w("### E1 — NCA Population (NCA01FL = 'Y')"); nl()
    w("> **Logic:** `RICHRFL = 'Y'` AND all of "
      "{" + ", ".join(sas_logic.get("ncaxfl_ncaxrs", [])) + "} are empty (NCAXFL = ' ')  "); nl()
    nca_pop = pops.get("nca_population", {})
    if nca_pop.get("status"):
        w(f"_{nca_pop['status']}_"); nl()
    else:
        w(f"Records: **{_n(nca_pop.get('n_records'))}** | "
          f"Subjects: **{_n(nca_pop.get('n_subjects'))}**"); nl()
        crit_y = nca_pop.get("crit_y_within_nca", {})
        if crit_y:
            w("⚠️ **CRITs = 'Y' within NCA population** *(intentional inclusions — verify each):*"); nl()
            rows = [{"CRITxFL": k, "N_records_Y": v} for k, v in crit_y.items()]
            w(_tbl(rows)); nl()
            w("*See sheet `NCA01FL_with_CRIT` in the XLSX for full record listing.*"); nl()
        else:
            w("✅ No CRITxFL = 'Y' within NCA population."); nl()
        ncaxrs_y = nca_pop.get("ncaxrs_within_nca", {})
        if ncaxrs_y:
            w("⚠️ **NCAxxXRS populated within NCA population** (these drive NCAXFL — should be empty):"); nl()
            rows = [{"NCAxxXRS": k, "N_records_populated": v} for k, v in ncaxrs_y.items()]
            w(_tbl(rows)); nl()
        else:
            w("✅ No NCAxxXRS driving variables populated within NCA population."); nl()
    nl()

    w("### E2 — Summary Statistics Population (PKSUMXFL missing)"); nl()
    w("> **Logic:** `PKSUMXFL` is missing or empty.  "); nl()
    w("> Includes NCA profile records AND trough concentrations.  "); nl()
    w("> `PKSUMXFL = 'Y'` is set when any of: "
      "{" + ", ".join(sas_logic.get("pksumxfl_crits", [])) + "} = 'Y'"); nl()

    sum_pop = pops.get("summary_population", {})
    if sum_pop.get("status"):
        w(f"_{sum_pop['status']}_"); nl()
    else:
        w(f"Records: **{_n(sum_pop.get('n_records'))}** | "
          f"Subjects: **{_n(sum_pop.get('n_subjects'))}**"); nl()
        crit_y_sum = sum_pop.get("crit_y_within_summary", {})
        if crit_y_sum:
            w("🔴 **CRITs = 'Y' within summary population** "
              "*(flagged records contributed to summary statistics — review required):*"); nl()
            rows = [{"CRITxFL": k, "N_records_Y": v} for k, v in crit_y_sum.items()]
            w(_tbl(rows)); nl()
            w("*See sheet `Summary_Pop_with_CRIT` in the XLSX.*"); nl()
        else:
            w("✅ No CRITxFL = 'Y' within summary statistics population."); nl()
    nl()

    w("### E3 — Trough Records (PKSUMXFL missing AND NCA01FL ≠ 'Y')"); nl()
    trough = pops.get("trough_records", {})
    w(f"Records: **{_n(trough.get('n_records', 0))}**  "); nl()
    w(f"*{trough.get('note', '')}*"); nl()
    if trough.get("by_visitcd"):
        w(_tbl(trough["by_visitcd"])); nl()

    w("### E4 — Population Contradiction Check"); nl()
    w("> `PKSUMXFL = 'Y'` AND `NCA01FL = 'Y'` simultaneously — must be zero."); nl()
    contra = pops.get("population_contradiction", {})
    n_contra = contra.get("n_records", 0)
    w(f"{_flag(n_contra, warn=1, error=1)} "
      f"Records: **{_n(n_contra)}**  "); nl()
    if n_contra > 0:
        w("🔴 **Critical error:** See sheet `Pop_Contradiction` in XLSX."); nl()
    nl()

    nl(); w("---"); nl()

    # ══════════════════════════════════════════════════════════════════════════
    # Section F: DTYPE Flags
    # ══════════════════════════════════════════════════════════════════════════
    w("## F — DTYPE Record Classification"); nl()
    # DTYPE in SAS program
    dtype_sas_note  = dtype_fl.get("dtype_sas_note", "")
    dtype_in_sas    = dtype_fl.get("dtype_in_sas")
    dtype_specs_note = dtype_fl.get("dtype_specs_note", "")
    dtype_in_specs  = dtype_fl.get("dtype_in_specs")

    if dtype_sas_note:
        icon = "✅" if dtype_in_sas else "⚠️"
        w(f"{icon} **SAS program:** {dtype_sas_note}"); nl()
    if dtype_specs_note:
        icon = "✅" if dtype_in_specs else "⚠️"
        w(f"{icon} **Specifications:** {dtype_specs_note}"); nl()
    nl()

    if dtype_fl.get("status") and not dtype_fl.get("distribution"):
        w(f"⚠️ _{dtype_fl['status']}_"); nl()
    else:
        w(_tbl(dtype_fl.get("distribution", []))); nl()
        n_del_nca = dtype_fl.get("deleted_in_nca_n", 0)
        w(f"{_flag(n_del_nca, warn=1, error=1)} "
          f"DELETED records with NCA01FL='Y': **{_n(n_del_nca)}**  "); nl()
        if n_del_nca > 0:
            w("🔴 **Critical:** DELETED records must never have NCA01FL='Y'. "
              "See sheet `DELETED_Records` in XLSX."); nl()
        else:
            w("✅ No DELETED records in NCA population."); nl()

    nl(); w("---"); nl()

    # ══════════════════════════════════════════════════════════════════════════
    # Section F2: Specifications Comparison
    # ══════════════════════════════════════════════════════════════════════════
    w("## F2 — Specifications Comparison"); nl()

    if specs.get("status"):
        w(f"_{specs['status']}_"); nl()
    else:
        w(f"**Specifications file:** `{Path(specs.get('specs_path','')).name}`  ")
        w(f"**Sheet used:** `{specs.get('specs_sheet', '—')}`  ")
        w(f"**Variables in specs:** {_n(specs.get('n_specs_vars', 0))}  ")
        w(f"**Variables in specs present in dataset:** "
          f"{_n(specs.get('n_in_dataset', 0))}  ")
        n_miss = specs.get("n_missing_from_dataset", 0)
        w(f"{_flag(n_miss)} "
          f"**Variables in specs MISSING from dataset:** **{_n(n_miss)}**"); nl()

        if n_miss == 0:
            w("✅ All variables listed in the specifications are present in the dataset."); nl()
        else:
            w("⚠️ The following variables are listed in the specifications but "
              "**not found in the dataset**. This may indicate they were omitted from "
              "the SAS program or the dataset was not regenerated after the specs were updated."); nl()
            miss_rows = specs.get("missing_from_dataset", [])
            w(_tbl(miss_rows, headers=["Name", "Label", "Core"])); nl()
            w("*See sheet `Specs_Comparison` in XLSX for the full variable list with in-dataset status.*"); nl()

        n_extra = specs.get("n_in_dataset_not_in_specs", 0)
        w(f"ℹ️ **Dataset columns not listed in specs** *(informational)*: "
          f"{_n(n_extra)}"); nl()
        if n_extra > 0 and specs.get("in_dataset_not_in_specs"):
            shown = specs["in_dataset_not_in_specs"][:30]
            suffix = f" … (+{n_extra-30} more)" if n_extra > 30 else ""
            w("`" + "`, `".join(shown) + "`" + suffix); nl()

    nl(); w("---"); nl()

    # ══════════════════════════════════════════════════════════════════════════
    # Section F3: APROFILE Alignment
    # ══════════════════════════════════════════════════════════════════════════
    w("## F3 — APROFILE / NCA01FL Alignment"); nl()
    w("> `APROFILE` is the true identifier of an individual PK profile: "
      "USUBJID × PCTEST × AVISIT. It is populated only for records that belong "
      "to a named NCA profile. It must align with `NCA01FL`:  "); nl()
    w("> - If `NCA01FL = 'Y'` → `APROFILE` must be populated  "); nl()
    w("> - If `APROFILE` is populated → `NCA01FL` should be 'Y'  "); nl()
    w("> Misalignment in either direction is a 🔴 QC finding."); nl()

    if aprofile.get("status"):
        icon = "⚠️" if "not found" in aprofile["status"].lower() else "ℹ️"
        w(f"{icon} _{aprofile['status']}_"); nl()
    else:
        w(f"  APROFILE column detected: `{aprofile.get('aprofile_col', '—')}`  ")
        w(f"  APROFILE populated: **{_n(aprofile.get('n_aprofile_populated'))}** records  ")
        w(f"  APROFILE blank: **{_n(aprofile.get('n_aprofile_blank'))}** records  ")
        w(f"  NCA01FL = 'Y': **{_n(aprofile.get('n_nca01fl_y'))}** records"); nl()

        n_c1 = aprofile.get("n_nca01_y_aprofile_blank", 0)
        n_c2 = aprofile.get("n_aprofile_nca01_not_y",   0)

        w(f"{_flag(n_c1, warn=1, error=1)} "
          f"**NCA01FL='Y' but APROFILE blank:** {_n(n_c1)} records  "); nl()
        if n_c1 > 0:
            w(f"🔴 {aprofile.get('case1_note', '')}  "); nl()
            w("*See sheet `APROFILE_NCA01_missing` in XLSX.*"); nl()

        w(f"{_flag(n_c2)} "
          f"**APROFILE populated but NCA01FL ≠ 'Y':** {_n(n_c2)} records  "); nl()
        if n_c2 > 0:
            w(f"⚠️ {aprofile.get('case2_note', '')}  "); nl()
            w("*See sheet `APROFILE_NCA01_mismatch` in XLSX.*"); nl()

        if n_c1 == 0 and n_c2 == 0:
            w("✅ APROFILE and NCA01FL are fully aligned."); nl()

        if aprofile.get("profile_summary"):
            w(f"**Profile summary** *(first 30 — see sheet `APROFILE_Summary` in XLSX for full list):*"); nl()
            w(_tbl(aprofile["profile_summary"][:30],
                   headers=["APROFILE", "N_records", "N_NCA01FL_Y", "All_NCA01FL_Y"])); nl()

    nl(); w("---"); nl()

    # ══════════════════════════════════════════════════════════════════════════
    # Section G: MRRLT Duplicate Timepoints
    # ══════════════════════════════════════════════════════════════════════════
    w("## G — MRRLT Duplicate Timepoints"); nl()
    w("> Duplicates are identified at the **MRRLT** level within "
      "USUBJID × PCTEST × VISITCD, because MRRLT is what the NCA software "
      "receives. NRRLT is shown for reference. After de-duplication in `adnca.sas` "
      "the NCA population should have zero duplicates."); nl()

    for pop_key, label in [
        ("all_records",       "All records"),
        ("nca_population",    "NCA population (NCA01FL='Y')"),
        ("summary_population","Summary population (PKSUMXFL missing)"),
    ]:
        dup = mrrlt_dup.get(pop_key, {})
        n_grp = dup.get("n_duplicate_groups", 0)
        n_rec = dup.get("n_duplicate_records", 0)
        expected_zero = pop_key == "nca_population"
        icon = _flag(n_grp, warn=1, error=1) if expected_zero else _flag(n_grp)
        w(f"**{label}:** {icon} "
          f"{_n(n_grp)} duplicate MRRLT groups | {_n(n_rec)} records involved"); nl()
        if n_rec > 0:
            w(f"*See sheet `MRRLT_Dups_{'_'.join(pop_key.split('_')[:2]).title()}` in XLSX.*"); nl()
    nl()

    nl(); w("---"); nl()

    # ══════════════════════════════════════════════════════════════════════════
    # Section H: Erroneous Records
    # ══════════════════════════════════════════════════════════════════════════
    w("## H — Erroneous Records"); nl()
    w("| Check | N | Note |")
    w("| --- | --- | --- |")
    for key, val in err_rec.items():
        if isinstance(val, dict):
            n   = val.get("n", 0)
            note = val.get("note", "")
            icon = _flag(n)
            w(f"| {key} | {icon} {_n(n)} | {note} |")
    nl()
    if any(isinstance(v, dict) and v.get("n", 0) > 0 for v in err_rec.values()):
        w("*See sheet `Erroneous_Records` in XLSX for full record listing.*"); nl()

    nl(); w("---"); nl()

    # ══════════════════════════════════════════════════════════════════════════
    # Section I: Dose Flags
    # ══════════════════════════════════════════════════════════════════════════
    w("## I — Dose Flags (CRIT3, CRIT7, Missed / Incomplete Dose)"); nl()
    if not miss_dose:
        w("*No dose flag variables found in dataset.*"); nl()
    else:
        w("| Flag | Label | N Records | N Subjects |")
        w("| --- | --- | --- | --- |")
        for flag, info in miss_dose.items():
            n_r = info.get("n_records", 0)
            n_s = info.get("n_subjects", 0)
            w(f"| {flag} | {info.get('label', '')} | "
              f"{_flag(n_r)} {_n(n_r)} | {_n(n_s)} |")
        nl()
        # List subjects for any flag with hits
        for flag, info in miss_dose.items():
            if info.get("subjects"):
                w(f"**{flag} subjects:** {_list_str(info['subjects'])}"); nl()

    nl(); w("---"); nl()

    # ══════════════════════════════════════════════════════════════════════════
    # Section J: Pre/Post Dose Flags
    # ══════════════════════════════════════════════════════════════════════════
    w("## J — Pre/Post Dose Timing Flags"); nl()
    if not prepost:
        w("*No pre/post dose flag variables found in dataset.*"); nl()
    else:
        w("| Flag | Label | N Records | N Subjects |")
        w("| --- | --- | --- | --- |")
        for flag, info in prepost.items():
            n_r = info.get("n_records", 0)
            n_s = info.get("n_subjects", 0)
            w(f"| {flag} | {info.get('label', '')} | "
              f"{_flag(n_r)} {_n(n_r)} | {_n(n_s)} |")
        nl()
        for flag, info in prepost.items():
            if info.get("subjects"):
                w(f"**{flag} subjects:** {_list_str(info['subjects'])}"); nl()

    nl(); w("---"); nl()

    # ══════════════════════════════════════════════════════════════════════════
    # Section K: Vomiting
    # ══════════════════════════════════════════════════════════════════════════
    w("## K — Vomiting (CRIT12, VOMFL)"); nl()
    if not vomit:
        w("*No vomiting flag variables found in dataset.*"); nl()
    else:
        for flag, info in vomit.items():
            n_r = info.get("n_records", 0)
            n_s = info.get("n_subjects", 0)
            w(f"**{flag}:** {_flag(n_r)} {_n(n_r)} records, {_n(n_s)} subjects"); nl()
            if info.get("by_visitcd"):
                w(_tbl(info["by_visitcd"])); nl()

    nl(); w("---"); nl()

    # ══════════════════════════════════════════════════════════════════════════
    # Section L: Profile Type Distribution
    # ══════════════════════════════════════════════════════════════════════════
    w("## L — Profile Type Distribution"); nl()
    if prof_type.get("status"):
        w(f"_{prof_type['status']}_"); nl()
    else:
        w("**Overall:**"); nl()
        w(_tbl(prof_type.get("overall_distribution", []))); nl()
        if prof_type.get("richrfl_distribution"):
            w("**RICHRFL distribution:**"); nl()
            w(_tbl(prof_type["richrfl_distribution"])); nl()
        if prof_type.get("by_group"):
            w("**By PHASE / ACTARM / PCTEST** *(first 40 rows — see XLSX for full table):*"); nl()
            w(_tbl(prof_type["by_group"], max_rows=40)); nl()

    nl(); w("---"); nl()

    # ══════════════════════════════════════════════════════════════════════════
    # Section M: RICHRFL / LISTRFL
    # ══════════════════════════════════════════════════════════════════════════
    w("## M — RICHRFL and LISTRFL Consistency"); nl()
    if rich_list.get("richrfl_dist"):
        w("**RICHRFL:**"); nl()
        rows = [{"Value": k, "N": v} for k, v in rich_list["richrfl_dist"].items()]
        w(_tbl(rows)); nl()
    if rich_list.get("listrfl_dist"):
        w("**LISTRFL:**"); nl()
        rows = [{"Value": k, "N": v} for k, v in rich_list["listrfl_dist"].items()]
        w(_tbl(rows)); nl()
    la = rich_list.get("listrfl_and_nca01fl_y", {})
    n_la = la.get("n_records", 0)
    w(f"{_flag(n_la, warn=1, error=1)} "
      f"LISTRFL='Y' AND NCA01FL='Y' simultaneously: **{_n(n_la)}**  "); nl()
    if n_la > 0:
        w("🔴 **Critical:** LISTRFL and NCA01FL must not co-occur."); nl()

    nl(); w("---"); nl()

    # ══════════════════════════════════════════════════════════════════════════
    # Section N — Gap Checks: Data Lineage & Derivation Verification
    # ══════════════════════════════════════════════════════════════════════════
    w("## N — Gap Checks: Data Lineage and Derivation Verification"); nl()
    w("> These checks verify exclusions, derived variables, subject-specific overrides,")
    w("> and manual patches that mirror specific SAS logic in adnca.sas."); nl()

    # N1: Phase 3 exclusion
    w("### N1 — PHASE 3 Exclusion (Gap 3)"); nl()
    if ph3_excl.get("status"):
        w(f"_{ph3_excl['status']}_"); nl()
    else:
        n_ph3 = ph3_excl.get("n_phase3_records", 0)
        w(f"{_flag(n_ph3, warn=1, error=1)} PHASE 3 records in dataset: **{_n(n_ph3)}**  "); nl()
        if n_ph3 > 0:
            w("🔴 **Critical:** PHASE 3 records must be excluded from adnca."); nl()
            w(_tbl(ph3_excl.get("records", []), max_rows=20)); nl()
        else:
            w("✅ No PHASE 3 records found — as expected."); nl()

    nl()

    # N2: ASTX030 exclusion
    w("### N2 — PCTESTCD = 'ASTX030' Exclusion (Gap 4)"); nl()
    if ax_excl.get("status"):
        w(f"_{ax_excl['status']}_"); nl()
    else:
        n_ax = ax_excl.get("n_astx030_records", 0)
        w(f"{_flag(n_ax, warn=1, error=1)} PCTESTCD='ASTX030' records: **{_n(n_ax)}**  "); nl()
        if n_ax > 0:
            w("🔴 **Critical:** PCTESTCD='ASTX030' (combination drug with no PCORRES) must be excluded."); nl()
            w(_tbl(ax_excl.get("records", []), max_rows=20)); nl()
        else:
            w("✅ No ASTX030 combination-drug records found — as expected."); nl()

    nl()

    # N3: AVALU consistency
    w("### N3 — AVALU Consistency (Gap 5)"); nl()
    if avalu_c.get("status"):
        w(f"_{avalu_c['status']}_"); nl()
    else:
        n_miss  = avalu_c.get("n_aval_present_avalu_blank", 0)
        n_extra = avalu_c.get("n_aval_missing_avalu_set",   0)
        w(f"{_flag(n_miss,  warn=1, error=1)} AVAL non-missing but AVALU blank: **{_n(n_miss)}**  "); nl()
        w(f"{_flag(n_extra, warn=1)}           AVAL missing but AVALU set:      **{_n(n_extra)}**  "); nl()
        if avalu_c.get("avalu_distribution"):
            w("**AVALU distinct values:**"); nl()
            w(_tbl(avalu_c["avalu_distribution"])); nl()

    nl()

    # N4: Sort key uniqueness
    w("### N4 — Final Sort Key Uniqueness (Gap 6)"); nl()
    if sort_key.get("status"):
        w(f"_{sort_key['status']}_"); nl()
    else:
        n_dup = sort_key.get("n_duplicate_groups", 0)
        w(f"Sort key: `{'` × `'.join(sort_key.get('sort_key_vars', []))}`  "); nl()
        w(f"{_flag(n_dup, warn=1, error=1)} Duplicate groups: **{_n(n_dup)}** "
          f"({_n(sort_key.get('n_duplicate_records', 0))} records)  "); nl()
        if n_dup > 0:
            w("🔴 **Critical:** Dataset is not unique on its declared sort key."); nl()
            w(_tbl(sort_key.get("records", []), max_rows=20)); nl()
        else:
            w("✅ Dataset is unique on the declared sort key."); nl()

    nl()

    # N5: BSACAT derivation
    w("### N5 — BSACAT Derivation (Gap 7)"); nl()
    if bsacat_d.get("status"):
        w(f"_{bsacat_d['status']}_"); nl()
    else:
        n_err = bsacat_d.get("n_derivation_errors", 0)
        w(f"Rule: `{bsacat_d.get('rule', '')}`  "); nl()
        w(f"{_flag(n_err, warn=1, error=1)} Derivation errors: **{_n(n_err)}**  "); nl()
        if bsacat_d.get("bsacat_distribution"):
            w("**BSACAT distribution:**"); nl()
            w(_tbl(bsacat_d["bsacat_distribution"])); nl()
        if n_err > 0:
            w(_tbl(bsacat_d.get("mismatch_records", []), max_rows=20)); nl()

    nl()

    # N6: APERIODC derivation
    w("### N6 — APERIODC Derivation (Gap 8)"); nl()
    if aperiodc_d.get("status"):
        w(f"_{aperiodc_d['status']}_"); nl()
    else:
        n_err = aperiodc_d.get("n_derivation_errors", 0)
        w(f"Rule: `{aperiodc_d.get('rule', '')}`  "); nl()
        w(f"{_flag(n_err, warn=1, error=1)} Derivation errors: **{_n(n_err)}**  "); nl()
        if aperiodc_d.get("aperiodc_distribution"):
            w("**APERIODC distribution:**"); nl()
            w(_tbl(aperiodc_d["aperiodc_distribution"])); nl()
        if n_err > 0:
            w(_tbl(aperiodc_d.get("mismatch_records", []), max_rows=20)); nl()

    nl()

    # N7: ASEQ derivation
    w("### N7 — ASEQ Derivation (Gap 9)"); nl()
    if aseq_d.get("status"):
        w(f"_{aseq_d['status']}_"); nl()
    else:
        n_err = aseq_d.get("n_derivation_errors", 0)
        w(f"Rule: `{aseq_d.get('rule', '')}`  "); nl()
        w(f"{_flag(n_err, warn=1, error=1)} Derivation errors: **{_n(n_err)}**  "); nl()
        if aseq_d.get("sequence_aseq_crosstab"):
            w("**SEQUENCE × ASEQ crosstab:**"); nl()
            w(_tbl(aseq_d["sequence_aseq_crosstab"])); nl()

    nl()

    # N8: COHORTCD derivation
    w("### N8 — COHORTCD Derivation (Gap 10)"); nl()
    if cohortcd_d.get("status"):
        w(f"_{cohortcd_d['status']}_"); nl()
    else:
        n_miss = cohortcd_d.get("n_phase1_cohortcd_missing", 0)
        w(f"{_flag(n_miss, warn=1, error=1)} Phase 1 records with missing COHORTCD: **{_n(n_miss)}**  "); nl()
        if cohortcd_d.get("cohortcd_distribution"):
            w("**COHORTCD distribution:**"); nl()
            w(_tbl(cohortcd_d["cohortcd_distribution"])); nl()
        if cohortcd_d.get("cohort_cohortcd_crosstab"):
            w("**COHORT → COHORTCD mapping (first 30 rows):**"); nl()
            w(_tbl(cohortcd_d["cohort_cohortcd_crosstab"], max_rows=30)); nl()

    nl()

    # N9: DRUGCAT derivation
    w("### N9 — DRUGCAT Derivation (Gap 11)"); nl()
    if drugcat_d.get("status"):
        w(f"_{drugcat_d['status']}_"); nl()
    else:
        n_ph2  = drugcat_d.get("n_phase2_drugcat_set",        0)
        n_bad  = drugcat_d.get("n_phase1_unexpected_drugcat", 0)
        w(f"Rule: `{drugcat_d.get('rule', '')}`  "); nl()
        w(f"{_flag(n_ph2, warn=1, error=1)} Phase 2 records with DRUGCAT set (should be blank): **{_n(n_ph2)}**  "); nl()
        w(f"{_flag(n_bad, warn=1, error=1)} Phase 1 records with unexpected DRUGCAT value:       **{_n(n_bad)}**  "); nl()
        if drugcat_d.get("phase_cohortcd_drugcat_crosstab"):
            w("**PHASE × COHORTCD × DRUGCAT cross-reference:**"); nl()
            w(_tbl(drugcat_d["phase_cohortcd_drugcat_crosstab"], max_rows=30)); nl()

    nl()

    # N10: TRTINT derivation
    w("### N10 — TRTINT Derivation (Gap 12)"); nl()
    if trtint_d.get("status"):
        w(f"_{trtint_d['status']}_"); nl()
    else:
        n_miss24 = trtint_d.get("n_should_be_24_but_not", 0)
        n_unexp  = trtint_d.get("n_unexpected_24",         0)
        w(f"Rule: `{trtint_d.get('rule', '')}`  "); nl()
        w(f"{_flag(n_miss24, warn=1, error=1)} Records that should be TRTINT=24 but are not: **{_n(n_miss24)}**  "); nl()
        w(f"{_flag(n_unexp,  warn=1)}           Records with unexpected TRTINT=24:             **{_n(n_unexp)}**  "); nl()
        if trtint_d.get("trtint_distribution"):
            w("**TRTINT distribution:**"); nl()
            w(_tbl(trtint_d["trtint_distribution"])); nl()
        w("**Expected TRTINT=24 combinations:**"); nl()
        w(_tbl(trtint_d.get("expected_24h_combinations", []))); nl()

    nl()

    # N11: NCA40XRS consistency
    w("### N11 — NCA40XRS Combined Datetime Flag (Gap 13)"); nl()
    if nca40_c.get("status"):
        w(f"_{nca40_c['status']}_"); nl()
    else:
        n_miss  = nca40_c.get("n_miss_fire",  0)
        n_false = nca40_c.get("n_false_fire", 0)
        w(f"Drivers found: `{'`, `'.join(nca40_c.get('drivers_found', []))}`  "); nl()
        w(f"{_flag(n_miss,  warn=1, error=1)} Miss-fires (driver='Y' but NCA40XRS blank): **{_n(n_miss)}**  "); nl()
        w(f"{_flag(n_false, warn=1)}           False-fires (NCA40XRS set but no driver):   **{_n(n_false)}**  "); nl()
        if n_miss > 0:
            w(_tbl(nca40_c.get("miss_fire_records", []), max_rows=20)); nl()

    nl()

    # N12: CRIT subject-specific overrides
    w("### N12 — Subject-Specific CRIT Flag Overrides (Gap 14)"); nl()
    if not crit_ov:
        w("*No override check data available.*"); nl()
    else:
        # CRIT2FL overrides
        crit2_ov = crit_ov.get("crit2fl_overrides", {})
        if isinstance(crit2_ov, list):
            n_bad_ov = sum(1 for r in crit2_ov if "MISMATCH" in str(r.get("Override_OK", "")))
            w(f"**CRIT2FL overrides** (must be 'N' for 7 subject/visit combinations):  "); nl()
            w(f"{_flag(n_bad_ov, warn=1, error=1)} Mismatch count: **{_n(n_bad_ov)}**  "); nl()
            w(_tbl(crit2_ov)); nl()
        # CRIT19FL
        c19 = crit_ov.get("crit19fl", {})
        if not c19.get("status"):
            n_unexp19 = c19.get("n_unexpected_y", 0)
            w(f"**CRIT19FL** (PPI inhibitor — should only be 'Y' for subject 101-053):  "); nl()
            w(f"{_flag(n_unexp19, warn=1, error=1)} Unexpected CRIT19FL='Y': **{_n(n_unexp19)}**  "); nl()
        # CRIT20FL
        c20 = crit_ov.get("crit20fl", {})
        if not c20.get("status"):
            n_unexp20 = c20.get("n_unexpected_y",   0)
            n_miss20  = len(c20.get("expected_missing", []))
            w(f"**CRIT20FL** (Atypical profile — 5 named subject/visit combinations):  "); nl()
            w(f"{_flag(n_unexp20, warn=1, error=1)} Unexpected CRIT20FL='Y': **{_n(n_unexp20)}**  "); nl()
            w(f"{_flag(n_miss20,  warn=1)}           Expected but not found:  **{_n(n_miss20)}**  "); nl()
            if c20.get("expected_missing"):
                w(_tbl(c20["expected_missing"])); nl()

    nl()

    # N13: Manual patches
    w("### N13 — Manual Patch Verification (Gap 15)"); nl()
    patch_list = patches.get("patches", [])
    if not patch_list:
        w("*No manual patch data available.*"); nl()
    else:
        for p in patch_list:
            n_errs = sum(
                v for k, v in p.items()
                if k.startswith("n_") and k != "n_records" and isinstance(v, int)
            )
            w(f"**Patch {p.get('patch')} — {p.get('subject')} / {p.get('visitcd')}:** "
              f"{p.get('description')}  "); nl()
            w(f"  Records found: **{_n(p.get('n_records', 0))}** | "
              f"{_flag(n_errs, warn=1, error=1)} Patch errors: **{_n(n_errs)}**  "); nl()
            if p.get("n_records", 0) == 0:
                w("  ⚠️ Subject/visit not found in dataset — verify SUBJID format."); nl()
    if patches.get("note"):
        w(f"> ℹ️ {patches['note']}"); nl()

    nl()

    # N14: Wide summaries
    w("### N14 — Profile-Level Wide Summaries (Gap 16)"); nl()
    tp_r = wide_sum.get("timepoints_per_profile", {})
    if isinstance(tp_r, dict) and not tp_r.get("status"):
        n_sparse = tp_r.get("n_profiles_sparse", 0)
        n_total  = tp_r.get("n_profiles_total",  0)
        w(f"Total profiles: **{_n(n_total)}**  "); nl()
        w(f"{_flag(n_sparse, warn=1)} Sparse profiles (≤ 3 timepoints): **{_n(n_sparse)}**  "); nl()
        if n_sparse > 0:
            w("*See sheet `Sparse_Profiles` in XLSX for full listing.*"); nl()
            w(_tbl(tp_r.get("sparse_profiles", []), max_rows=20)); nl()
    sv_r = wide_sum.get("subject_per_visit", {})
    if isinstance(sv_r, list) and sv_r:
        w(f"Subject-per-visit table: **{_n(len(sv_r))}** combinations. "
          "See sheet `Subjects_Per_Visit` in XLSX."); nl()
    elif isinstance(sv_r, dict) and sv_r.get("status"):
        w(f"_{sv_r['status']}_"); nl()

    nl()

    # N15: AVAL increase flags
    w("### N15 — AVAL Increase Flag Verification — CRIT16/17 (Gap 17)"); nl()
    if aval_inc.get("status"):
        w(f"_{aval_inc['status']}_"); nl()
    else:
        for lbl, tp in [("CRIT16FL", 8), ("CRIT17FL", 24)]:
            info = aval_inc.get(lbl, {})
            if not info:
                w(f"_{lbl} not found in dataset._"); nl(); continue
            n_profiles = info.get("n_profiles_with_increase", 0)
            n_missed   = info.get("n_increase_not_flagged",   0)
            n_false    = info.get("n_flagged_but_no_increase", 0)
            w(f"**{lbl} (increase at {tp}h):**  "); nl()
            w(f"  Profiles with detected increase: **{_n(n_profiles)}**  "); nl()
            w(f"  {_flag(n_missed, warn=1, error=1)} Increases not flagged:   **{_n(n_missed)}**  "); nl()
            w(f"  {_flag(n_false,  warn=1)}           Flagged but no increase: **{_n(n_false)}**  "); nl()

    nl(); w("---"); nl()

    # ══════════════════════════════════════════════════════════════════════════
    # Section O: SAS Log Analysis
    # ══════════════════════════════════════════════════════════════════════════
    w("## O — SAS Log Analysis"); nl()

    if not sas_log:
        w("> ℹ️ No SAS log was provided. Re-run with `--log path/to/adnca.log` "
          "or enter the log path at the Step 2.7 prompt to include log analysis."); nl()
    elif not sas_log.get("read_ok", True):
        w(f"🔴 **Could not read log file:** `{sas_log.get('log_path', '')}`  "); nl()
        w(f"Error: `{sas_log.get('read_error', '')}`"); nl()
    else:
        log_path_disp = sas_log.get("log_path", "")
        n_err     = sas_log.get("n_errors",          0)
        n_warn    = sas_log.get("n_warnings",         0)
        n_err_u   = sas_log.get("n_errors_unique",    0)
        n_warn_u  = sas_log.get("n_warnings_unique",  0)
        n_warn_rv = sas_log.get("n_warnings_review",  0)
        n_warn_bn = sas_log.get("n_warnings_benign",  0)
        errors    = sas_log.get("errors",   [])
        warnings  = sas_log.get("warnings", [])

        w(f"> **Log file:** `{log_path_disp}`  "); nl()
        w(f"> Scan mirrors `%check_log` in `check_log.sas`: lines matching "
          "`/^\\s*ERROR:/i` or `/^\\s*WARNING:/i` are captured. "
          "Identical messages repeated on multiple lines are deduplicated."); nl()
        nl()

        # ── Headline counts ───────────────────────────────────────────────────
        err_icon  = "🔴" if n_err  else "✅"
        warn_icon = "⚠️" if n_warn_rv else ("ℹ️" if n_warn_bn else "✅")
        w(f"| Metric | Count |")
        w(f"| --- | --- |")
        w(f"| {err_icon} ERROR lines (raw)               | **{_n(n_err)}**   |")
        w(f"| {err_icon} ERROR lines (unique messages)   | **{_n(n_err_u)}** |")
        w(f"| {warn_icon} WARNING lines (raw)            | **{_n(n_warn)}**  |")
        w(f"| {warn_icon} WARNING lines (unique messages)| **{_n(n_warn_u)}**|")
        w(f"| ⚠️ WARNINGs requiring review              | **{_n(n_warn_rv)}**|")
        w(f"| ℹ️ Known-benign WARNINGs                  | **{_n(n_warn_bn)}**|")
        nl()

        # ── ERRORs ────────────────────────────────────────────────────────────
        w("### O1 — ERRORs"); nl()
        if not errors:
            w("✅ No ERROR lines found in the log."); nl()
        else:
            w(f"🔴 **{_n(n_err_u)} unique ERROR message(s)** "
              f"({_n(n_err)} total occurrences):"); nl()
            err_rows = []
            for e in errors:
                also = e.get("also_at_lines", [])
                also_str = (", ".join(str(x) for x in also[:10]) +
                            (f" … (+{len(also)-10} more)" if len(also) > 10 else "")
                            ) if also else ""
                err_rows.append({
                    "Line":      _n(e["line_no"]),
                    "Also at":   also_str or "—",
                    "ERROR text": e["text"],
                })
            w(_tbl(err_rows, headers=["Line", "Also at", "ERROR text"])); nl()

        # ── WARNINGs requiring review ─────────────────────────────────────────
        w("### O2 — WARNINGs Requiring Review"); nl()
        warn_review = [w2 for w2 in warnings if w2.get("category") == "review"]
        if not warn_review:
            w("✅ No WARNINGs requiring review."); nl()
        else:
            w(f"⚠️ **{_n(len(warn_review))} WARNING(s) to review:**"); nl()
            rv_rows = []
            for w2 in warn_review:
                also = w2.get("also_at_lines", [])
                also_str = (", ".join(str(x) for x in also[:10]) +
                            (f" … (+{len(also)-10} more)" if len(also) > 10 else "")
                            ) if also else ""
                rv_rows.append({
                    "Line":         _n(w2["line_no"]),
                    "Also at":      also_str or "—",
                    "WARNING text": w2["text"],
                })
            w(_tbl(rv_rows, headers=["Line", "Also at", "WARNING text"])); nl()

        # ── Known-benign WARNINGs ─────────────────────────────────────────────
        w("### O3 — Known-Benign WARNINGs"); nl()
        warn_benign = [w2 for w2 in warnings if w2.get("category") == "known-benign"]
        if not warn_benign:
            w("✅ No known-benign WARNINGs."); nl()
        else:
            w(f"ℹ️ **{_n(len(warn_benign))} known-benign WARNING(s)** "
              "(uninitialized variables, format not found, etc.):"); nl()
            bn_rows = []
            for w2 in warn_benign:
                also = w2.get("also_at_lines", [])
                also_str = (", ".join(str(x) for x in also[:10]) +
                            (f" … (+{len(also)-10} more)" if len(also) > 10 else "")
                            ) if also else ""
                bn_rows.append({
                    "Line":         _n(w2["line_no"]),
                    "Also at":      also_str or "—",
                    "WARNING text": w2["text"],
                })
            w(_tbl(bn_rows, headers=["Line", "Also at", "WARNING text"], max_rows=30)); nl()

        if sas_log.get("note"):
            w(f"> ℹ️ {sas_log['note']}"); nl()

    nl(); w("---"); nl()

    # ══════════════════════════════════════════════════════════════════════════
    # Summary — Action Items
    # ══════════════════════════════════════════════════════════════════════════
    w("## Summary — Action Items"); nl()
    w("| Priority | Check | Status | N |")
    w("| --- | --- | --- | --- |")

    def _item(priority, label, n, error=None):
        icon = _flag(n, warn=1, error=error)
        w(f"| {priority} | {label} | {icon} | {_n(n)} |")

    _item("🔴 Critical", "Population contradiction (PKSUMXFL='Y' AND NCA01FL='Y')",
          pops.get("population_contradiction", {}).get("n_records", 0), error=1)
    _item("🔴 Critical", "DELETED records with NCA01FL='Y'",
          dtype_fl.get("deleted_in_nca_n", 0), error=1)
    _item("🔴 Critical", "LISTRFL='Y' AND NCA01FL='Y'",
          rich_list.get("listrfl_and_nca01fl_y", {}).get("n_records", 0), error=1)
    _item("🔴 Critical", "MRRLT duplicates within NCA population (NCA01FL='Y')",
          mrrlt_dup.get("nca_population", {}).get("n_duplicate_groups", 0), error=1)
    _item("🔴 Critical", "NCA01FL='Y' but APROFILE blank",
          aprofile.get("n_nca01_y_aprofile_blank", 0), error=1)
    _item("⚠️ Review",  "Specs variables missing from dataset",
          specs.get("n_missing_from_dataset", 0))
    _item("⚠️ Review",  "APROFILE populated but NCA01FL ≠ 'Y'",
          aprofile.get("n_aprofile_nca01_not_y", 0))
    _item("⚠️ Review",  "CRITs='Y' within NCA population (intentional inclusions)",
          sum(pops.get("nca_population", {}).get("crit_y_within_nca", {}).values()) if isinstance(
              pops.get("nca_population", {}).get("crit_y_within_nca"), dict) else 0)
    _item("⚠️ Review",  "CRITs='Y' within summary population",
          sum(pops.get("summary_population", {}).get("crit_y_within_summary", {}).values()) if isinstance(
              pops.get("summary_population", {}).get("crit_y_within_summary"), dict) else 0)
    _item("⚠️ Review",  "Erroneous datetime records (bad year / PCDTCEFL)",
          sum(v.get("n", 0) for v in err_rec.values() if isinstance(v, dict)))
    _item("⚠️ Review",  "Missed dose records (CRIT3FL='Y')",
          miss_dose.get("CRIT3FL", {}).get("n_records", 0))
    _item("⚠️ Review",  "Pre-dose taken post-dose (CRIT10FL='Y')",
          prepost.get("CRIT10FL", {}).get("n_records", 0))
    _item("⚠️ Review",  "Post-dose taken pre-dose (CRIT2FL='Y')",
          prepost.get("CRIT2FL", {}).get("n_records", 0))
    _item("⚠️ Review",  "Vomiting (CRIT12FL='Y')",
          vomit.get("CRIT12FL", {}).get("n_records", 0))
    _item("ℹ️ Info",    "Trough records (in summary, not in NCA)",
          pops.get("trough_records", {}).get("n_records", 0))
    _item("ℹ️ Info",    "MRRLT duplicates across all records",
          mrrlt_dup.get("all_records", {}).get("n_duplicate_groups", 0))
    # ── Gap checks ───────────────────────────────────────────────────────────
    w("| **— Gap Checks —** | | | |")
    _item("🔴 Critical", "PHASE 3 records in dataset (must be 0)",
          ph3_excl.get("n_phase3_records", 0), error=1)
    _item("🔴 Critical", "PCTESTCD='ASTX030' records in dataset (must be 0)",
          ax_excl.get("n_astx030_records", 0), error=1)
    _item("🔴 Critical", "Sort key duplicate groups (NRRLT-based)",
          sort_key.get("n_duplicate_groups", 0), error=1)
    _item("🔴 Critical", "AVAL non-missing but AVALU blank",
          avalu_c.get("n_aval_present_avalu_blank", 0), error=1)
    _item("🔴 Critical", "NCA40XRS miss-fires (driver='Y' but flag blank)",
          nca40_c.get("n_miss_fire", 0), error=1)
    _item("🔴 Critical", "CRIT2FL override mismatches",
          sum(1 for r in crit_ov.get("crit2fl_overrides", [])
              if isinstance(r, dict) and "MISMATCH" in str(r.get("Override_OK", ""))), error=1)
    _item("⚠️ Review",  "BSACAT derivation errors",
          bsacat_d.get("n_derivation_errors", 0))
    _item("⚠️ Review",  "APERIODC derivation errors",
          aperiodc_d.get("n_derivation_errors", 0))
    _item("⚠️ Review",  "ASEQ derivation errors",
          aseq_d.get("n_derivation_errors", 0))
    _item("⚠️ Review",  "Phase 1 records with missing COHORTCD",
          cohortcd_d.get("n_phase1_cohortcd_missing", 0))
    _item("⚠️ Review",  "Phase 2 DRUGCAT set (should be blank)",
          drugcat_d.get("n_phase2_drugcat_set", 0))
    _item("⚠️ Review",  "TRTINT=24 expected but not set",
          trtint_d.get("n_should_be_24_but_not", 0))
    _item("⚠️ Review",  "Unexpected TRTINT=24",
          trtint_d.get("n_unexpected_24", 0))
    _item("⚠️ Review",  "CRIT19FL='Y' outside subject 101-053",
          crit_ov.get("crit19fl", {}).get("n_unexpected_y", 0))
    _item("⚠️ Review",  "CRIT20FL='Y' outside expected 5 combinations",
          crit_ov.get("crit20fl", {}).get("n_unexpected_y", 0))
    _item("⚠️ Review",  "CRIT16FL increases not flagged (8h)",
          aval_inc.get("CRIT16FL", {}).get("n_increase_not_flagged", 0))
    _item("⚠️ Review",  "CRIT17FL increases not flagged (24h)",
          aval_inc.get("CRIT17FL", {}).get("n_increase_not_flagged", 0))
    _item("ℹ️ Info",    "Sparse profiles (≤ 3 timepoints)",
          wide_sum.get("timepoints_per_profile", {}).get("n_profiles_sparse", 0))
    # ── SAS log ──────────────────────────────────────────────────────────────
    w("| **— SAS Log —** | | | |")
    _item("🔴 Critical", "SAS log ERRORs",
          sas_log.get("n_errors_unique", 0) if sas_log else 0, error=1)
    _item("⚠️ Review",  "SAS log WARNINGs requiring review",
          sas_log.get("n_warnings_review", 0) if sas_log else 0)
    _item("ℹ️ Info",    "SAS log known-benign WARNINGs",
          sas_log.get("n_warnings_benign", 0) if sas_log else 0)

    nl()

    if sas_logic.get("parse_warnings"):
        w("### SAS Parse Warnings"); nl()
        for pw in sas_logic["parse_warnings"]:
            w(f"- ⚠️ {pw}"); nl()
        nl()

    w(f"*Report generated by adpc_output_qc_report.py | {run_dt}*")
    return "\n".join(L)


# ── HTML wrapper ──────────────────────────────────────────────────────────────

def md_to_html(md_text: str, title: str = "ADPC Output QC") -> str:
    try:
        import markdown as md_lib
        body = md_lib.markdown(md_text, extensions=["tables", "nl2br", "fenced_code"])
    except ImportError:
        body = "<pre>" + md_text.replace("&", "&amp;").replace("<", "&lt;") + "</pre>"

    css = (
        "body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;"
        "font-size:15px;line-height:1.7;color:#1a1a1a;max-width:1100px;margin:40px auto;padding:0 24px}"
        "h1{font-size:1.7em;border-bottom:3px solid #1a252f;padding-bottom:.3em;color:#1a252f}"
        "h2{font-size:1.25em;border-bottom:2px solid #2980b9;padding-bottom:.2em;"
        "   margin-top:2.5em;color:#2980b9;page-break-before:always}"
        "h3{font-size:1.05em;margin-top:1.5em;color:#333}"
        "table{border-collapse:collapse;width:100%;margin:1em 0;font-size:13px}"
        "th{background:#1a252f;color:#fff;text-align:left;padding:7px 12px;border:1px solid #111}"
        "td{padding:6px 12px;border:1px solid #ddd;vertical-align:top}"
        "tr:nth-child(even) td{background:#f8f9fa}"
        "code,pre{background:#f0f0f0;font-family:'SFMono-Regular',Consolas,monospace;font-size:.87em}"
        "pre{padding:12px;border-radius:4px;overflow-x:auto;border-left:4px solid #2980b9}"
        "hr{border:none;border-top:2px solid #e0e0e0;margin:2.5em 0}"
        "blockquote{border-left:4px solid #f0a000;margin:1em 0;padding:.5em 1em;"
        "           background:#fffbf0;color:#5a3e00;border-radius:0 4px 4px 0}"
        "@media print{h2{page-break-before:always}body{max-width:100%;margin:0;padding:0 1cm}}"
    )
    return (
        f'<!DOCTYPE html>\n<html lang="en">\n<head>\n<meta charset="UTF-8">\n'
        f'<title>{title}</title>\n<style>{css}</style>\n</head>\n<body>\n'
        + body
        + "\n</body>\n</html>"
    )


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="ADPC Output QC Report Generator")
    parser.add_argument("--json",    required=True, help="Path to qc_output_results.json")
    parser.add_argument("--out_dir", default=None,  help="Output directory (default: same as JSON)")
    args = parser.parse_args()

    data     = load(args.json)
    md_text  = generate_md(data)

    out_dir  = args.out_dir or os.path.dirname(os.path.abspath(args.json))
    md_path  = os.path.join(out_dir, "qc_output_report.md")
    html_path = os.path.join(out_dir, "qc_output_report.html")

    with open(md_path,   "w", encoding="utf-8") as f:
        f.write(md_text)
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(md_to_html(md_text))

    print(f"[adpc_output_qc_report] Markdown : {md_path}")
    print(f"[adpc_output_qc_report] HTML     : {html_path}")
