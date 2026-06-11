"""
modules/checks_gap.py
Gap QC checks identified against adnca.sas logic (Gaps 1–17).

Each function accepts (df, mapping) and returns a dict stored in
qc_output_results.json under its section key.
"""

import pandas as pd

from .utils import col, col_list, is_blank


def check_phase3_exclusion(df, mapping):
    """Gap 3 — PHASE 3 records must be absent from the final dataset."""
    ph = col(df, "PHASE", mapping)
    if not ph:
        return {"status": "PHASE variable not found in dataset."}
    mask = df[ph].astype(str).str.strip().str.upper() == "PHASE 3"
    uid  = col(df, "USUBJID", mapping)
    test = col(df, "PCTEST",  mapping)
    keep = [c for c in [uid, ph, test] if c]
    return {
        "n_phase3_records": int(mask.sum()),
        "records": df[mask][keep].head(50).to_dict("records"),
        "note": "PHASE 3 records must be excluded from adnca. Any non-zero count is a critical error.",
    }


def check_astx030_exclusion(df, mapping):
    """Gap 4 — PCTESTCD/PARAMCD = 'ASTX030' must be absent (combination drug, no PCORRES)."""
    c = next((c for c in ["PCTESTCD", "PARAMCD", "TESTCD"] if c in df.columns), None)
    if not c:
        return {"status": "PCTESTCD / PARAMCD not found in dataset."}
    mask = df[c].astype(str).str.strip().str.upper() == "ASTX030"
    uid  = col(df, "USUBJID", mapping)
    aval = col(df, "AVAL",    mapping)
    keep = [x for x in [uid, c, aval] if x]
    return {
        "testcd_col":        c,
        "n_astx030_records": int(mask.sum()),
        "records":           df[mask][keep].head(50).to_dict("records"),
        "note": (
            "PCTESTCD='ASTX030' is the combination drug entry with no PCORRES. "
            "It must be excluded from adnca. Any non-zero count is a critical error."
        ),
    }


def check_avalu_consistency(df, mapping):
    """Gap 5 — AVALU must be populated whenever AVAL is non-missing."""
    aval  = col(df, "AVAL",    mapping)
    avalu = col(df, "AVALU",   mapping)
    uid   = col(df, "USUBJID", mapping)
    test  = col(df, "PCTEST",  mapping)
    visit = col(df, "VISITCD", mapping) or col(df, "AVISIT", mapping)
    if not aval or not avalu:
        return {"status": f"AVAL or AVALU not found (AVAL={aval}, AVALU={avalu})."}
    mask_miss  = df[aval].notna() & is_blank(df[avalu])
    mask_extra = df[aval].isna()  & ~is_blank(df[avalu])
    keep = [c for c in [uid, test, visit, aval, avalu] if c]
    dist = df[avalu].value_counts(dropna=False).reset_index()
    dist.columns = ["AVALU", "N"]
    return {
        "avalu_distribution":         dist.to_dict("records"),
        "n_aval_present_avalu_blank": int(mask_miss.sum()),
        "n_aval_missing_avalu_set":   int(mask_extra.sum()),
        "records_aval_no_avalu":      df[mask_miss][keep].head(30).to_dict("records"),
        "note": "AVALU should be populated whenever AVAL is non-missing.",
    }


def check_sort_key_uniqueness(df, mapping):
    """Gap 6 — Final sort key USUBJID × PARAMCD × PARAM × AVISIT × VISITCD × NRRLT must be unique."""
    uid     = col(df, "USUBJID", mapping)
    paramcd = next((c for c in ["PARAMCD", "PCTESTCD"] if c in df.columns), None)
    param   = next((c for c in ["PARAM",   "PCTEST"]   if c in df.columns), None)
    avisit  = col(df, "AVISIT",  mapping)
    visitcd = col(df, "VISITCD", mapping)
    nrrlt   = col(df, "NRRLT",   mapping)
    key_vars = [v for v in [uid, paramcd, param, avisit, visitcd, nrrlt] if v]
    if len(key_vars) < 3:
        return {"status": f"Too few sort-key variables found: {key_vars}"}
    counts = df.groupby(key_vars, dropna=False).size().reset_index(name="_cnt")
    dups   = counts[counts["_cnt"] > 1]
    dup_records = []
    if len(dups) > 0:
        merged = dups.merge(df, on=key_vars, how="left")
        dup_records = merged[key_vars + ["_cnt"]].sort_values(key_vars).head(100).to_dict("records")
    return {
        "sort_key_vars":       key_vars,
        "n_total_rows":        int(len(df)),
        "n_duplicate_groups":  int(len(dups)),
        "n_duplicate_records": int(dups["_cnt"].sum()) if len(dups) > 0 else 0,
        "records":             dup_records,
        "note": "Duplicate groups indicate the dataset is not unique on its declared sort key.",
    }


def check_bsacat_derivation(df, mapping):
    """Gap 7 — BSACAT must match BSABL thresholds: Low ≤1.55, Intermediate, High ≥2.15."""
    bsabl  = col(df, "BSABL",  mapping)
    bsacat = col(df, "BSACAT", mapping)
    uid    = col(df, "USUBJID", mapping)
    if not bsabl or not bsacat:
        return {"status": f"BSABL or BSACAT not found (BSABL={bsabl}, BSACAT={bsacat})."}
    df2       = df[[c for c in [uid, bsabl, bsacat] if c]].drop_duplicates().copy()
    bsabl_num = pd.to_numeric(df2[bsabl], errors="coerce")
    df2["_expected"] = bsabl_num.apply(
        lambda v: None if pd.isna(v) else ("Low" if v <= 1.55 else ("High" if v >= 2.15 else "Intermediate"))
    )
    check_mask      = df2["_expected"].notna() & ~is_blank(df2[bsacat])
    df2["_mismatch"] = check_mask & (df2[bsacat].astype(str).str.strip() != df2["_expected"])
    mismatches = df2[df2["_mismatch"]]
    dist = df[bsacat].value_counts(dropna=False).reset_index()
    dist.columns = ["BSACAT", "N"]
    return {
        "bsacat_distribution": dist.to_dict("records"),
        "n_bsabl_missing":     int(bsabl_num.isna().sum()),
        "n_bsacat_missing":    int(is_blank(df[bsacat]).sum()),
        "n_derivation_errors": int(len(mismatches)),
        "mismatch_records":    mismatches.drop(columns=["_mismatch"]).head(50).to_dict("records"),
        "rule": "Low if BSABL ≤ 1.55 | Intermediate if 1.55 < BSABL < 2.15 | High if BSABL ≥ 2.15",
    }


def check_aperiodc_derivation(df, mapping):
    """Gap 8 — APERIODC must equal 'Period {ACYCLE}' for all records."""
    aperiodc = next((c for c in ["APERIODC"] if c in df.columns), None)
    acycle   = next((c for c in ["ACYCLE"]   if c in df.columns), None)
    uid      = col(df, "USUBJID", mapping)
    if not aperiodc or not acycle:
        return {"status": f"APERIODC or ACYCLE not found (APERIODC={aperiodc}, ACYCLE={acycle})."}
    df2        = df[[c for c in [uid, acycle, aperiodc] if c]].drop_duplicates().copy()
    acycle_num = pd.to_numeric(df2[acycle], errors="coerce")
    df2["_expected"]  = "Period " + acycle_num.apply(lambda v: str(int(v)) if pd.notna(v) else "")
    df2["_mismatch"]  = acycle_num.notna() & ~is_blank(df2[aperiodc]) & \
                        (df2[aperiodc].astype(str).str.strip() != df2["_expected"])
    mismatches = df2[df2["_mismatch"]]
    dist = df[aperiodc].value_counts(dropna=False).reset_index()
    dist.columns = ["APERIODC", "N"]
    return {
        "aperiodc_distribution": dist.to_dict("records"),
        "n_acycle_missing":      int(acycle_num.isna().sum()),
        "n_aperiodc_missing":    int(is_blank(df[aperiodc]).sum()),
        "n_derivation_errors":   int(len(mismatches)),
        "mismatch_records":      mismatches.drop(columns=["_mismatch"]).head(50).to_dict("records"),
        "rule": "APERIODC = 'Period ' || strip(ACYCLE)",
    }


def check_aseq_derivation(df, mapping):
    """Gap 9 — ASEQ must be 'AB' for SEQUENCE A and 'BA' for SEQUENCE B."""
    aseq     = next((c for c in ["ASEQ"]     if c in df.columns), None)
    sequence = next((c for c in ["SEQUENCE"] if c in df.columns), None)
    uid      = col(df, "USUBJID", mapping)
    if not aseq or not sequence:
        return {"status": f"ASEQ or SEQUENCE not found (ASEQ={aseq}, SEQUENCE={sequence})."}
    df2 = df[[c for c in [uid, sequence, aseq] if c]].drop_duplicates().copy()
    df2["_expected"] = df2[sequence].apply(
        lambda v: "AB" if "SEQUENCE A" in str(v).upper() else
                  ("BA" if "SEQUENCE B" in str(v).upper() else None)
    )
    df2["_mismatch"] = df2["_expected"].notna() & ~is_blank(df2[aseq]) & \
                       (df2[aseq].astype(str).str.strip() != df2["_expected"])
    mismatches = df2[df2["_mismatch"]]
    crosstab   = (df[[sequence, aseq]].fillna("").astype(str)
                  .value_counts().reset_index().rename(columns={0: "N"}))
    return {
        "sequence_aseq_crosstab": crosstab.to_dict("records"),
        "n_derivation_errors":    int(len(mismatches)),
        "mismatch_records":       mismatches.drop(columns=["_mismatch"]).head(50).to_dict("records"),
        "rule": "ASEQ = 'AB' if SEQUENCE = 'SEQUENCE A'; 'BA' if SEQUENCE = 'SEQUENCE B'",
    }


def check_cohortcd_derivation(df, mapping):
    """Gap 10 — COHORTCD must be non-missing for Phase 1 and consistent with COHORT string."""
    cohortcd = next((c for c in ["COHORTCD"] if c in df.columns), None)
    cohort   = next((c for c in ["COHORT"]   if c in df.columns), None)
    phase    = col(df, "PHASE",   mapping)
    uid      = col(df, "USUBJID", mapping)
    if not cohortcd:
        return {"status": "COHORTCD not found in dataset."}
    dist = df[cohortcd].value_counts(dropna=False).reset_index()
    dist.columns = ["COHORTCD", "N"]
    result = {"cohortcd_distribution": dist.to_dict("records")}
    if phase:
        ph1 = df[df[phase].astype(str).str.strip().str.upper() == "PHASE 1"]
        n_miss = int(is_blank(ph1[cohortcd]).sum())
        result["n_phase1_cohortcd_missing"] = n_miss
        keep = [c for c in [uid, phase, cohort, cohortcd] if c]
        result["phase1_missing_records"] = ph1[is_blank(ph1[cohortcd])][keep].head(30).to_dict("records")
    if cohort:
        cross = (df[[cohort, cohortcd]].fillna("").astype(str)
                 .drop_duplicates().sort_values([cohort, cohortcd]).reset_index(drop=True))
        result["cohort_cohortcd_crosstab"] = cross.head(50).to_dict("records")
    result["note"] = "COHORTCD is derived from the COHORT string. Phase 1 records must always have COHORTCD."
    return result


def check_drugcat_derivation(df, mapping):
    """Gap 11 — DRUGCAT: 'IR'/'DR' for Phase 1 cohort records; blank for Phase 2."""
    drugcat  = next((c for c in ["DRUGCAT"]  if c in df.columns), None)
    cohortcd = next((c for c in ["COHORTCD"] if c in df.columns), None)
    phase    = col(df, "PHASE",   mapping)
    uid      = col(df, "USUBJID", mapping)
    if not drugcat:
        return {"status": "DRUGCAT not found in dataset."}
    dist = df[drugcat].value_counts(dropna=False).reset_index()
    dist.columns = ["DRUGCAT", "N"]
    result = {"drugcat_distribution": dist.to_dict("records")}
    group_cols = [c for c in [phase, cohortcd, drugcat] if c]
    if group_cols:
        cross = (df[group_cols].fillna("").astype(str)
                 .drop_duplicates().sort_values(group_cols).reset_index(drop=True))
        result["phase_cohortcd_drugcat_crosstab"] = cross.to_dict("records")
    if phase and drugcat:
        keep = [c for c in [uid, phase, cohortcd, drugcat] if c]
        ph2 = df[df[phase].astype(str).str.strip().str.upper() == "PHASE 2"]
        ph2_set = ph2[~is_blank(ph2[drugcat])]
        result["n_phase2_drugcat_set"]       = int(len(ph2_set))
        result["phase2_drugcat_set_records"] = ph2_set[keep].head(30).to_dict("records")
        ph1 = df[df[phase].astype(str).str.strip().str.upper() == "PHASE 1"]
        ph1_bad = ph1[~ph1[drugcat].astype(str).str.strip().isin({"IR", "DR", ""})]
        result["n_phase1_unexpected_drugcat"]       = int(len(ph1_bad))
        result["phase1_unexpected_drugcat_records"] = ph1_bad[keep].head(30).to_dict("records")
    result["rule"] = (
        "DRUGCAT = 'IR' for Phase 1 cohorts 1, 2a, 2b, 101, 102; "
        "'DR' for other Phase 1 cohorts; blank for Phase 2."
    )
    return result


def check_trtint_derivation(df, mapping):
    """Gap 12 — TRTINT must be 24 (TRTINTU='h') for specific Phase × TREAT × VISITCD combinations."""
    trtint  = next((c for c in ["TRTINT"]  if c in df.columns), None)
    trtintu = next((c for c in ["TRTINTU"] if c in df.columns), None)
    phase   = col(df, "PHASE",   mapping)
    treat   = next((c for c in ["TREAT"]   if c in df.columns), None)
    visitcd = col(df, "VISITCD", mapping)
    uid     = col(df, "USUBJID", mapping)
    if not trtint:
        return {"status": "TRTINT not found in dataset."}
    dist = df[trtint].value_counts(dropna=False).reset_index()
    dist.columns = ["TRTINT", "N"]
    EXPECTED_24 = [
        ("PHASE 1", "Oral ASTX030", "C1D7"), ("PHASE 1", "Oral ASTX030", "C2D7"),
        ("PHASE 2", "Oral ASTX030", "C1D2"), ("PHASE 2", "Oral ASTX030", "C1D7"),
        ("PHASE 2", "Oral ASTX030", "C2D2"), ("PHASE 2", "Oral ASTX030", "C2D7"),
        ("PHASE 2", "SC Aza",       "C1D7"), ("PHASE 2", "SC Aza",       "C2D7"),
    ]
    result = {
        "trtint_distribution": dist.to_dict("records"),
        "expected_24h_combinations": [{"PHASE": p, "TREAT": t, "VISITCD": v} for p, t, v in EXPECTED_24],
    }
    if phase and treat and visitcd and trtint:
        trtint_num   = pd.to_numeric(df[trtint], errors="coerce")
        keep = [c for c in [uid, phase, treat, visitcd, trtint] + ([trtintu] if trtintu else []) if c]
        should_be_24 = pd.Series(False, index=df.index)
        for ph, tr, vc in EXPECTED_24:
            should_be_24 |= (
                df[phase].astype(str).str.strip().str.upper()   == ph.upper()
            ) & (
                df[treat].astype(str).str.strip()               == tr
            ) & (
                df[visitcd].astype(str).str.strip().str.upper() == vc.upper()
            )
        bad_24        = should_be_24 & (trtint_num != 24)
        unexpected_24 = ~should_be_24 & (trtint_num == 24)
        result["n_should_be_24_but_not"]       = int(bad_24.sum())
        result["records_should_be_24_but_not"] = df[bad_24][keep].head(30).to_dict("records")
        result["n_unexpected_24"]              = int(unexpected_24.sum())
        result["records_unexpected_24"]        = df[unexpected_24][keep].head(30).to_dict("records")
    result["rule"] = (
        "TRTINT=24, TRTINTU='h' for: Phase 1 Oral ASTX030 C1D7/C2D7; "
        "Phase 2 Oral ASTX030 C1D2/C1D7/C2D2/C2D7; Phase 2 SC Aza C1D7/C2D7."
    )
    return result


def check_nca40xrs_consistency(df, mapping):
    """Gap 13 — NCA40XRS must fire when any of ARRLTSF/ARRLT0F/ARRLTEF/ARRLTDF = 'Y'."""
    nca40   = next((c for c in ["NCA40XRS"] if c in df.columns), None)
    drivers = {c: c for c in ["ARRLTSF", "ARRLT0F", "ARRLTEF", "ARRLTDF"] if c in df.columns}
    uid     = col(df, "USUBJID", mapping)
    test    = col(df, "PCTEST",  mapping)
    visit   = col(df, "VISITCD", mapping) or col(df, "AVISIT", mapping)
    if not nca40:
        return {"status": "NCA40XRS not found in dataset."}
    if not drivers:
        return {"status": "None of ARRLTSF/ARRLT0F/ARRLTEF/ARRLTDF found — cannot verify NCA40XRS."}
    any_driver = pd.Series(False, index=df.index)
    for c in drivers.values():
        any_driver |= (df[c].astype(str).str.strip() == "Y")
    nca40_set  = ~is_blank(df[nca40])
    miss_fire  = any_driver & ~nca40_set
    false_fire = nca40_set  & ~any_driver
    keep = [c for c in [uid, test, visit, nca40] + list(drivers.values()) if c]
    return {
        "drivers_found":      list(drivers.keys()),
        "n_any_driver_y":     int(any_driver.sum()),
        "n_nca40xrs_set":     int(nca40_set.sum()),
        "n_miss_fire":        int(miss_fire.sum()),
        "n_false_fire":       int(false_fire.sum()),
        "miss_fire_records":  df[miss_fire][keep].head(30).to_dict("records"),
        "false_fire_records": df[false_fire][keep].head(30).to_dict("records"),
        "note": (
            "NCA40XRS must be populated when ANY of ARRLTSF/ARRLT0F/ARRLTEF/ARRLTDF = 'Y', "
            "and blank when all are missing. Miss-fires and false-fires are both errors."
        ),
    }


def check_crit_subject_overrides(df, mapping):
    """
    Gap 14 — Verify subject-specific CRIT flag overrides from adnca.sas.
      CRIT2FL must be 'N' for 7 named subject/visit combinations.
      CRIT19FL must only be 'Y' for subject 101-053.
      CRIT20FL must only be 'Y' for 5 named subject/visit combinations.
    """
    uid     = col(df, "USUBJID", mapping)
    visitcd = col(df, "VISITCD", mapping)
    subjid  = next((c for c in ["SUBJID"] if c in df.columns), uid)
    result  = {}

    CRIT2_OVERRIDES = [
        ("101-047", "C2D7"), ("103-004", "C2D7"),  ("103-012", "C1D22"),
        ("104-014", "C1D2"), ("104-022", "C1D22"),  ("101-032", "C1D22"),
        ("102-007", "C1D-3"),
    ]
    crit2 = "CRIT2FL" if "CRIT2FL" in df.columns else None
    if crit2 and subjid and visitcd:
        rows = []
        for sbj, vc in CRIT2_OVERRIDES:
            mask = df[subjid].astype(str).str.contains(sbj, na=False) & \
                   (df[visitcd].astype(str).str.strip() == vc)
            sub  = df[mask][[subjid, visitcd, crit2]].copy() if mask.sum() > 0 else pd.DataFrame()
            ok   = all(str(v).strip() == "N" for v in sub[crit2].tolist()) if len(sub) > 0 else None
            rows.append({
                "SUBJID": sbj, "VISITCD": vc,
                "N_records": int(len(sub)),
                "CRIT2FL_values": str(list(sub[crit2].astype(str).str.strip().unique())) if len(sub) > 0 else "not found",
                "Override_OK": "✅" if ok else ("⚠️ not found" if ok is None else "🔴 MISMATCH"),
            })
        result["crit2fl_overrides"] = rows
    else:
        result["crit2fl_overrides"] = {"status": "CRIT2FL or key variables not found."}

    crit19 = "CRIT19FL" if "CRIT19FL" in df.columns else None
    if crit19 and subjid:
        crit19_y    = df[df[crit19].astype(str).str.strip() == "Y"]
        unexpected  = crit19_y[~crit19_y[subjid].astype(str).str.contains("101-053", na=False)]
        result["crit19fl"] = {
            "n_total_y": int(len(crit19_y)), "n_unexpected_y": int(len(unexpected)),
            "unexpected_records": unexpected[[subjid]].head(20).to_dict("records"),
            "note": "CRIT19FL='Y' (PPI inhibitor) should only apply to subject 101-053.",
        }
    else:
        result["crit19fl"] = {"status": "CRIT19FL not found in dataset."}

    CRIT20_EXPECTED = [
        ("101-039", "C1D7"), ("101-040", "C1D7"), ("132-002", "C2D1"),
        ("101-035", "C1D1"), ("101-049", "C2D7"),
    ]
    crit20 = "CRIT20FL" if "CRIT20FL" in df.columns else None
    if crit20 and subjid and visitcd:
        crit20_y = df[df[crit20].astype(str).str.strip() == "Y"]
        def _in_expected(row):
            return any(sbj in str(row[subjid]) and str(row[visitcd]).strip() == vc
                       for sbj, vc in CRIT20_EXPECTED)
        unexpected_c20 = crit20_y[~crit20_y.apply(_in_expected, axis=1)] if len(crit20_y) > 0 else pd.DataFrame()
        missing_flags  = [
            {"SUBJID": sbj, "VISITCD": vc, "Status": "⚠️ CRIT20FL='Y' not found"}
            for sbj, vc in CRIT20_EXPECTED
            if not (df[subjid].astype(str).str.contains(sbj, na=False) &
                    (df[visitcd].astype(str).str.strip() == vc) &
                    (df[crit20].astype(str).str.strip() == "Y")).any()
        ]
        result["crit20fl"] = {
            "n_total_y": int(len(crit20_y)), "n_unexpected_y": int(len(unexpected_c20)),
            "unexpected_records": unexpected_c20[[subjid, visitcd, crit20]].head(20).to_dict("records"),
            "expected_missing":   missing_flags,
            "note": "CRIT20FL='Y' (Atypical) should only apply to 5 named subject/visit combinations.",
        }
    else:
        result["crit20fl"] = {"status": "CRIT20FL or key variables not found."}

    return result


def check_manual_patches(df, mapping):
    """
    Gap 15 — Verify active manual patches from adnca.sas:
      Patch 1: 101-044 C1D22 → DOSEA=20, M_NRRLT_FL='Y'
      Patch 2: 104-022 C2D7  → MRRLT=NRRLT, M_NRRLT_FL='Y'
      Patch 5: 104-014 C1D2  → MRRLT=NRRLT, M_NRRLT_FL='Y'
    """
    uid     = col(df, "USUBJID", mapping)
    subjid  = next((c for c in ["SUBJID"]      if c in df.columns), uid)
    visitcd = col(df, "VISITCD", mapping)
    pctest  = col(df, "PCTEST",  mapping)
    mrrlt   = col(df, "MRRLT",   mapping)
    nrrlt   = col(df, "NRRLT",   mapping)
    dosea   = next((c for c in ["DOSEA"]        if c in df.columns), None)
    mnrrlt  = next((c for c in ["M_NRRLT_FL"]   if c in df.columns), None)
    result  = {"patches": []}
    keep    = [c for c in [subjid, visitcd, pctest, dosea, mnrrlt, mrrlt, nrrlt] if c]

    def _get(subj, vc, extra=None):
        if not subjid or not visitcd:
            return pd.DataFrame()
        mask = df[subjid].astype(str).str.contains(subj, na=False) & \
               (df[visitcd].astype(str).str.strip() == vc)
        if extra is not None:
            mask &= extra
        return df[mask].copy()

    # Patch 1
    p1 = _get("101-044", "C1D22",
               df[pctest].astype(str).str.strip().isin(["Cedazuridine", "E-7727 Epimer"]) if pctest else None)
    p1_info = {"patch": "1", "subject": "101-044", "visitcd": "C1D22",
               "description": "Dose date mismatch → DOSEA=20, M_NRRLT_FL='Y'", "n_records": int(len(p1))}
    p1_info["records"] = p1[keep].to_dict("records") if len(p1) > 0 else []
    if dosea and len(p1) > 0:
        p1_info["n_dosea_not_20"]    = int((pd.to_numeric(p1[dosea], errors="coerce") != 20).sum())
    if mnrrlt and len(p1) > 0:
        p1_info["n_mnrrlt_fl_not_y"] = int((p1[mnrrlt].astype(str).str.strip() != "Y").sum())
    result["patches"].append(p1_info)

    # Patch 2
    p2 = _get("104-022", "C2D7")
    p2_info = {"patch": "2", "subject": "104-022", "visitcd": "C2D7",
               "description": "Date mismatch → MRRLT=NRRLT, M_NRRLT_FL='Y'", "n_records": int(len(p2))}
    p2_info["records"] = p2[keep].to_dict("records") if len(p2) > 0 else []
    if mrrlt and nrrlt and len(p2) > 0:
        p2_info["n_mrrlt_ne_nrrlt"]  = int((pd.to_numeric(p2[mrrlt], errors="coerce") != pd.to_numeric(p2[nrrlt], errors="coerce")).sum())
    if mnrrlt and len(p2) > 0:
        p2_info["n_mnrrlt_fl_not_y"] = int((p2[mnrrlt].astype(str).str.strip() != "Y").sum())
    result["patches"].append(p2_info)

    # Patch 5
    p5 = _get("104-014", "C1D2")
    p5_info = {"patch": "5", "subject": "104-014", "visitcd": "C1D2",
               "description": "Incorrect dates → MRRLT=NRRLT, M_NRRLT_FL='Y'", "n_records": int(len(p5))}
    p5_info["records"] = p5[keep].to_dict("records") if len(p5) > 0 else []
    if mrrlt and nrrlt and len(p5) > 0:
        p5_info["n_mrrlt_ne_nrrlt"]  = int((pd.to_numeric(p5[mrrlt], errors="coerce") != pd.to_numeric(p5[nrrlt], errors="coerce")).sum())
    if mnrrlt and len(p5) > 0:
        p5_info["n_mnrrlt_fl_not_y"] = int((p5[mnrrlt].astype(str).str.strip() != "Y").sum())
    result["patches"].append(p5_info)

    result["note"] = "Only active patches (1, 2, 5) are verified. Patches 3, 4, 6–13 were commented out."
    return result


def check_wide_summaries(df, mapping):
    """
    Gap 16 — Subject-per-visit and timepoint-per-subject wide summaries.
    Flags profiles with ≤ 3 timepoints as sparse.
    """
    uid     = col(df, "USUBJID",  mapping)
    subjid  = next((c for c in ["SUBJID"] if c in df.columns), uid)
    actarm  = col(df, "ACTARM",   mapping)
    test    = col(df, "PCTEST",   mapping)
    visitcd = col(df, "VISITCD",  mapping)
    nrrlt   = col(df, "NRRLT",    mapping)
    result  = {}

    group1 = [c for c in [actarm, test, visitcd] if c]
    if group1 and subjid:
        rows = []
        for name, sub in df.groupby(group1, dropna=False):
            if not isinstance(name, tuple): name = (name,)
            d = dict(zip(group1, name))
            d["N_SUBJECTS"] = int(sub[subjid].nunique())
            d["N_RECORDS"]  = int(len(sub))
            rows.append(d)
        rows.sort(key=lambda r: [str(r.get(c, "")) for c in group1])
        result["subject_per_visit"] = rows
    else:
        result["subject_per_visit"] = {"status": "Required variables not found."}

    group2 = [c for c in [subjid, test, visitcd] if c]
    if group2 and nrrlt:
        rows = []
        for name, sub in df.groupby(group2, dropna=False):
            if not isinstance(name, tuple): name = (name,)
            d  = dict(zip(group2, name))
            tp = sub[nrrlt].dropna().unique()
            d["N_TIMEPOINTS"] = int(len(tp))
            d["TIMEPOINTS"]   = ", ".join(str(v) for v in sorted(tp))
            d["SPARSE"]       = "⚠️ ≤3" if len(tp) <= 3 else ""
            rows.append(d)
        sparse = [r for r in rows if r["SPARSE"]]
        result["timepoints_per_profile"] = {
            "n_profiles_total":  len(rows),
            "n_profiles_sparse": len(sparse),
            "sparse_profiles":   sparse[:100],
            "note": "Profiles with ≤ 3 timepoints may be insufficient for reliable NCA estimation.",
        }
    else:
        result["timepoints_per_profile"] = {"status": "Required variables not found."}

    return result


def check_aval_increase_flags(df, mapping):
    """
    Gap 17 — Re-derive CRIT16FL (increase at 8h) and CRIT17FL (increase at 24h)
    independently and compare against dataset values.
    """
    uid     = col(df, "USUBJID",  mapping)
    test    = col(df, "PCTEST",   mapping)
    visitcd = col(df, "VISITCD",  mapping)
    nrrlt   = col(df, "NRRLT",    mapping)
    aval    = col(df, "AVAL",     mapping)
    crit16  = "CRIT16FL" if "CRIT16FL" in df.columns else None
    crit17  = "CRIT17FL" if "CRIT17FL" in df.columns else None

    if not all([uid, test, visitcd, nrrlt, aval]):
        return {"status": "Required variables not found (need USUBJID, PCTEST, VISITCD, NRRLT, AVAL)."}

    profile_key = [uid, test, visitcd]
    df2 = df[profile_key + [nrrlt, aval]].copy()
    df2[nrrlt] = pd.to_numeric(df2[nrrlt], errors="coerce")
    df2[aval]  = pd.to_numeric(df2[aval],  errors="coerce")
    df2 = df2.sort_values(profile_key + [nrrlt])
    df2["_prev_aval"] = df2.groupby(profile_key, sort=False)[aval].shift(1)
    df2["_increase"]  = df2[aval] > df2["_prev_aval"]

    result = {}
    for tp, crit_col, label in [(8, crit16, "CRIT16FL"), (24, crit17, "CRIT17FL")]:
        tp_mask        = df2[nrrlt] == tp
        increase_at_tp = df2[tp_mask & df2["_increase"]]
        derived_keys   = set(zip(increase_at_tp[uid], increase_at_tp[test], increase_at_tp[visitcd]))
        info = {
            "timepoint": tp,
            "n_profiles_with_increase": len(increase_at_tp.groupby(profile_key).size()),
            "profiles_with_increase":   increase_at_tp[profile_key].drop_duplicates().head(30).to_dict("records"),
        }
        if crit_col:
            crit_vals = df.set_index([uid, test, visitcd])[crit_col].astype(str).str.strip()
            missed = []
            for key in derived_keys:
                try:
                    val = crit_vals.loc[key]
                    val = val.iloc[0] if hasattr(val, "iloc") else val
                    if str(val) != "Y":
                        missed.append({uid: key[0], test: key[1], visitcd: key[2], crit_col: str(val)})
                except KeyError:
                    pass
            flagged_y   = df[df[crit_col] == "Y"][profile_key].drop_duplicates()
            false_flags = [r.to_dict() for _, r in flagged_y.iterrows()
                           if tuple(r[c] for c in profile_key) not in derived_keys]
            info["n_increase_not_flagged"]    = len(missed)
            info["n_flagged_but_no_increase"] = len(false_flags)
            info["increase_not_flagged"]      = missed[:20]
            info["flagged_but_no_increase"]   = false_flags[:20]
        result[label] = info

    return result
