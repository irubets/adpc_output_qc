"""
adpc_output_qc_agent.py
ADPC / ADNCA output dataset QC agent.

Reads the final ADPC or ADNCA dataset (CSV / XPT / SAS7BDAT), parses the
production SAS program for flag derivation logic, runs all QC checks, and
writes:
  - qc_output_results.json   (structured results consumed by the report)
  - adpc_output_qc.xlsx      (multi-sheet Excel workbook for manual review)

All logic lives in modules/. This file is the thin orchestrator.

Usage:
    python adpc_output_qc_agent.py --config _adpc_output_qc_config.json
    python adpc_output_qc_agent.py --config config.json --out results.json --xlsx out.xlsx

Requirements:
    pip install pandas pyreadstat openpyxl
"""

import argparse
import json
import sys
import warnings
from datetime import datetime
from pathlib import Path

# ── Allow running from the package directory without installing ───────────────
sys.path.insert(0, str(Path(__file__).resolve().parent))

from modules.utils import (
    load_dataset, json_safe,
    DEFAULT_PKSUMXFL_CRITS, DEFAULT_NCAXFL_NCAXRS,
)
from modules.checks_core import (
    parse_sas_flag_logic, check_specs_comparison,
    check_meta, check_subject_counts, check_treatment_mapping,
    check_route_table, check_crit_nca_crossref, check_populations,
    check_dtype_flags, check_aprofile_alignment, check_mrrlt_duplicates,
    check_erroneous_records, check_missing_dose, check_prepost_flags,
    check_vomiting, check_profile_types, check_richrfl_listrfl,
)
from modules.checks_gap import (
    check_phase3_exclusion, check_astx030_exclusion,
    check_avalu_consistency, check_sort_key_uniqueness,
    check_bsacat_derivation, check_aperiodc_derivation,
    check_aseq_derivation, check_cohortcd_derivation,
    check_drugcat_derivation, check_trtint_derivation,
    check_nca40xrs_consistency, check_crit_subject_overrides,
    check_manual_patches, check_wide_summaries, check_aval_increase_flags,
)
from modules.xlsx_builder import build_xlsx

warnings.filterwarnings("ignore")

_DEFAULTS = {
    "pksumxfl_crits": DEFAULT_PKSUMXFL_CRITS,
    "ncaxfl_ncaxrs":  DEFAULT_NCAXFL_NCAXRS,
}


def run_all(config: dict, out_json: str, out_xlsx: str):
    data_path            = config["data_path"]
    sas_path             = config.get("sas_path", "")
    specs_path           = config.get("specs_path", "")
    mapping              = {k.upper(): v for k, v in config.get("mapping", {}).items()}
    user_override_misses = config.get("user_override_misses", [])

    print(f"[adpc_output_qc_agent] Loading dataset: {data_path}")
    df = load_dataset(data_path)
    print(f"  {len(df):,} records × {len(df.columns)} columns")

    print("[adpc_output_qc_agent] Parsing SAS flag derivation logic …")
    sas_logic = parse_sas_flag_logic(sas_path, _DEFAULTS)
    for w in sas_logic.get("parse_warnings", []):
        print(f"  ⚠ {w}")
    print(f"  Parse method : {sas_logic['parse_method']}")
    print(f"  PKSUMXFL CRITs ({len(sas_logic['pksumxfl_crits'])}): "
          f"{', '.join(sas_logic['pksumxfl_crits'])}")
    print(f"  NCAXFL XRS    ({len(sas_logic['ncaxfl_ncaxrs'])}): "
          f"{', '.join(sas_logic['ncaxfl_ncaxrs'])}")

    print("[adpc_output_qc_agent] Running: specs_comparison …")
    specs_result = check_specs_comparison(df, specs_path)
    specs_vars   = [v["Name"] for v in specs_result.get("specs_vars", [])]
    print(f"  ✓ specs_comparison  "
          f"({specs_result.get('n_specs_vars', '—')} spec vars, "
          f"{specs_result.get('n_missing_from_dataset', '—')} missing from dataset)")

    results = {
        "meta": {
            "generated":            datetime.now().isoformat(timespec="seconds"),
            "data_path":            data_path,
            "sas_path":             sas_path,
            "specs_path":           specs_path,
            "phase_filter":         config.get("phase_filter", "ALL"),
            "n_records":            int(len(df)),
            "n_columns":            int(len(df.columns)),
            "user_override_misses": user_override_misses,
        },
        "sas_flag_logic":   sas_logic,
        "specs_comparison": specs_result,
    }

    sections = [
        # ── Core checks ───────────────────────────────────────────────────────
        ("meta_detail",        check_meta,              (df, data_path, mapping, user_override_misses)),
        ("subject_counts",     check_subject_counts,    (df, mapping)),
        ("treatment_mapping",  check_treatment_mapping, (df, mapping)),
        ("route_table",        check_route_table,       (df, mapping)),
        ("crit_nca_crossref",  check_crit_nca_crossref, (df, mapping, sas_logic)),
        ("populations",        check_populations,       (df, mapping, sas_logic)),
        ("dtype_flags",        check_dtype_flags,       (df, mapping, sas_path, specs_vars or None)),
        ("aprofile_alignment", check_aprofile_alignment,(df, mapping)),
        ("mrrlt_duplicates",   check_mrrlt_duplicates,  (df, mapping)),
        ("erroneous_records",  check_erroneous_records, (df, mapping)),
        ("missing_dose",       check_missing_dose,      (df, mapping, sas_logic)),
        ("prepost_flags",      check_prepost_flags,     (df, mapping)),
        ("vomiting",           check_vomiting,          (df, mapping)),
        ("profile_types",      check_profile_types,     (df, mapping)),
        ("richrfl_listrfl",    check_richrfl_listrfl,   (df, mapping)),
        # ── Gap checks ────────────────────────────────────────────────────────
        ("phase3_exclusion",      check_phase3_exclusion,       (df, mapping)),
        ("astx030_exclusion",     check_astx030_exclusion,      (df, mapping)),
        ("avalu_consistency",     check_avalu_consistency,      (df, mapping)),
        ("sort_key_uniqueness",   check_sort_key_uniqueness,    (df, mapping)),
        ("bsacat_derivation",     check_bsacat_derivation,      (df, mapping)),
        ("aperiodc_derivation",   check_aperiodc_derivation,    (df, mapping)),
        ("aseq_derivation",       check_aseq_derivation,        (df, mapping)),
        ("cohortcd_derivation",   check_cohortcd_derivation,    (df, mapping)),
        ("drugcat_derivation",    check_drugcat_derivation,     (df, mapping)),
        ("trtint_derivation",     check_trtint_derivation,      (df, mapping)),
        ("nca40xrs_consistency",  check_nca40xrs_consistency,   (df, mapping)),
        ("crit_subject_overrides",check_crit_subject_overrides, (df, mapping)),
        ("manual_patches",        check_manual_patches,         (df, mapping)),
        ("wide_summaries",        check_wide_summaries,         (df, mapping)),
        ("aval_increase_flags",   check_aval_increase_flags,    (df, mapping)),
    ]

    for name, fn, args in sections:
        print(f"[adpc_output_qc_agent] Running: {name} …")
        try:
            results[name] = fn(*args)
            print(f"  ✓ {name}")
        except Exception as e:
            import traceback
            print(f"  ✗ {name}: {e}")
            results[name] = {"error": str(e), "traceback": traceback.format_exc()}

    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(json_safe(results), f, indent=2)
    print(f"[adpc_output_qc_agent] JSON written: {out_json}")

    print("[adpc_output_qc_agent] Building XLSX …")
    try:
        build_xlsx(df, mapping, results, sas_logic, out_xlsx)
    except Exception as e:
        import traceback
        print(f"  ✗ XLSX build failed: {e}")
        traceback.print_exc()

    return results


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="ADPC Output QC Agent")
    parser.add_argument("--config", required=True, help="Path to _adpc_output_qc_config.json")
    parser.add_argument("--out",    default=None,  help="Output JSON path")
    parser.add_argument("--xlsx",   default=None,  help="Output XLSX path")
    args = parser.parse_args()

    with open(args.config, encoding="utf-8") as f:
        config = json.load(f)

    out_dir  = Path(config["data_path"]).parent
    out_json = args.out  or str(out_dir / "qc_output_results.json")
    out_xlsx = args.xlsx or str(out_dir / "adpc_output_qc.xlsx")

    run_all(config, out_json, out_xlsx)
