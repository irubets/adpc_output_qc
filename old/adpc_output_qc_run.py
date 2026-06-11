"""
adpc_output_qc_run.py  —  Interactive launcher for the ADPC/ADNCA Output QC pipeline.

Asks for the ADPC/ADNCA dataset path, the production SAS program path, an optional
SAS log file, variable mappings (auto-detected from the dataset, confirmed interactively),
and any optional settings, then runs:
  1. adpc_output_qc_agent.py  → qc_output_results.json + adpc_output_qc.xlsx
  2. adpc_output_qc_report.py → qc_output_report.md + qc_output_report.html

The SAS log is scanned for ERROR and WARNING lines (mirroring %check_log in check_log.sas)
and the results appear as Section O in the QC report.

Usage:
    python adpc_output_qc_run.py                                     # fully interactive
    python adpc_output_qc_run.py --data path/to/adnca.xpt
    python adpc_output_qc_run.py --data path/to/adnca.sas7bdat \\
                                  --sas  path/to/adnca.sas     \\
                                  --log  path/to/adnca.log
    python adpc_output_qc_run.py --data path/ --yes
"""

import argparse
import json
import os
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

# ── Allow running from the package directory without installing ───────────────
sys.path.insert(0, str(Path(__file__).resolve().parent))

from modules.log_analyser import analyse_sas_log, append_log_xlsx

# ── Expected variables and their aliases ──────────────────────────────────────
VARIABLE_REGISTRY = [
    ("USUBJID",  ["SUBJID", "SUBJECTID", "SUBJECT"],             True),
    ("PCTESTCD", ["PARAMCD", "TESTCD"],                          True),
    ("PCTEST",   ["PARAM", "ANALYTE", "TEST", "ATEST"],          True),
    ("AVISIT",   ["VISIT", "ANALYSIS_VISIT"],                    True),
    ("APROFILE", ["PROFILE", "PROFILEID"],                       True),
    ("VISITCD",  ["AVISITCD", "VISITCODE"],                      False),
    ("PHASE",    ["STUDYPHASE"],                                  False),
    ("ACTARM",   ["ARM", "ACTARMCD"],                            False),
    ("ACTARMCD", ["ARMCD"],                                      False),
    ("NRRLT",    ["PCTPTNUM", "NOMINAL_TIME", "NOMTIME"],        True),
    ("ARRLT",    ["RELTM", "ACTUAL_TIME"],                       False),
    ("MRRLT",    ["MRRLT"],                                      True),
    ("ADTM",     ["PCDTC", "SAMPLEDTM"],                         False),
    ("ARFTDTM",  ["PCRFTDTM", "REFDTM", "DOSEDTM"],             True),
    ("AVAL",     ["CONC", "DV"],                                 True),
    ("ARESC",    ["PCORRES", "CONCENTRATION_CHAR"],              False),
    ("AVALU",    ["PCORRESU", "UNIT"],                           False),
    ("DTYPE",    ["RECORDTYPE"],                                 True),
    ("RICHRFL",  ["RICH_FL", "RICHFL"],                          False),
    ("NCA01FL",  ["NCAFL", "NCA_FL"],                            True),
    ("PKSUMXFL", ["SUMFL", "PCSUMXFL", "PKSUM_FL"],             True),
    ("NCAXFL",   ["NCA_XRS_FL"],                                 False),
    ("LISTRFL",  ["LIST_FL", "LISTFL"],                          False),
    ("DOSEA",    ["DOSE", "DOSEAMT"],                            False),
    ("DOSEP",    ["DOSEPROT"],                                   False),
    ("DOSEU",    ["DOSEU"],                                      False),
    ("EXROUTE",  ["ECROUTE", "ROUTE"],                           False),
    ("TREAT",    ["TRT", "TREATMENT"],                           False),
    ("SATRT",    ["SHORT_TRT"],                                   False),
    ("ATRT",     ["ANALYSIS_TRT"],                               False),
    ("TRT01A",   ["TRT01A"],                                     False),
    ("BSABL",    ["BSA", "BSABASE"],                             False),
    ("BSACAT",   ["BSACATEGORY"],                                False),
    ("SEQUENCE", ["SEQ", "CROSSOVERSEQ"],                        False),
    ("VOMFL",    ["VOM_FL"],                                     False),
    ("PCVOMYN",  ["VOMITING"],                                   False),
]


# ── Colour helpers ────────────────────────────────────────────────────────────
def _coloured(text, code):
    if sys.stdout.isatty() and os.name != "nt":
        return f"\033[{code}m{text}\033[0m"
    return text

def bold(t):   return _coloured(t, "1")
def green(t):  return _coloured(t, "32")
def yellow(t): return _coloured(t, "33")
def red(t):    return _coloured(t, "31")
def cyan(t):   return _coloured(t, "36")
def grey(t):   return _coloured(t, "90")


def _prompt(label, default=None, required=False):
    hint = f"  [{default}]" if default is not None else ""
    while True:
        raw = input(f"  {bold(label)}{hint}: ").strip()
        if raw:         return raw
        if default is not None: return str(default)
        if not required: return ""
        print(f"  {yellow('(required — please enter a value)')}")


def _confirm(question, default_yes=True):
    hint = "[Y/n]" if default_yes else "[y/N]"
    raw  = input(f"  {bold(question)} {hint}: ").strip().lower()
    return default_yes if not raw else raw in ("y", "yes")


def _here():
    return Path(__file__).resolve().parent


def _find_script(name):
    for c in [_here() / name, Path.cwd() / name]:
        if c.exists(): return c
    print(red(f"\n  ✗ Cannot find {name}"))
    sys.exit(1)


def _run_step(label, cmd):
    print(f"\n{bold('─' * 60)}\n{bold(label)}\n{bold('─' * 60)}")
    result = subprocess.run(cmd, text=True)
    if result.returncode != 0:
        print(red(f"\n  ✗ {label} failed (exit code {result.returncode})"))
        return False
    return True


def _read_column_names(path):
    ext = Path(path).suffix.lower()
    try:
        if ext == ".csv":
            import pandas as pd
            return [c.upper() for c in pd.read_csv(path, nrows=0).columns]
        elif ext in (".xpt", ".sas7bdat"):
            import pyreadstat
            fn = pyreadstat.read_xport if ext == ".xpt" else pyreadstat.read_sas7bdat
            _, meta = fn(path, metadataonly=True)
            return [c.upper() for c in meta.column_names]
    except Exception as e:
        print(yellow(f"  Could not read column names: {e}"))
    return []


def _autodetect_mapping(columns):
    col_upper = {c.upper(): c for c in columns}
    return {
        canonical: next(
            (col_upper[cand.upper()] for cand in [canonical] + aliases if cand.upper() in col_upper),
            None
        )
        for canonical, aliases, _ in VARIABLE_REGISTRY
    }


def _confirm_mapping(mapping, columns, args_yes=False):
    print(f"\n{bold('Step 3 — Variable mapping')}")
    print(grey("  Auto-detected from dataset columns. Press Enter to accept,"))
    print(grey("  or type the correct column name. Leave blank to skip optional variables.\n"))
    col_upper = {c.upper() for c in columns}
    confirmed, user_override_misses = {}, []
    required_vars = [(c, a, r) for c, a, r in VARIABLE_REGISTRY if r]
    optional_vars = [(c, a, r) for c, a, r in VARIABLE_REGISTRY if not r]

    def _handle(canonical, aliases, required):
        detected = mapping.get(canonical)
        if args_yes and detected:
            confirmed[canonical] = detected; return
        tag    = bold("[required]") if required else grey("[optional]")
        status = green(f"✅ {detected}") if detected else yellow("— not found")
        if args_yes:
            confirmed[canonical] = detected; return
        raw = input(f"  {tag} {canonical:<30} {status}: ").strip()
        if not raw:
            confirmed[canonical] = detected
        elif raw.upper() in col_upper:
            confirmed[canonical] = raw.upper()
            print(f"    {green('✓')} mapped to {cyan(raw.upper())}")
        else:
            confirmed[canonical] = None
            user_override_misses.append((canonical, raw))
            print(f"    {red('✗')} '{raw}' not found in dataset — logged as QC finding")

    print(bold("  — Required variables —"))
    for c, a, r in required_vars:
        _handle(c, a, r)
    print()
    show_opt = False if args_yes else _confirm("Review optional variable mappings?", default_yes=False)
    if show_opt:
        print(bold("\n  — Optional variables —"))
        for c, a, r in optional_vars:
            _handle(c, a, r)
    else:
        for c, a, r in optional_vars:
            confirmed[c] = mapping.get(c)

    return confirmed, user_override_misses


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    cli = argparse.ArgumentParser(
        description="Interactive launcher for adpc_output_qc pipeline",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    cli.add_argument("--data",  default=None, help="Path to ADPC/ADNCA dataset (CSV/XPT/SAS7BDAT)")
    cli.add_argument("--sas",   default=None, help="Path to production SAS program (adnca.sas)")
    cli.add_argument("--log",   default=None, help="Path to SAS log file (adnca.log)")
    cli.add_argument("--specs", default=None, help="Path to specifications Excel file")
    cli.add_argument("--phase", default=None, help="Phase filter (e.g. 'PHASE 1', default ALL)")
    cli.add_argument("--yes",   action="store_true", help="Skip confirmations, accept auto-detected mapping")
    args = cli.parse_args()

    print()
    print(bold("╔══════════════════════════════════════════════════════╗"))
    print(bold("║     ADPC / ADNCA Output Data QC — Interactive Run    ║"))
    print(bold("╚══════════════════════════════════════════════════════╝"))
    print()

    # Step 1 — dataset
    print(bold("Step 1 — ADPC / ADNCA dataset"))
    if args.data:
        data_path = Path(args.data).expanduser().resolve()
        print(f"  Dataset : {cyan(str(data_path))}")
    else:
        print("  Supported formats: .csv  .xpt  .sas7bdat")
        data_path = Path(_prompt("Dataset path", required=True)).expanduser().resolve()
    if not data_path.is_file():
        print(red(f"\n  ✗ File not found: {data_path}")); sys.exit(1)
    out_dir = data_path.parent / ("adnca_review_" + datetime.now().strftime("%Y-%m-%dT%H_%M"))
    out_dir.mkdir(parents=True, exist_ok=True)
    print(f"  {green('✓')} Output directory: {cyan(str(out_dir))}")

    # Step 2 — SAS program
    print(); print(bold("Step 2 — Production SAS program"))
    print(grey("  Used to parse whichc() / cmiss() derivation logic."))
    print(grey("  Leave blank to skip (flag lists will use defaults)."))
    print()
    sas_path = args.sas or _prompt("Path to adnca.sas (or blank to skip)", default="")
    if sas_path:
        sp = Path(sas_path).expanduser().resolve()
        if not sp.exists():
            print(yellow(f"  ⚠ SAS file not found: {sp} — flag derivation will use defaults"))
            sas_path = ""
        else:
            sas_path = str(sp); print(green("  ✓ SAS program found"))

    # Step 2.5 — specs file
    print(); print(bold("Step 2.5 — Specifications file (Excel)"))
    print(grey("  Expects a sheet named 'ADPC' or 'ADNCA' with columns: Name, Label, Core."))
    print(grey("  Leave blank to skip."))
    print()
    specs_path = args.specs or _prompt("Path to specifications Excel file (or blank to skip)", default="")
    if specs_path:
        xp = Path(specs_path).expanduser().resolve()
        if not xp.exists():
            print(yellow(f"  ⚠ Specs file not found: {xp} — skipping"))
            specs_path = ""
        elif xp.suffix.lower() not in (".xlsx", ".xlsm", ".xls"):
            print(yellow("  ⚠ Not an Excel file — skipping"))
            specs_path = ""
        else:
            specs_path = str(xp); print(green("  ✓ Specs file found"))

    # Step 2.7 — SAS log
    print(); print(bold("Step 2.7 — SAS log file"))
    print(grey("  Scanned for ERROR and WARNING lines (mirrors %check_log)."))
    print(grey("  Results appear as Section O in the QC report."))
    print(grey("  Leave blank to skip."))
    print()
    if args.log:
        log_path = args.log
    else:
        _auto_log = ""
        for _dir in [data_path.parent, Path(sas_path).parent if sas_path else None]:
            if _dir is None: continue
            for _stem in [data_path.stem, "adnca"]:
                _cand = _dir / f"{_stem}.log"
                if _cand.exists():
                    _auto_log = str(_cand); break
            if _auto_log: break
        log_path = _prompt("Path to SAS log file (or blank to skip)", default=_auto_log or "")
    log_path_resolved = ""
    if log_path:
        lp = Path(log_path).expanduser().resolve()
        if not lp.exists():
            print(yellow(f"  ⚠ Log file not found: {lp} — log analysis will be skipped"))
        else:
            log_path_resolved = str(lp); print(green("  ✓ Log file found"))

    # Step 3 — variable mapping
    print(); print(bold("Step 3 — Reading dataset column names …"))
    columns      = _read_column_names(str(data_path))
    auto_mapping = _autodetect_mapping(columns) if columns else {c: None for c, _, _ in VARIABLE_REGISTRY}
    if columns:
        n_found    = sum(1 for v in auto_mapping.values() if v)
        n_req      = sum(1 for _, _, r in VARIABLE_REGISTRY if r)
        req_found  = sum(1 for c, _, r in VARIABLE_REGISTRY if r and auto_mapping.get(c))
        print(f"  {green('✓')} {len(columns)} columns found")
        print(f"  Auto-detected: {green(str(n_found))} / {len(auto_mapping)} variables "
              f"({green(str(req_found))} / {n_req} required)")
        req_miss = [c for c, _, r in VARIABLE_REGISTRY if r and not auto_mapping.get(c)]
        if req_miss:
            print(f"  {yellow('⚠ Required variables not auto-detected:')} {', '.join(req_miss)}")
    confirmed_mapping, user_override_misses = _confirm_mapping(auto_mapping, columns, args_yes=args.yes)

    # Step 4 — options
    print(); print(bold("Step 4 — Options"))
    phase = args.phase or _prompt("Phase filter (e.g. 'PHASE 1', blank for ALL)", default="ALL")

    # Step 5 — review
    json_out = out_dir / "qc_output_results.json"
    xlsx_out = out_dir / "adpc_output_qc.xlsx"
    md_out   = out_dir / "qc_output_report.md"
    html_out = out_dir / "qc_output_report.html"

    print(); print(bold("Step 5 — Review and confirm"))
    print(f"  Dataset      : {cyan(str(data_path))}")
    print(f"  SAS program  : {cyan(sas_path) if sas_path else yellow('(not provided)')}")
    print(f"  SAS log      : {cyan(log_path_resolved) if log_path_resolved else yellow('(not provided — log analysis skipped)')}")
    print(f"  Specs file   : {cyan(specs_path) if specs_path else yellow('(not provided)')}")
    print(f"  Phase filter : {cyan(phase)}")
    print()
    print("  Output files:")
    for p in (json_out, xlsx_out, md_out, html_out):
        print(f"    {cyan(str(p))}")
    print()

    if user_override_misses:
        print(yellow(f"  ⚠ {len(user_override_misses)} user-provided variable name(s) not found in dataset:"))
        for canon, provided in user_override_misses:
            print(yellow(f"      {canon} → '{provided}' (not in dataset — will appear as QC finding)"))
        print()

    if not args.yes:
        if not _confirm("Run the QC pipeline now?", default_yes=True):
            print(yellow("\n  Aborted.\n")); sys.exit(0)

    # Serialise config
    run_config = {
        "data_path":            str(data_path),
        "sas_path":             sas_path or "",
        "log_path":             log_path_resolved or "",
        "specs_path":           specs_path or "",
        "phase_filter":         phase,
        "mapping":              {k: v for k, v in confirmed_mapping.items() if v},
        "user_override_misses": [{"canonical": c, "provided": p} for c, p in user_override_misses],
    }
    config_path = out_dir / "_adpc_output_qc_config.json"
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(run_config, f, indent=2)

    # Copy input files
    print(bold("\nCopying input files into review folder …"))
    for src, label in [
        (str(data_path),        "Dataset "),
        (sas_path,              "SAS prog"),
        (specs_path,            "Specs   "),
        (log_path_resolved,     "SAS log "),
    ]:
        if src:
            try:
                shutil.copy2(src, str(out_dir / Path(src).name))
                print(f"  {green('✓')} {label}: {cyan(str(out_dir / Path(src).name))}")
            except Exception as e:
                print(yellow(f"  ⚠ Could not copy {label}: {e}"))

    # Locate sibling scripts
    agent_script  = _find_script("adpc_output_qc_agent.py")
    report_script = _find_script("adpc_output_qc_report.py")

    agent_cmd = [sys.executable, str(agent_script),
                 "--config", str(config_path), "--out", str(json_out), "--xlsx", str(xlsx_out)]
    report_cmd = [sys.executable, str(report_script),
                  "--json", str(json_out), "--out_dir", str(out_dir)]

    # Step A — run agent
    ok = _run_step("Step A — Running QC agent (adpc_output_qc_agent.py)", agent_cmd)
    if not ok:
        print(red("\n  Pipeline stopped.\n")); sys.exit(1)
    if not json_out.exists():
        print(red(f"\n  ✗ Expected JSON not found: {json_out}")); sys.exit(1)

    # Step C — inject log results
    if log_path_resolved:
        print(f"\n{bold('─' * 60)}")
        print(bold("Step C — Analysing SAS log for ERRORs and WARNINGs"))
        print(bold("─" * 60))
        log_results = analyse_sas_log(log_path_resolved)
        _ne = log_results.get("n_errors",   0)
        _nw = log_results.get("n_warnings", 0)
        print(f"  ERRORs   : {red(str(_ne))   if _ne   else green('0')}")
        print(f"  WARNINGs : {yellow(str(_nw)) if _nw else green('0')}")
        try:
            with open(json_out, encoding="utf-8") as _f:
                _qc = json.load(_f)
            _qc["sas_log_analysis"] = log_results
            with open(json_out, "w", encoding="utf-8") as _f:
                json.dump(_qc, _f, indent=2)
            print(f"  {green('✓')} Log results merged into JSON")
        except Exception as _e:
            print(yellow(f"  ⚠ Could not merge log results into JSON: {_e}"))
    else:
        print(f"\n  {grey('(Step C skipped — no SAS log provided)')}")

    # Step B — generate report
    ok = _run_step("Step B — Generating report (adpc_output_qc_report.py)", report_cmd)
    if not ok:
        print(red("\n  Report generation failed — JSON and XLSX are intact."))
        print(f"  Re-run manually:\n    python {report_script} --json {json_out}\n")
        sys.exit(1)

    # Step D — append log sheets to XLSX
    if log_path_resolved and xlsx_out.exists():
        print(f"\n{bold('─' * 60)}")
        print(bold("Step D — Appending SAS log sheets to XLSX"))
        print(bold("─" * 60))
        append_log_xlsx(xlsx_out, log_path_resolved)

    # Done
    print()
    print(bold("╔══════════════════════════════════════════════════════╗"))
    print(bold("║                     ✅  Done                         ║"))
    print(bold("╚══════════════════════════════════════════════════════╝"))
    print()
    print("  Output files:")
    for p in (json_out, xlsx_out, md_out, html_out):
        icon = green("✅") if p.exists() else red("✗ missing")
        print(f"    {icon}  {cyan(str(p))}")
    if log_path_resolved:
        _lr  = json.loads(json_out.read_text(encoding="utf-8")).get("sas_log_analysis", {})
        _ne  = _lr.get("n_errors",   0)
        _nw  = _lr.get("n_warnings_review", 0)
        _icon = red("🔴") if _ne else (yellow("⚠️") if _nw else green("✅"))
        print(f"\n  SAS log  : {_icon}  {_ne} ERROR(s), {_nw} WARNING(s) to review "
              f"— see Section O in the HTML report")
    print()


if __name__ == "__main__":
    main()
