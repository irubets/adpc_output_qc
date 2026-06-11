"""
adpc_output_qc_run.py  —  Launcher for the ADPC/ADNCA Output QC pipeline.

TWO MODES:

  1. Config-file mode (recommended):
       python adpc_output_qc_run.py --input adnca_qc.cfg
       python adpc_output_qc_run.py --input adnca_qc.cfg --yes

  2. Interactive mode (no config file):
       python adpc_output_qc_run.py

  3. Generate a pre-filled config template from a dataset:
       python adpc_output_qc_run.py --generate-config adnca_qc.cfg --data path/to/adnca.xpt

Config file format (.cfg):

    [sources]
    data  = /path/to/adnca.xpt          # required
    sas   = /path/to/adnca.sas          # optional
    log   = /path/to/adnca.log          # optional
    specs = /path/to/Specifications.xlsx # optional

    [options]
    phase = ALL      # ALL | PHASE 1 | PHASE 2
    yes   = false    # true = skip all confirmation prompts

    [mapping]
    USUBJID  = USUBJID   # CANONICAL = ACTUAL_COLUMN_IN_DATASET
    AVAL     = AVAL
    # ... only lines that differ from auto-detection are needed

Paths in the config file may be absolute or relative to the config file's
own directory, making the file portable across machines.
"""

import argparse
import configparser
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
    ("USUBJID",  ["SUBJID", "SUBJECTID", "SUBJECT"],          True),
    ("PCTESTCD", ["PARAMCD", "TESTCD"],                       True),
    ("PCTEST",   ["PARAM", "ANALYTE", "TEST", "ATEST"],       True),
    ("AVISIT",   ["VISIT", "ANALYSIS_VISIT"],                 True),
    ("APROFILE", ["PROFILE", "PROFILEID"],                    True),
    ("VISITCD",  ["AVISITCD", "VISITCODE"],                   False),
    ("PHASE",    ["STUDYPHASE"],                               False),
    ("ACTARM",   ["ARM", "ACTARMCD"],                         False),
    ("ACTARMCD", ["ARMCD"],                                    False),
    ("NRRLT",    ["PCTPTNUM", "NOMINAL_TIME", "NOMTIME"],     True),
    ("ARRLT",    ["RELTM", "ACTUAL_TIME"],                    False),
    ("MRRLT",    ["MRRLT"],                                    True),
    ("ADTM",     ["PCDTC", "SAMPLEDTM"],                      False),
    ("ARFTDTM",  ["PCRFTDTM", "REFDTM", "DOSEDTM"],          True),
    ("AVAL",     ["CONC", "DV"],                               True),
    ("ARESC",    ["PCORRES", "CONCENTRATION_CHAR"],           False),
    ("AVALU",    ["PCORRESU", "UNIT"],                        False),
    ("DTYPE",    ["RECORDTYPE"],                               True),
    ("RICHRFL",  ["RICH_FL", "RICHFL"],                       False),
    ("NCA01FL",  ["NCAFL", "NCA_FL"],                         True),
    ("PKSUMXFL", ["SUMFL", "PCSUMXFL", "PKSUM_FL"],          True),
    ("NCAXFL",   ["NCA_XRS_FL"],                              False),
    ("LISTRFL",  ["LIST_FL", "LISTFL"],                       False),
    ("DOSEA",    ["DOSE", "DOSEAMT"],                         False),
    ("DOSEP",    ["DOSEPROT"],                                 False),
    ("DOSEU",    ["DOSEU"],                                    False),
    ("EXROUTE",  ["ECROUTE", "ROUTE"],                        False),
    ("TREAT",    ["TRT", "TREATMENT"],                        False),
    ("SATRT",    ["SHORT_TRT"],                                False),
    ("ATRT",     ["ANALYSIS_TRT"],                            False),
    ("TRT01A",   ["TRT01A"],                                   False),
    ("BSABL",    ["BSA", "BSABASE"],                          False),
    ("BSACAT",   ["BSACATEGORY"],                             False),
    ("SEQUENCE", ["SEQ", "CROSSOVERSEQ"],                     False),
    ("VOMFL",    ["VOM_FL"],                                   False),
    ("PCVOMYN",  ["VOMITING"],                                 False),
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
        if raw:               return raw
        if default is not None: return str(default)
        if not required:      return ""
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
            (col_upper[cand.upper()] for cand in [canonical] + aliases
             if cand.upper() in col_upper),
            None,
        )
        for canonical, aliases, _ in VARIABLE_REGISTRY
    }


def _resolve_path(raw: str, base_dir: Path) -> str:
    """Resolve a path that may be absolute or relative to base_dir."""
    if not raw:
        return ""
    p = Path(raw.strip())
    if not p.is_absolute():
        p = base_dir / p
    return str(p.resolve())


# ═════════════════════════════════════════════════════════════════════════════
# Config-file reader
# ═════════════════════════════════════════════════════════════════════════════

def load_cfg(cfg_path: Path) -> dict:
    """
    Parse an INI-style .cfg file and return a normalised settings dict with keys:
      data_path, sas_path, log_path, specs_path, phase, yes, mapping_overrides
    Paths are resolved relative to the cfg file's own directory.
    """
    if not cfg_path.exists():
        print(red(f"\n  ✗ Config file not found: {cfg_path}"))
        sys.exit(1)

    base_dir = cfg_path.parent

    # configparser requires at least one section; add a fallback header
    text = cfg_path.read_text(encoding="utf-8")
    cp   = configparser.ConfigParser(
        inline_comment_prefixes=("#",),
        comment_prefixes=("#",),
    )
    cp.read_string(text)

    def _get(section, key, fallback=""):
        try:
            return cp.get(section, key).strip()
        except (configparser.NoSectionError, configparser.NoOptionError):
            return fallback

    data_path  = _resolve_path(_get("sources", "data"),  base_dir)
    sas_path   = _resolve_path(_get("sources", "sas"),   base_dir)
    log_path   = _resolve_path(_get("sources", "log"),   base_dir)
    specs_path = _resolve_path(_get("sources", "specs"), base_dir)
    phase      = _get("options", "phase", "ALL")
    yes_val    = _get("options", "yes",   "false").lower() in ("true", "yes", "1")

    # Read [mapping] section — all keys are canonical variable names
    mapping_overrides = {}
    if cp.has_section("mapping"):
        for canonical, actual in cp.items("mapping"):
            actual = actual.strip()
            if actual:
                mapping_overrides[canonical.upper()] = actual.upper()

    return {
        "data_path":        data_path,
        "sas_path":         sas_path,
        "log_path":         log_path,
        "specs_path":       specs_path,
        "phase":            phase,
        "yes":              yes_val,
        "mapping_overrides":mapping_overrides,
    }


def _validate_cfg_paths(cfg: dict) -> bool:
    """Check that required paths exist and optional ones warn if missing. Returns True if OK."""
    ok = True

    if not cfg["data_path"]:
        print(red("  ✗ [sources] data = ... is required but not set in the config file."))
        ok = False
    elif not Path(cfg["data_path"]).exists():
        print(red(f"  ✗ Dataset not found: {cfg['data_path']}"))
        ok = False
    else:
        print(green(f"  ✓ data   : {cfg['data_path']}"))

    for key, label in [("sas_path", "sas  "), ("log_path", "log  "), ("specs_path", "specs")]:
        val = cfg[key]
        if val:
            if Path(val).exists():
                print(green(f"  ✓ {label} : {val}"))
            else:
                print(yellow(f"  ⚠ {label} : not found — {val}"))
                cfg[key] = ""   # blank it so downstream skips it
        else:
            print(grey(f"  — {label} : (not provided)"))

    return ok


# ═════════════════════════════════════════════════════════════════════════════
# Config template generator
# ═════════════════════════════════════════════════════════════════════════════

def generate_config(out_path: Path, data_path_hint: str = ""):
    """
    Write a pre-filled .cfg template.  If data_path_hint points to a real
    dataset, the [mapping] section is populated with auto-detected column names.
    """
    columns      = _read_column_names(data_path_hint) if data_path_hint else []
    auto_mapping = _autodetect_mapping(columns)       if columns else {}

    lines = [
        "# ─────────────────────────────────────────────────────────────",
        "#  adnca_qc.cfg  —  Input file for adpc_output_qc_run.py",
        "#",
        "#  Usage:",
        "#      python adpc_output_qc_run.py --input adnca_qc.cfg",
        "#",
        "#  Rules:",
        "#    • Lines starting with # are comments",
        "#    • Blank lines are ignored",
        "#    • Values must NOT be quoted",
        "#    • Paths may be absolute or relative to this file's location",
        "# ─────────────────────────────────────────────────────────────",
        "",
        "[sources]",
        f"data  = {data_path_hint or '/path/to/adnca.xpt'}",
        "sas   = ",
        "log   = ",
        "specs = ",
        "",
        "[options]",
        "# Phase filter: ALL (default), PHASE 1, or PHASE 2",
        "phase = ALL",
        "# Set yes = true to skip the final confirmation prompt",
        "yes   = false",
        "",
        "[mapping]",
        "# Map canonical variable names to actual column names in your dataset.",
        "# Delete or comment out lines where auto-detection is correct.",
        "# Format:  CANONICAL_NAME = ACTUAL_COLUMN_IN_DATASET",
        "#",
    ]

    req_vars = [(c, r) for c, _, r in VARIABLE_REGISTRY if r]
    opt_vars = [(c, r) for c, _, r in VARIABLE_REGISTRY if not r]

    lines.append("# — Required variables —")
    for canonical, _ in req_vars:
        detected = auto_mapping.get(canonical, "")
        tag      = f"  # auto-detected" if detected else "  # ⚠ not found — please set manually"
        lines.append(f"{canonical:<12} = {detected or ''}{tag}")

    lines.append("#")
    lines.append("# — Optional variables (remove # to activate) —")
    for canonical, _ in opt_vars:
        detected = auto_mapping.get(canonical, "")
        lines.append(f"# {canonical:<10} = {detected or ''}")

    out_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(green(f"\n  ✓ Config template written: {out_path}"))
    if columns:
        n_det = sum(1 for v in auto_mapping.values() if v)
        print(f"    {n_det} / {len(auto_mapping)} variables auto-detected from dataset.")
    print(grey("    Edit the file, then run:"))
    print(grey(f"    python adpc_output_qc_run.py --input {out_path}\n"))


# ═════════════════════════════════════════════════════════════════════════════
# Shared pipeline runner  (called from both modes)
# ═════════════════════════════════════════════════════════════════════════════

def _build_confirmed_mapping(auto_mapping: dict, overrides: dict,
                              columns: list, skip_interactive: bool):
    """
    Merge auto-detected mapping with cfg [mapping] overrides (or interactive input).
    Returns (confirmed_mapping, user_override_misses).
    """
    col_upper   = {c.upper() for c in columns}
    confirmed   = {}
    misses      = []

    for canonical, aliases, required in VARIABLE_REGISTRY:
        # Priority: explicit override > auto-detected > None
        if canonical in overrides:
            actual = overrides[canonical]
            if actual in col_upper:
                confirmed[canonical] = actual
            else:
                confirmed[canonical] = None
                misses.append((canonical, actual))
        else:
            confirmed[canonical] = auto_mapping.get(canonical)

    if not skip_interactive:
        # Interactive confirmation step (only used in interactive mode)
        print(f"\n{bold('Step 3 — Variable mapping')}")
        print(grey("  Auto-detected from dataset. Press Enter to accept, or type a correction.\n"))

        req_vars = [(c, a, r) for c, a, r in VARIABLE_REGISTRY if r]
        opt_vars = [(c, a, r) for c, a, r in VARIABLE_REGISTRY if not r]

        def _handle(canonical, required):
            current = confirmed.get(canonical)
            tag     = bold("[required]") if required else grey("[optional]")
            status  = green(f"✅ {current}") if current else yellow("— not found")
            raw     = input(f"  {tag} {canonical:<30} {status}: ").strip()
            if not raw:
                return  # keep current
            if raw.upper() in col_upper:
                confirmed[canonical] = raw.upper()
                print(f"    {green('✓')} mapped to {cyan(raw.upper())}")
            else:
                confirmed[canonical] = None
                misses.append((canonical, raw))
                print(f"    {red('✗')} '{raw}' not found — logged as QC finding")

        print(bold("  — Required variables —"))
        for c, a, r in req_vars:
            _handle(c, r)
        print()
        if _confirm("Review optional variable mappings?", default_yes=False):
            print(bold("\n  — Optional variables —"))
            for c, a, r in opt_vars:
                _handle(c, r)

    return confirmed, misses


def run_pipeline(data_path: Path, sas_path: str, log_path: str, specs_path: str,
                 phase: str, confirmed_mapping: dict, user_override_misses: list,
                 skip_confirm: bool):
    """Execute Steps A → D and print the final summary."""

    out_dir  = data_path.parent / ("adnca_review_" + datetime.now().strftime("%Y-%m-%dT%H_%M"))
    out_dir.mkdir(parents=True, exist_ok=True)
    print(f"\n  {green('✓')} Output directory: {cyan(str(out_dir))}")

    json_out = out_dir / "qc_output_results.json"
    xlsx_out = out_dir / "adpc_output_qc.xlsx"
    md_out   = out_dir / "qc_output_report.md"
    html_out = out_dir / "qc_output_report.html"

    # ── Review summary ────────────────────────────────────────────────────────
    print(); print(bold("Review and confirm"))
    print(f"  Dataset      : {cyan(str(data_path))}")
    print(f"  SAS program  : {cyan(sas_path) if sas_path else yellow('(not provided)')}")
    print(f"  SAS log      : {cyan(log_path)  if log_path  else yellow('(not provided — log analysis skipped)')}")
    print(f"  Specs file   : {cyan(specs_path) if specs_path else yellow('(not provided)')}")
    print(f"  Phase filter : {cyan(phase)}")
    print(); print("  Output files:")
    for p in (json_out, xlsx_out, md_out, html_out):
        print(f"    {cyan(str(p))}")

    if user_override_misses:
        print()
        print(yellow(f"  ⚠ {len(user_override_misses)} mapping override(s) not found in dataset:"))
        for canon, provided in user_override_misses:
            print(yellow(f"      {canon} → '{provided}' (will appear as QC finding)"))

    print()
    if not skip_confirm:
        if not _confirm("Run the QC pipeline now?", default_yes=True):
            print(yellow("\n  Aborted.\n")); sys.exit(0)

    # ── Serialise run config ──────────────────────────────────────────────────
    run_config = {
        "data_path":            str(data_path),
        "sas_path":             sas_path   or "",
        "log_path":             log_path   or "",
        "specs_path":           specs_path or "",
        "phase_filter":         phase,
        "mapping":              {k: v for k, v in confirmed_mapping.items() if v},
        "user_override_misses": [{"canonical": c, "provided": p}
                                 for c, p in user_override_misses],
    }
    config_path = out_dir / "_adpc_output_qc_config.json"
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(run_config, f, indent=2)

    # ── Copy input files into review folder ───────────────────────────────────
    print(bold("\nCopying input files into review folder …"))
    for src, label in [(str(data_path), "Dataset "), (sas_path, "SAS prog"),
                        (specs_path, "Specs   "), (log_path, "SAS log ")]:
        if src:
            try:
                shutil.copy2(src, str(out_dir / Path(src).name))
                print(f"  {green('✓')} {label}: {cyan(str(out_dir / Path(src).name))}")
            except Exception as e:
                print(yellow(f"  ⚠ Could not copy {label}: {e}"))

    # ── Locate sibling scripts ────────────────────────────────────────────────
    agent_script  = _find_script("adpc_output_qc_agent.py")
    report_script = _find_script("adpc_output_qc_report.py")
    agent_cmd  = [sys.executable, str(agent_script),
                  "--config", str(config_path), "--out", str(json_out), "--xlsx", str(xlsx_out)]
    report_cmd = [sys.executable, str(report_script),
                  "--json", str(json_out), "--out_dir", str(out_dir)]

    # ── Step A — QC agent ────────────────────────────────────────────────────
    ok = _run_step("Step A — Running QC agent (adpc_output_qc_agent.py)", agent_cmd)
    if not ok:
        print(red("\n  Pipeline stopped.\n")); sys.exit(1)
    if not json_out.exists():
        print(red(f"\n  ✗ Expected JSON not found: {json_out}")); sys.exit(1)

    # ── Step C — SAS log analysis ─────────────────────────────────────────────
    if log_path:
        print(f"\n{bold('─' * 60)}")
        print(bold("Step C — Analysing SAS log for ERRORs and WARNINGs"))
        print(bold("─" * 60))
        log_results = analyse_sas_log(log_path)
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
            print(yellow(f"  ⚠ Could not merge log results: {_e}"))
    else:
        print(f"\n  {grey('(Step C skipped — no SAS log provided)')}")

    # ── Step B — Report ──────────────────────────────────────────────────────
    ok = _run_step("Step B — Generating report (adpc_output_qc_report.py)", report_cmd)
    if not ok:
        print(red("\n  Report generation failed — JSON and XLSX are intact."))
        print(f"  Re-run:\n    python {report_script} --json {json_out}\n")
        sys.exit(1)

    # ── Step D — Append log sheets to XLSX ───────────────────────────────────
    if log_path and xlsx_out.exists():
        print(f"\n{bold('─' * 60)}")
        print(bold("Step D — Appending SAS log sheets to XLSX"))
        print(bold("─" * 60))
        append_log_xlsx(xlsx_out, log_path)

    # ── Done ─────────────────────────────────────────────────────────────────
    print()
    print(bold("╔══════════════════════════════════════════════════════╗"))
    print(bold("║                     ✅  Done                         ║"))
    print(bold("╚══════════════════════════════════════════════════════╝"))
    print()
    print("  Output files:")
    for p in (json_out, xlsx_out, md_out, html_out):
        icon = green("✅") if p.exists() else red("✗ missing")
        print(f"    {icon}  {cyan(str(p))}")
    if log_path:
        _lr   = json.loads(json_out.read_text(encoding="utf-8")).get("sas_log_analysis", {})
        _ne   = _lr.get("n_errors",          0)
        _nw   = _lr.get("n_warnings_review", 0)
        _icon = red("🔴") if _ne else (yellow("⚠️") if _nw else green("✅"))
        print(f"\n  SAS log  : {_icon}  {_ne} ERROR(s), {_nw} WARNING(s) to review "
              f"— see Section O in the HTML report")
    print()


# ═════════════════════════════════════════════════════════════════════════════
# Entry point
# ═════════════════════════════════════════════════════════════════════════════

def main():
    cli = argparse.ArgumentParser(
        description="Launcher for adpc_output_qc pipeline",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    cli.add_argument("--input",           default=None,
                     metavar="FILE.cfg",
                     help="Config file (.cfg) with all four source paths and options")
    cli.add_argument("--generate-config", default=None,
                     metavar="OUT.cfg",
                     help="Write a pre-filled config template and exit")
    cli.add_argument("--data",  default=None, help="Dataset path (overrides config; also used by --generate-config)")
    cli.add_argument("--yes",   action="store_true", help="Skip confirmation prompts")
    args = cli.parse_args()

    print()
    print(bold("╔══════════════════════════════════════════════════════╗"))
    print(bold("║     ADPC / ADNCA Output Data QC — Launcher           ║"))
    print(bold("╚══════════════════════════════════════════════════════╝"))
    print()

    # ── --generate-config mode ────────────────────────────────────────────────
    if args.generate_config:
        out_cfg = Path(args.generate_config).expanduser().resolve()
        print(bold(f"Generating config template → {out_cfg}"))
        generate_config(out_cfg, data_path_hint=args.data or "")
        sys.exit(0)

    # ── Config-file mode ──────────────────────────────────────────────────────
    if args.input:
        cfg_path = Path(args.input).expanduser().resolve()
        print(bold(f"Config-file mode — reading: {cyan(str(cfg_path))}"))
        print()
        cfg = load_cfg(cfg_path)

        # --data on command line overrides cfg
        if args.data:
            cfg["data_path"] = str(Path(args.data).expanduser().resolve())
        # --yes on command line overrides cfg
        if args.yes:
            cfg["yes"] = True

        print(bold("Validating paths …"))
        if not _validate_cfg_paths(cfg):
            sys.exit(1)

        data_path  = Path(cfg["data_path"])
        sas_path   = cfg["sas_path"]
        log_path   = cfg["log_path"]
        specs_path = cfg["specs_path"]
        phase      = cfg["phase"] or "ALL"
        skip_all   = cfg["yes"]
        overrides  = cfg["mapping_overrides"]

        print()
        print(bold("Reading dataset column names …"))
        columns      = _read_column_names(str(data_path))
        auto_mapping = _autodetect_mapping(columns) if columns else \
                       {c: None for c, _, _ in VARIABLE_REGISTRY}

        if columns:
            n_det     = sum(1 for v in auto_mapping.values() if v)
            req_found = sum(1 for c, _, r in VARIABLE_REGISTRY if r and auto_mapping.get(c))
            n_req     = sum(1 for _, _, r in VARIABLE_REGISTRY if r)
            print(f"  {green('✓')} {len(columns)} columns — "
                  f"{green(str(n_det))} / {len(auto_mapping)} variables auto-detected "
                  f"({green(str(req_found))} / {n_req} required)")
            req_miss = [c for c, _, r in VARIABLE_REGISTRY if r and not auto_mapping.get(c)
                        and c not in overrides]
            if req_miss:
                print(yellow(f"  ⚠ Required variables not found and not in [mapping]: "
                             f"{', '.join(req_miss)}"))

        # Merge auto-detection with cfg overrides (no interactive prompts)
        confirmed_mapping, misses = _build_confirmed_mapping(
            auto_mapping, overrides, columns, skip_interactive=True
        )

        if misses:
            print()
            print(yellow(f"  ⚠ {len(misses)} [mapping] override(s) not found in dataset:"))
            for canon, provided in misses:
                print(yellow(f"      {canon} → '{provided}'"))

        run_pipeline(data_path, sas_path, log_path, specs_path, phase,
                     confirmed_mapping, misses, skip_confirm=skip_all)

    # ── Interactive mode ──────────────────────────────────────────────────────
    else:
        print(bold("Interactive mode"))
        print(grey("  Tip: run with --generate-config adnca_qc.cfg --data path/to/adnca.xpt"))
        print(grey("       to create a config file and skip this wizard next time."))
        print()

        # Step 1 — dataset
        print(bold("Step 1 — ADPC / ADNCA dataset"))
        data_path = Path(
            args.data or _prompt("Dataset path", required=True)
        ).expanduser().resolve()
        if not data_path.is_file():
            print(red(f"\n  ✗ File not found: {data_path}")); sys.exit(1)
        print(f"  {green('✓')} Found: {cyan(str(data_path))}")

        # Step 2 — SAS program
        print(); print(bold("Step 2 — Production SAS program"))
        print(grey("  Leave blank to skip (flag lists will use defaults)."))
        sas_raw  = _prompt("Path to adnca.sas (or blank to skip)", default="")
        sas_path = ""
        if sas_raw:
            sp = Path(sas_raw).expanduser().resolve()
            sas_path = str(sp) if sp.exists() else ""
            print((green("  ✓ SAS program found") if sas_path
                   else yellow(f"  ⚠ Not found: {sp} — using defaults")))

        # Step 2.5 — specs
        print(); print(bold("Step 2.5 — Specifications file (Excel)"))
        print(grey("  Leave blank to skip."))
        specs_raw  = _prompt("Path to specifications Excel (or blank to skip)", default="")
        specs_path = ""
        if specs_raw:
            xp = Path(specs_raw).expanduser().resolve()
            if xp.exists() and xp.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
                specs_path = str(xp); print(green("  ✓ Specs file found"))
            else:
                print(yellow(f"  ⚠ Not found or not an Excel file — skipping"))

        # Step 2.7 — log
        print(); print(bold("Step 2.7 — SAS log file"))
        print(grey("  Leave blank to skip."))
        _auto_log = ""
        for _dir in [data_path.parent,
                     Path(sas_path).parent if sas_path else None]:
            if _dir is None: continue
            for _stem in [data_path.stem, "adnca"]:
                _cand = _dir / f"{_stem}.log"
                if _cand.exists():
                    _auto_log = str(_cand); break
            if _auto_log: break
        log_raw  = _prompt("Path to SAS log (or blank to skip)", default=_auto_log or "")
        log_path = ""
        if log_raw:
            lp = Path(log_raw).expanduser().resolve()
            log_path = str(lp) if lp.exists() else ""
            print((green("  ✓ Log file found") if log_path
                   else yellow(f"  ⚠ Not found: {lp} — log analysis skipped")))

        # Step 3 — variable mapping
        print(); print(bold("Step 3 — Reading dataset column names …"))
        columns      = _read_column_names(str(data_path))
        auto_mapping = _autodetect_mapping(columns) if columns else \
                       {c: None for c, _, _ in VARIABLE_REGISTRY}
        if columns:
            n_det     = sum(1 for v in auto_mapping.values() if v)
            req_found = sum(1 for c, _, r in VARIABLE_REGISTRY if r and auto_mapping.get(c))
            n_req     = sum(1 for _, _, r in VARIABLE_REGISTRY if r)
            print(f"  {green('✓')} {len(columns)} columns — "
                  f"{green(str(n_det))} / {len(auto_mapping)} auto-detected "
                  f"({green(str(req_found))} / {n_req} required)")
            req_miss = [c for c, _, r in VARIABLE_REGISTRY if r and not auto_mapping.get(c)]
            if req_miss:
                print(yellow(f"  ⚠ Required not found: {', '.join(req_miss)}"))

        confirmed_mapping, misses = _build_confirmed_mapping(
            auto_mapping, {}, columns, skip_interactive=False
        )

        # Step 4 — options
        print(); print(bold("Step 4 — Options"))
        phase = _prompt("Phase filter (blank for ALL)", default="ALL")

        run_pipeline(data_path, sas_path, log_path, specs_path, phase,
                     confirmed_mapping, misses, skip_confirm=args.yes)


if __name__ == "__main__":
    main()
