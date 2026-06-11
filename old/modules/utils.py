"""
modules/utils.py
Shared helpers used across all adpc_output_qc modules.
"""

from pathlib import Path
import warnings

import pandas as pd

warnings.filterwarnings("ignore")

# ── Bad years that flag erroneous datetimes (mirrors adnca.sas) ──────────────
BAD_YEARS = ["1900", "1923"]

# ── Default flag lists (used when SAS parsing fails) ─────────────────────────
DEFAULT_PKSUMXFL_CRITS = [
    "CRIT1FL", "CRIT2FL", "CRIT3FL", "CRIT4FL", "CRIT5FL",
    "CRIT9FL", "CRIT10FL", "CRIT11FL", "CRIT12FL", "CRIT13FL",
    "CRIT14FL", "CRIT15FL", "CRIT16FL", "CRIT19FL", "CRIT20FL",
]
DEFAULT_NCAXFL_NCAXRS = [
    "NCA01XRS", "NCA02XRS", "NCA03XRS", "NCA04XRS",
    "NCA05XRS", "NCA07XRS", "NCA16XRS",
]


def load_dataset(path: str) -> pd.DataFrame:
    """Load a CSV / XPT / SAS7BDAT dataset and return a DataFrame with uppercased columns."""
    ext = Path(path).suffix.lower()
    if ext == ".csv":
        df = pd.read_csv(path, low_memory=False)
    elif ext in (".xpt", ".sas7bdat"):
        import pyreadstat
        reader = pyreadstat.read_xport if ext == ".xpt" else pyreadstat.read_sas7bdat
        df, _ = reader(path)
    else:
        raise ValueError(f"Unsupported file format: {ext}")
    df.columns = [c.upper() for c in df.columns]
    for c in df.select_dtypes(include="object").columns:
        first_valid = df[c].dropna().iloc[0] if df[c].notna().any() else None
        if first_valid is None or isinstance(first_valid, str):
            df[c] = df[c].str.strip()
    return df


def col(df: pd.DataFrame, canonical: str, mapping: dict) -> "str | None":
    """Return the actual column name for a canonical variable, or None."""
    mapped = mapping.get(canonical)
    if mapped and mapped.upper() in df.columns:
        return mapped.upper()
    return None


def is_blank(series: pd.Series) -> pd.Series:
    """True where a Series value is NaN, empty string, '.', 'nan', or 'NA'."""
    return series.isna() | series.astype(str).str.strip().isin(["", ".", "nan", "NA"])


def col_list(df: pd.DataFrame, names: list) -> list:
    """Return the subset of names that actually exist as columns in df."""
    return [n for n in names if n.upper() in df.columns]


def json_safe(obj):
    """Recursively make an object JSON-serialisable."""
    if isinstance(obj, dict):
        return {k: json_safe(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [json_safe(v) for v in obj]
    if isinstance(obj, float) and (obj != obj or abs(obj) == float("inf")):
        return None
    if hasattr(obj, "item"):
        return obj.item()
    if hasattr(obj, "isoformat"):
        return obj.isoformat()
    if isinstance(obj, pd.Series):
        return obj.tolist()
    return obj
