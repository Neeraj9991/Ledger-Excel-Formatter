"""
daybook_parser.py — Parses Tally daybook Excel and normalises into a flat DataFrame.

Excel layout (starting row 8, 0-indexed row 7):
  Col A  = Date
  Col B  = Particulars (ledger name on header row; cost-centre / creditor on sub-rows)
  Col C  = Sub-amount (Dr side, cost-centre amount)
  Col D  = Dr/Cr label for col C
  Col E  = Vch Type
  Col F  = Vch No.
  Col G  = Debit Amount  (total debit on ledger header row)
  Col H  = Credit Amount (on creditor sub-rows)

A TRANSACTION BLOCK starts when:
  - Col A has a date value (non-NaN), AND
  - Col B matches a ledger in the ledger set.

Output schema (one row per cost-centre OR creditor entry):
  Date | Ledger | Vch Type | Vch No. | Particulars | Dr/Cr | Amount | Type
  where Type = "Cost Centre" or "Creditor"
  Dr/Cr = "Dr" for cost-centre rows, "Cr" for creditor rows
"""

import re
import pandas as pd
from pathlib import Path


# ─── Entries that are never cost centres (GST, tax lines, rounding) ──────────

EXCLUDED_PARTICULARS = {
    "sgst input", "cgst input", "igst input",
    "sgst", "cgst", "igst",
    "roundoff", "round off", "round-off",
    "tds", "tcs",
}


# ─────────────────────────── helpers ────────────────────────────────────────

def load_ledger_set(ledger_path: str | Path) -> set[str]:
    """Return a set of ledger names (stripped, lowercase) for fast lookup."""
    p = Path(ledger_path)
    if not p.exists():
        return set()
    lines = p.read_text(encoding="utf-8", errors="ignore").splitlines()
    return {line.strip().lower() for line in lines if line.strip()}


def _clean(val) -> str:
    """Return stripped string or ''."""
    if val is None:
        return ""
    return str(val).strip()


def _to_float(val) -> float | None:
    """Best-effort numeric conversion; strip Dr/Cr suffix."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = re.sub(r"[Dd][Rr]|[Cc][Rr]|\s", "", str(val)).replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


# Explicit formats tried in priority order — avoids MM/DD ambiguity.
_DATE_FORMATS = [
    "%d-%b-%y",           # 09-Feb-26   ← Tally export default
    "%d-%b-%Y",           # 09-Feb-2026
    "%d/%b/%y",           # 09/Feb/26
    "%Y-%m-%d %H:%M:%S",  # 2026-02-09 00:00:00   ← openpyxl datetime→str
    "%Y-%m-%d",           # 2026-02-09
    "%d-%m-%Y",           # 09-02-2026
    "%d-%m-%y",           # 09-02-26
    "%d/%m/%Y",           # 09/02/2026
    "%d/%m/%y",           # 09/02/26
]

def _format_date(val: str) -> str:
    """
    Parse a date string using a list of explicit formats (DD-first)
    and return dd-mm-yyyy.  Falls back to the original string if nothing matches.
    """
    if not val:
        return val
    for fmt in _DATE_FORMATS:
        try:
            dt = pd.to_datetime(val, format=fmt, errors="raise")
            return dt.strftime("%d-%m-%Y")
        except (ValueError, TypeError):
            continue
    return val   # unrecognised — return as-is


def _is_excluded(name: str) -> bool:
    """Return True if the particulars name is a GST/roundoff line to skip."""
    return name.lower().strip() in EXCLUDED_PARTICULARS


# ─────────────────────────── main parser ────────────────────────────────────

def parse_daybook(file, ledger_path: str | Path = "ledger.txt") -> pd.DataFrame:
    """
    Parse the uploaded Excel daybook and return a normalised DataFrame.

    Parameters
    ----------
    file        : file-like object or path accepted by pd.read_excel
    ledger_path : path to ledger.txt

    Returns
    -------
    pd.DataFrame with columns:
        Date, Ledger, Vch Type, Vch No., Total Debit, Particulars, Amount, Type
    """
    ledgers = load_ledger_set(ledger_path)

    # Read raw — header is split across two merged rows (7 & 8 in Excel).
    raw = pd.read_excel(
        file,
        header=None,
        skiprows=7,          # skip rows 1-7 (Excel), data starts at row 8
        dtype=str,
        engine="openpyxl",
    )

    # Trim to cols A-H (0-7). Guard against sheets with fewer columns.
    max_col = min(8, raw.shape[1])
    raw = raw.iloc[:, :max_col]
    raw.columns = list("ABCDEFGH")[:max_col]
    for col in list("ABCDEFGH"):
        if col not in raw.columns:
            raw[col] = ""

    raw = raw.replace({"nan": "", "None": "", "NaN": ""}).fillna("")

    records = []
    current_date = None
    current_ledger = None
    current_vch_type = None
    current_vch_no = None

    # Accumulate sub-rows separately
    cost_centres: list[dict] = []   # {"name": str, "amount": float|None, "dr_cr": str}
    creditors: list[dict] = []      # {"name": str, "amount": float|None}

    def flush():
        """Emit flat records for the current transaction block."""
        if current_ledger is None:
            return

        common = {
            "Date": current_date,
            "Ledger": current_ledger,
            "Vch Type": current_vch_type,
            "Vch No.": current_vch_no,
        }

        if not cost_centres and not creditors:
            records.append({**common, "Particulars": "", "Debit/Credit": "", "Amount": None, "Type": ""})
            return

        for cc in cost_centres:
            records.append({**common,
                            "Particulars": cc["name"],
                            "Amount": cc["amount"],
                            "Debit/Credit": cc.get("dr_cr", "Dr"),
                            "Type": "Cost Centre"})

        for cr in creditors:
            records.append({**common,
                            "Particulars": cr["name"],
                            "Amount": cr["amount"],
                            "Debit/Credit": "Cr",
                            "Type": "Creditor"})

    # ── Row iteration ────────────────────────────────────────────────────────
    for _, row in raw.iterrows():
        col_a = _clean(row.get("A", ""))
        col_b = _clean(row.get("B", ""))
        col_c = _clean(row.get("C", ""))
        col_d = _clean(row.get("D", ""))
        col_e = _clean(row.get("E", ""))
        col_f = _clean(row.get("F", ""))
        col_g = _clean(row.get("G", ""))
        col_h = _clean(row.get("H", ""))

        # Skip completely empty rows.
        if not any([col_a, col_b, col_c, col_d, col_e, col_f, col_g, col_h]):
            continue

        # Skip footer / summary rows (e.g. "Total:", "Grand Total").
        if col_a.lower().startswith("total") or col_b.lower().startswith("total"):
            continue

        # Detect ledger header row: col_a has a date AND col_b is a known ledger.
        is_ledger_row = bool(col_b) and (col_b.lower() in ledgers)

        if is_ledger_row and col_a:
            flush()
            cost_centres = []
            creditors = []

            current_date = _format_date(col_a)
            current_ledger = col_b
            current_vch_type = col_e
            current_vch_no = col_f
            continue

        # Sub-rows below a ledger header.
        if current_ledger is None:
            continue

        # Propagate date if col_a has a value.
        if col_a:
            current_date = _format_date(col_a)

        # Skip GST / Roundoff lines entirely.
        if _is_excluded(col_b):
            continue

        # Cost-centre row: col_b = CC name, col_c = Dr amount, col_d = Dr/Cr tag.
        if col_b and col_c:
            # col_d normally contains "Dr" — fall back to "Dr" if blank.
            dr_cr_tag = col_d.strip() if col_d.strip() else "Dr"
            cost_centres.append({"name": col_b, "amount": _to_float(col_c), "dr_cr": dr_cr_tag})
            continue

        # Creditor row: col_b = creditor name, col_h = credit amount.
        if col_b and col_h:
            creditors.append({"name": col_b, "amount": _to_float(col_h)})
            continue

    # Flush the last block.
    flush()

    df = pd.DataFrame(records)
    if df.empty:
        return df

    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")

    return df
