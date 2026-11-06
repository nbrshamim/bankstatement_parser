"""Unified parser for Bangladeshi bank statement PDFs.

The parser does not rely on one bank's layout. It finds the amount columns
(Debit / Credit / Balance) from the table header and uses the x-coordinate of
each word to decide which column an amount belongs to.

Only rows that contain a transaction date are used for amounts. Description
and reference lines are never scanned for money, which is what used to turn
account / reference numbers (for example 1161010049849) into huge balances.
"""

from __future__ import annotations

import difflib
import hashlib
import io
import os
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from typing import BinaryIO

import pdfplumber

# Bump this whenever parsing logic changes. The Streamlit app includes it in
# the cache key, so a code change can never be masked by an old cached result.
PARSER_VERSION = "2026.11.05.1"

# ---------------------------------------------------------------------------
# Patterns
# ---------------------------------------------------------------------------

# 01-01-2026, 01/01/26, 01-JAN-26, 15 MAR 2026, 2026-01-01
_DATE_TOKEN = (
    r"(?:"
    r"\d{1,2}[-/.](?:\d{1,2}|[A-Za-z]{3})[-/.]\d{2,4}"
    r"|\d{1,2}\s+[A-Za-z]{3}\s+\d{2,4}"
    r"|\d{4}[-/.]\d{1,2}[-/.]\d{1,2}"
    r")"
)

DATE_RE = re.compile(rf"^(?P<date>{_DATE_TOKEN})")

# Some statements (for example Standard Bank) print a serial number before
# the date: "1 07/12/09 CS DEP ...".
SERIAL_DATE_RE = re.compile(
    rf"^(?P<serial>\d{{1,3}})\s+(?P<date>{_DATE_TOKEN})"
)

TIME_RE = re.compile(
    r"^\s*\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM)?\b", re.IGNORECASE
)

# 1,234.56 / 1234.56 / (1,234.56) / -1,234.56 / 1,234.56CR
# Indian grouping (1,23,456.78) is accepted as well.
MONEY_RE = re.compile(
    r"^\(?-?(?:"
    r"\d{1,3}(?:,\d{3})+(?:\.\d{2})?"
    r"|\d{1,2}(?:,\d{2})+,\d{3}(?:\.\d{2})?"
    r"|\d+\.\d{2}"
    r")\)?(?:CR|DR)?$",
    re.IGNORECASE,
)

OPENING_KEYWORDS = (
    "OPENING BALANCE",
    "OPENING BAL",
    "BALANCE AT PERIOD START",
    "BALANCE AT START",
    "STATEMENT OPENING",
    "BROUGHT FORWARD",
    "BROUGHT FWD",
    "BALANCE B/F",
    "B/F BALANCE",
)

SUMMARY_RE = re.compile(
    r"^(?:STATEMENT\s+)?(?:OPENING|CLOSING)\s+BALANCE\b"
    r"|^TOTAL\b"
    r"|^GRAND\s+TOTAL\b"
    r"|^CUMULATIVE\s+TOTAL\b"
    r"|^UNCOLLECTED\b"
    r"|^SUMMARY\b"
    r"|^DEBITS\b"
    r"|^CREDITS\b"
    r"|^DR\s*COUNT\b"
    r"|^CR\s*COUNT\b"
    r"|^\*\s*="
    r"|^-+\s*END\b"
    r"|^IMPORTANT\s+NOTICE\b"
    r"|^ALL\s+TRANSACTIONS\b"
    r"|^AUTHORIZED\b"
    r"|^CONFIDENTIAL\b"
    r"|^STATEMENT\s+GENERATED\b",
    re.IGNORECASE,
)

# Text that marks the start of a statement summary / footer. If one of these
# appears inside a transaction description, everything from that point on is
# dropped.
TRIM_MARKERS = (
    "STATEMENT CLOSING BALANCE",
    "CLOSING BALANCE",
    "OPENING BALANCE",
    "OPENING BAL ",
    "TOTAL DEBIT",
    "TOTAL CREDIT",
    "TOTAL DEBITS",
    "TOTAL CREDITS",
    "TOTAL WITHDRAWAL",
    "TOTAL DEPOSIT",
    "DEBITS ",
    "CREDITS ",
    "DR COUNT",
    "CR COUNT",
    "UNCOLLECTED FUNDS",
    "END OF STATEMENT",
    "IMPORTANT NOTICE",
    "ALL TRANSACTIONS POSTED",
    "AUTHORIZED CONFIDENTIAL",
    "STATEMENT GENERATED",
    "PLEASE NOTE THAT",
    "TRANSACTION CODE DESCRIPTIONS",
    "THANKS FOR BANKING",
    "THE CUSTOMER SHOULD EXAMINE",
    "BAL. AT END",
    "BALANCE AT END",
    "PRINTED BY",
)

DEBIT_HEADERS = {"DEBIT", "DEBITS", "WITHDRAW", "WITHDRAWAL", "WITHDRAWALS", "DR"}
CREDIT_HEADERS = {"CREDIT", "CREDITS", "DEPOSIT", "DEPOSITS", "CR"}
BALANCE_HEADERS = {"BALANCE", "BALANCES"}

FOOTER_RE = re.compile(
    r"(page\s+\d+\s+of\s+\d+|print(ed)?\s*date|this is a computer|"
    r"end of statement|statement generated|\*+\s*end)",
    re.IGNORECASE,
)


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------


@dataclass
class BankInfo:
    bank_name: str = ""
    branch: str = ""
    account_no: str = ""
    account_holder: str = ""
    period: str = ""


@dataclass
class Transaction:
    date: str
    description: str
    debit: Decimal | None = None
    credit: Decimal | None = None
    balance: Decimal | None = None
    page: int = 0
    is_opening: bool = False

    # Internal fields, not shown in the UI / Excel
    top: float = field(default=0.0, repr=False)
    parts: list[tuple[float, str]] = field(default_factory=list, repr=False)


@dataclass
class ParsedStatement:
    info: BankInfo
    transactions: list[Transaction]
    warnings: list[str] = field(default_factory=list)
    source_name: str = ""


@dataclass
class _Row:
    top: float
    words: list[dict]

    @property
    def text(self) -> str:
        return " ".join(w["text"] for w in self.words).strip()


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------


def _money(text: str) -> Decimal | None:
    """Convert a money-looking token to Decimal, otherwise return None."""
    token = text.strip()
    if not MONEY_RE.match(token):
        return None

    negative = token.startswith("(") or token.startswith("-")
    token = token.strip("()")
    token = re.sub(r"(CR|DR)$", "", token, flags=re.IGNORECASE)
    token = token.replace(",", "").replace(" ", "").lstrip("-")
    try:
        value = Decimal(token)
    except InvalidOperation:
        return None
    return -value if negative else value


def _group_rows(words: list[dict], tolerance: float = 3.0) -> list[_Row]:
    """Group words that sit on the same visual line."""
    rows: list[_Row] = []
    for word in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if rows and abs(word["top"] - rows[-1].top) <= tolerance:
            rows[-1].words.append(word)
        else:
            rows.append(_Row(top=word["top"], words=[word]))
    for row in rows:
        row.words.sort(key=lambda w: w["x0"])
    return rows


def _leading_date(text: str) -> str | None:
    """Return the date at the start of a transaction line, if any.

    Supports an optional leading serial number (Standard Bank) and the
    ``DD MON YYYY`` form used by Agrani Bank.
    """
    stripped = text.lstrip()
    match = DATE_RE.match(stripped) or SERIAL_DATE_RE.match(stripped)
    return match.group("date") if match else None


def _row_date(
    row: _Row,
    learned_x_ratio: float | None = None,
    page_width: float = 1.0,
) -> tuple[str | None, int | None]:
    """Find a date at its learned column, falling back to the leading date."""
    if learned_x_ratio is not None:
        target_x = learned_x_ratio * page_width
        candidates: list[tuple[float, str, int]] = []
        for index, word in enumerate(row.words):
            match = DATE_RE.match(word["text"]) or SERIAL_DATE_RE.match(word["text"])
            if match:
                center = (word["x0"] + word["x1"]) / 2
                candidates.append(
                    (abs(center - target_x), match.group("date"), index)
                )
        if candidates:
            distance, date_value, index = min(candidates)
            if distance <= page_width * 0.12:
                return date_value, index

    date_value = _leading_date(row.text)
    if date_value is None:
        return None, None
    for index, word in enumerate(row.words):
        if DATE_RE.match(word["text"]) or SERIAL_DATE_RE.match(word["text"]):
            return date_value, index
    return date_value, None


def _row_fingerprint(text: str) -> str:
    tokens = re.findall(r"[A-Z]{2,}", text.upper())
    if not tokens:
        return ""
    return hashlib.sha256(" ".join(tokens).encode("utf-8")).hexdigest()


def _strip_leading_dates(text: str) -> str:
    """Remove the transaction date, time and value date from a description."""
    stripped = text.lstrip()
    match = DATE_RE.match(stripped) or SERIAL_DATE_RE.match(stripped)
    if not match:
        return stripped

    rest = stripped[match.end():].lstrip(" -:,")
    rest = TIME_RE.sub("", rest).lstrip(" -:,")

    # Many statements repeat the value date after the transaction date.
    value_match = DATE_RE.match(rest)
    if value_match:
        rest = rest[value_match.end():].lstrip(" -:,")
        rest = TIME_RE.sub("", rest).lstrip(" -:,")

    return rest


def _find_columns(rows: list[_Row]) -> tuple[dict[str, tuple[float, float]], float] | None:
    """Locate the amount columns from a table header.

    Returns ``({"debit": (x0, x1), "credit": (x0, x1), "balance": (x0, x1)},
    header_top)``. A column's span (not just one x coordinate) is kept because
    amount values are right-aligned inside the header cell and can sit a little
    outside the printed header text.
    """
    best: tuple[tuple[int, int], dict[str, tuple[float, float]], float] | None = None

    for row in rows:
        upper = row.text.upper()
        # "Statement Summary ... Total Debits Total Credits Balance" is not the
        # table header even though it contains the same words.
        if "SUMMARY" in upper:
            continue

        columns: dict[str, tuple[float, float]] = {}
        for index, word in enumerate(row.words):
            # "Dr/Cr" is a direction column, not the debit amount column.
            if "/" in word["text"]:
                continue
            match = re.match(r"[A-Z]+", word["text"].upper())
            if not match:
                continue
            text = match.group(0)
            if text in DEBIT_HEADERS:
                key = "debit"
            elif text in CREDIT_HEADERS:
                key = "credit"
            elif text in BALANCE_HEADERS:
                key = "balance"
            else:
                continue

            x0, x1 = word["x0"], word["x1"]
            # Merge a following "Amount" / "Amt" / "BDT" into the same
            # column. This fixes headers such as "Debit Amount".
            next_index = index + 1
            while next_index < len(row.words):
                following = re.match(
                    r"[A-Z]+", row.words[next_index]["text"].upper()
                )
                if not following:
                    break
                token = following.group(0)
                if token not in ("AMOUNT", "AMT", "BDT"):
                    break
                if row.words[next_index]["x0"] - x1 > 25:
                    break
                x1 = row.words[next_index]["x1"]
                next_index += 1
            # Keep the widest matching header ("Debit Amount" beats "Debit").
            old = columns.get(key)
            if old is not None and (old[1] - old[0]) >= (x1 - x0):
                continue
            columns[key] = (x0, x1)

        score = len(columns)
        if "balance" not in columns or score < 2:
            continue

        has_context = any(
            word in upper
            for word in ("DATE", "DESCRIPTION", "PARTICULARS", "NARRATION", "TRANS")
        )
        priority = (1 if has_context else 0, score)
        if best is None or priority > best[0]:
            best = (priority, columns, row.top)

    if best is None:
        return None
    return best[1], best[2]


def _column_for(
    word: dict, columns: dict[str, tuple[float, float]], max_distance: float = 35.0
) -> str | None:
    """Return the amount column whose right edge is closest to a word."""
    best_key: str | None = None
    best_distance = max_distance
    for key, (_, x1) in columns.items():
        distance = abs(word["x1"] - x1)
        if distance <= best_distance:
            best_key, best_distance = key, distance
    return best_key


def _read_amounts(
    row: _Row, columns: dict[str, tuple[float, float]] | None
) -> tuple[dict[str, Decimal | None], set[int]]:
    """Read debit / credit / balance from one dated row.

    Only money-shaped tokens are considered, and a token must line up with an
    amount column. Returns the values and the indexes of amount words so they
    can be removed from the description.
    """
    values: dict[str, Decimal | None] = {
        "debit": None,
        "credit": None,
        "balance": None,
    }
    amount_indexes: set[int] = set()

    if columns:
        for index, word in enumerate(row.words):
            value = _money(word["text"])
            if value is None:
                continue
            key = _column_for(word, columns)
            if key is None:
                continue
            if values[key] is None:
                values[key] = value
                amount_indexes.add(index)
        return values, amount_indexes

    # Fallback for PDFs where the header could not be found: use the last
    # money-shaped token as balance, but still avoid plain integers such as
    # account numbers (they are not accepted by MONEY_RE).
    candidates: list[tuple[int, Decimal]] = []
    for index, word in enumerate(row.words):
        value = _money(word["text"])
        if value is not None:
            candidates.append((index, value))

    if candidates:
        amount_indexes = {index for index, _ in candidates}
        values["balance"] = candidates[-1][1]
        if len(candidates) >= 3:
            values["debit"] = candidates[-3][1]
            values["credit"] = candidates[-2][1]
        elif len(candidates) == 2:
            values["credit"] = candidates[-2][1]

    return values, amount_indexes


def _description(row: _Row, amount_indexes: set[int]) -> str:
    """Join the non-amount words of a row, dropping the leading date."""
    words = [
        word["text"]
        for index, word in enumerate(row.words)
        if index not in amount_indexes
    ]
    text = re.sub(r"\s{2,}", " ", " ".join(words)).strip(" -")
    return _strip_leading_dates(text)


def _is_opening(text: str) -> bool:
    upper = text.upper()
    # Footer totals sometimes repeat "Opening Balance" ("Total Withdrawal ...
    # Opening Balance"). They are summaries, not opening rows.
    if re.search(r"\b(?:TOTAL|SUMMARY|CLOSING)\b", upper):
        return False
    return any(keyword in upper for keyword in OPENING_KEYWORDS)


def _is_footer(text: str, top: float, page_height: float) -> bool:
    if FOOTER_RE.search(text):
        return True
    # Repeated bank line at the very bottom of the page
    if top > page_height * 0.9 and re.search(
        r"\b(BANK|PLC|LIMITED|LTD)\b", text.upper()
    ):
        return True
    return False


def _clean_description(
    parts: list[tuple[float, str]], keep_summary: bool = False
) -> str:
    seen: set[str] = set()
    cleaned: list[str] = []
    for _, text in sorted(parts, key=lambda item: item[0]):
        text = re.sub(r"\s{2,}", " ", text).strip()
        if not text:
            continue
        key = text.lower()
        if key in seen:
            continue
        seen.add(key)
        cleaned.append(text)

    description = " ".join(cleaned).strip(" -")
    if keep_summary or not description:
        return description

    upper = description.upper()
    cut = len(description)
    for marker in TRIM_MARKERS:
        position = upper.find(marker)
        if position > 0:
            cut = min(cut, position)
    description = description[:cut].rstrip(" -,:;")
    # Drop a trailing direction / stray single-letter fragment that the PDF
    # text layer produces ("... deducted. T", "CASH l", "... 004 r").
    description = re.sub(r"\s+(?:DR|CR)\.?$", "", description, flags=re.IGNORECASE)
    description = re.sub(r"\s+[A-Za-z]$", "", description)
    return description.rstrip(" -,:;")


# ---------------------------------------------------------------------------
# Bank name printed as a logo (OCR fallback)
# ---------------------------------------------------------------------------

# Some statements (for example AB Bank) print the bank name only as a logo
# image, so it never appears in the text layer. When the text based detection
# finds nothing, the logo area of page 1 is rendered and read with Tesseract.
KNOWN_BANK_NAMES: tuple[str, ...] = (
    "AB Bank",
    "Agrani Bank",
    "Al-Arafah Islami Bank",
    "Ansar VDP Unnayan Bank",
    "Bangladesh Krishi Bank",
    "Bank Asia",
    "Basic Bank",
    "BRAC Bank",
    "Citibank",
    "City Bank",
    "Commercial Bank of Ceylon",
    "Community Bank Bangladesh",
    "Dhaka Bank",
    "Dutch-Bangla Bank",
    "Eastern Bank",
    "EXIM Bank",
    "First Security Islami Bank",
    "Habib Bank",
    "HSBC",
    "IFIC Bank",
    "Islami Bank Bangladesh",
    "Jamuna Bank",
    "Mercantile Bank",
    "Midland Bank",
    "Modhumoti Bank",
    "Mutual Trust Bank",
    "National Bank",
    "NCC Bank",
    "NRB Bank",
    "One Bank",
    "Padma Bank",
    "Prime Bank",
    "Probashi Kallyan Bank",
    "Pubali Bank",
    "Rajshahi Krishi Unnayan Bank",
    "Rupali Bank",
    "Shahjalal Islami Bank",
    "Shimanto Bank",
    "Social Islami Bank",
    "Sonali Bank",
    "Southeast Bank",
    "Standard Bank",
    "Standard Chartered",
    "The Premier Bank",
    "Trust Bank",
    "United Commercial Bank",
    "Uttara Bank",
)

_BANK_SUFFIX_RE = re.compile(
    r"\b(?:LIMITED|LTD|PLC|P\.?L\.?C\.?|BANGLADESH)\b\.?", re.IGNORECASE
)

_TESSERACT_CACHE: str | None = None
_TESSERACT_CHECKED = False


def _tesseract_exe() -> str | None:
    """Locate the Tesseract binary, or return None if OCR is unavailable."""
    global _TESSERACT_CACHE, _TESSERACT_CHECKED
    if _TESSERACT_CHECKED:
        return _TESSERACT_CACHE

    _TESSERACT_CHECKED = True
    candidates = [
        os.environ.get("TESSERACT_CMD", ""),
        shutil.which("tesseract") or "",
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        "/usr/bin/tesseract",
        "/usr/local/bin/tesseract",
    ]
    for candidate in candidates:
        if candidate and os.path.isfile(candidate):
            _TESSERACT_CACHE = candidate
            break
    return _TESSERACT_CACHE


def _ocr_image(page, box: tuple[float, float, float, float], resolution: int = 400) -> str:
    """Render a region of a PDF page and return the text Tesseract reads."""
    exe = _tesseract_exe()
    if exe is None:
        return ""

    x0, top, x1, bottom = box
    x0 = max(0.0, min(x0, page.width - 1))
    x1 = max(x0 + 1, min(x1, page.width))
    top = max(0.0, min(top, page.height - 1))
    bottom = max(top + 1, min(bottom, page.height))

    path = ""
    try:
        image = page.crop((x0, top, x1, bottom)).to_image(resolution=resolution)
        handle, path = tempfile.mkstemp(suffix=".png")
        os.close(handle)
        image.save(path)

        chunks: list[str] = []
        for psm in ("6", "7"):
            result = subprocess.run(
                [exe, path, "stdout", "-l", "eng", "--psm", psm],
                capture_output=True,
                text=True,
                timeout=30,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            if result.stdout.strip():
                chunks.append(result.stdout)
        return "\n".join(chunks)
    except Exception:
        return ""
    finally:
        if path:
            try:
                os.unlink(path)
            except OSError:
                pass


def _compact(text: str) -> str:
    return re.sub(r"[^A-Z]", "", text.upper())


def _window_similarity(compact: str, key: str) -> float:
    """Compare ``key`` against the best matching window inside ``compact``.

    This tolerates OCR noise before / after the bank name ("Sy AB Bank") and
    still rejects a bare "BANK" that would otherwise fuzzy-match a short name.
    """
    if not compact or not key or len(compact) < len(key) - 1:
        return 0.0

    best = 0.0
    lowest = max(1, len(key) - 2)
    highest = min(len(compact), len(key) + 2)
    for size in range(lowest, highest + 1):
        for start in range(0, len(compact) - size + 1):
            window = compact[start : start + size]
            ratio = difflib.SequenceMatcher(None, window, key).ratio()
            if ratio > best:
                best = ratio
                if best == 1.0:
                    return best
    return best


def _match_known_bank(text: str) -> str:
    best_name = ""
    best_score = 0.0

    for raw_line in text.splitlines():
        line = _BANK_SUFFIX_RE.sub(" ", raw_line)
        compact = _compact(line)
        if len(compact) < 4:
            continue
        for name in KNOWN_BANK_NAMES:
            key = _compact(name)
            score = _window_similarity(compact, key)
            if score > best_score:
                best_name, best_score = name, score

    return best_name if best_score >= 0.85 else ""


def _bank_from_ocr_text(text: str) -> str:
    """Turn raw OCR output into a bank name, using a known list first."""
    known = _match_known_bank(text)
    if known:
        return known

    # Unknown bank: accept a short header line that ends with the word BANK,
    # e.g. "Sy AB Bank" -> "AB Bank".
    for line in text.splitlines():
        tokens = [t for t in line.strip().split() if re.fullmatch(r"[A-Za-z]+", t)]
        if not 2 <= len(tokens) <= 5:
            continue
        if tokens[-1].upper() not in ("BANK", "BANKING"):
            continue
        if any(len(token) < 2 for token in tokens):
            continue
        candidate = " ".join(
            token if len(token) <= 3 else token.capitalize() for token in tokens
        )
        if len(candidate) <= 40:
            return candidate

    return ""


def _detect_bank_from_logo(page) -> str:
    """Read the bank name from the logo / masthead area of the first page."""
    if _tesseract_exe() is None:
        return ""

    header_limit = page.height * 0.18
    boxes: list[tuple[float, float, float, float]] = []

    for image in page.images:
        width = image.get("x1", 0) - image.get("x0", 0)
        height = image.get("bottom", 0) - image.get("top", 0)
        if image.get("top", page.height) > header_limit:
            continue
        if 15 <= width <= 320 and 8 <= height <= 90:
            boxes.append(
                (
                    image["x0"] - 6,
                    image["top"] - 6,
                    image["x1"] + 6,
                    image["bottom"] + 6,
                )
            )

    # Fallback: the left half of the masthead, where most banks put the logo.
    boxes.append((0, 0, min(page.width, 320), header_limit))

    for box in boxes:
        name = _bank_from_ocr_text(_ocr_image(page, box))
        if name:
            return name
    return ""


# ---------------------------------------------------------------------------
# Metadata
# ---------------------------------------------------------------------------

# Labels that usually sit in the right-hand column of a statement header.
# Text after one of these belongs to another field, not to the value we want.
_RIGHT_LABELS = tuple(
    sorted(
        (
            "BALANCE AT PERIOD START",
            "STATEMENT GENERATE DATE",
            "ACCOUNT OPENING DATE",
            "ACCOUNT OPEN DATE",
            "GENERATION DATE/TIME",
            "GENERATION DATE",
            "Statement Generate Date",
            "Statement Period",
            "STATEMENT FOR THE PERIOD",
            "STATEMENT FROM",
            "Statement From",
            "PRINTING DATE",
            "Printing Date",
            "Print Date",
            "Print Time",
            "Account Number",
            "ACCOUNT NUMBER",
            "Account No.",
            "ACCOUNT NO.",
            "Account No",
            "ACCOUNT NO",
            "A/C No",
            "A/C Title",
            "ACCOUNT TITLE",
            "Account Title",
            "ACCOUNT NAME",
            "Account Name",
            "Account Holder",
            "ACCOUNT HOLDER",
            "Customer Name",
            "CUSTOMER NAME",
            "Customer ID",
            "CUSTOMER ID",
            "Customer Id",
            "Customer Type",
            "CUSTOMER TYPE",
            "Currency Name",
            "CURRENCY NAME",
            "Currency",
            "CURRENCY",
            "Branch Code",
            "BRANCH CODE",
            "Product Name",
            "PRODUCT NAME",
            "Territory Name",
            "TERRITORY NAME",
            "From Date",
            "FROM DATE",
            "To Date",
            "TO DATE",
            "Interest Rate",
            "INTEREST RATE",
            "A/C Open Date",
            "Open Date",
            "OPEN DATE",
            "Address",
            "ADDRESS",
            "Cust. ID",
            "CUST. ID",
            "Contact No",
            "CONTACT NO",
            "Routing Number",
            "ROUTING NUMBER",
            "Maturity Date",
            "MATURITY DATE",
            "Expiry Date",
            "EXPIRY DATE",
            "Rate of Int.",
            "LIMIT AMOUNT",
            "Limit Amount",
            "Old Account No",
            "OLD ACCOUNT NO",
            "Account Type",
            "ACCOUNT TYPE",
            "A/C Type",
            "A/C Status",
            "Account Status",
            "Status",
            "STATUS",
            "Telephone",
            "TELEPHONE",
            "Period From",
            "PERIOD FROM",
            "Period",
            "PERIOD",
        ),
        key=len,
        reverse=True,
    )
)

_RIGHT_LABEL_RE = re.compile(
    r"\s+(?=(?:" + "|".join(re.escape(label) for label in _RIGHT_LABELS) + r")\b)",
    re.IGNORECASE,
)


def _cut_right(text: str) -> str:
    """Drop the right-hand column from a header line."""
    match = _RIGHT_LABEL_RE.search(text)
    if match:
        text = text[: match.start()]
    return re.sub(r"\s+", " ", text).strip(" :-,")


def _header_region(lines: list[str]) -> list[str]:
    """Return the header lines that precede the transaction table."""
    for index, line in enumerate(lines):
        upper = line.upper()
        has_debit = any(word in upper for word in ("DEBIT", "WITHDRAW"))
        has_credit = any(word in upper for word in ("CREDIT", "DEPOSIT"))
        if has_debit and has_credit and "BALANCE" in upper:
            return lines[:index]

    # No table header found: stop at the first transaction line.
    for index, line in enumerate(lines):
        if _leading_date(line):
            return lines[:index]
    return lines


def _detect_bank_name(lines: list[str]) -> str:
    for line in lines:
        cleaned = _cut_right(line)
        known = _match_known_bank(cleaned)
        if known:
            return known

    for line in lines:
        cleaned = _cut_right(line)
        upper = cleaned.upper()
        if "BANK" not in upper or len(cleaned) > 100:
            continue
        if _leading_date(cleaned):
            continue
        if re.search(r"\d{1,3}(?:,\d{3})+\.\d{2}", cleaned):
            continue
        if any(
            word in upper
            for word in (
                "ACCOUNT",
                "STATEMENT",
                "BRANCH",
                "ONLINE",
                "CUSTOMER",
                "TELEPHONE",
                "CURRENCY",
            )
        ):
            continue
        return cleaned.replace("©", "").replace("®", "").strip()
    return ""


def _bank_from_filename(filename: str) -> str:
    """Best-effort bank name from a file name such as ``City.pdf``."""
    stem = os.path.splitext(os.path.basename(filename))[0]
    compact = _compact(stem)
    if len(compact) < 3:
        return ""

    aliases = {
        "AB": "AB Bank",
        "ABBANK": "AB Bank",
        "DBBL": "Dutch-Bangla Bank",
        "DUTCHBANGLA": "Dutch-Bangla Bank",
        "EBL": "Eastern Bank",
        "IBBL": "Islami Bank Bangladesh",
        "ISLAMIBANK": "Islami Bank Bangladesh",
        "BASIC": "Basic Bank",
        "EXIM": "EXIM Bank",
        "EXIMBANK": "EXIM Bank",
        "PRIME": "Prime Bank",
        "SOUTHEAST": "Southeast Bank",
        "STANDARD": "Standard Bank",
        "AGRANI": "Agrani Bank",
        "JAMUNA": "Jamuna Bank",
        "SONALI": "Sonali Bank",
        "CITY": "City Bank",
        "DHAKABANK": "Dhaka Bank",
    }
    if compact in aliases:
        return aliases[compact]

    known = _match_known_bank(stem)
    if known:
        return known

    for name in KNOWN_BANK_NAMES:
        key = _compact(name)
        if key.startswith(compact) or compact.startswith(key):
            return name
    return ""


def _looks_like_holder(text: str) -> bool:
    if not text or len(text) > 90:
        return False
    if re.search(r"\d{4,}", text) or "#" in text or "," in text:
        return False

    upper = text.upper()
    bad_words = (
        "BANK",
        "BRANCH",
        "STATEMENT",
        "ACCOUNT",
        "ADDRESS",
        "HOUSE",
        "ROAD",
        "SECTOR",
        "BLOCK",
        "VILLA",
        "FLOOR",
        "APARTMENT",
        "FLAT",
        "LANE",
        "AVENUE",
        "POST",
        "VILLAGE",
        "DIST",
        "CITY",
        "PHONE",
        "EMAIL",
        "MOBILE",
        "FAX",
        "TELEPHONE",
        "WEBSITE",
        "SWIFT",
        "ROUTING",
        "CUSTOMER",
        "PRODUCT",
        "CURRENCY",
        "PERIOD",
        "DATE",
        "STATUS",
        "PRINT",
        "GENERATED",
        "USER",
        "HEAD",
        "OFFICE",
        "TIME",
        "PAGE",
    )
    if any(word in upper for word in bad_words):
        return False
    if not re.search(r"[A-Za-z]{2}", text):
        return False
    return True


def _clean_holder(value: str) -> str:
    value = _cut_right(value)
    value = re.sub(r"\s+", " ", value).strip(" :-,")
    return value


def _extract_holder(lines: list[str]) -> str:
    # 1. Explicit labels such as "Account Title", "A/C Title" and "Name".
    label_re = re.compile(
        r"^\s*(?:ACCOUNT\s+TITLE|A/C\s+TITLE|ACCOUNT\s+NAME|ACCOUNT\s+HOLDER"
        r"|CUSTOMER\s+NAME|NAME)\s*(?:[:\-]\s*|\s+)(?P<value>.+)$",
        re.IGNORECASE,
    )
    for line in lines:
        match = label_re.match(line)
        if not match:
            continue
        value = _clean_holder(match.group("value"))
        if value and not re.search(r"\d{6,}", value):
            return value

    # 2. Two-column headers: the holder name is printed to the left of a
    #    label such as "Account No" / "Print Date" / "Customer ID".
    for line in lines:
        match = _RIGHT_LABEL_RE.search(line)
        if not match:
            continue
        value = line[: match.start()].strip(" :-,")
        if _looks_like_holder(value):
            return re.sub(r"\s+", " ", value)

    # 3. Standalone name block (for example EBL and Sonali print the name on
    #    its own line just below the address / branch block).
    for line in lines:
        value = _cut_right(line)
        if _looks_like_holder(value):
            return re.sub(r"\s+", " ", value)

    return ""


def _clean_branch(value: str) -> str:
    value = re.sub(r"\s+", " ", value).strip(" :-,")
    value = re.sub(r"^\d{2,6}\s*[-–]\s*", "", value)
    value = re.sub(r"\s*\.\s*$", "", value)
    return value.strip(" :-,")


def _strip_bank_prefix(text: str, bank_name: str) -> str:
    match = re.search(
        r"\bBANK\b(?:\s+(?:LIMITED|LTD|PLC\.?))?\s*", text, re.IGNORECASE
    )
    if match and match.start() <= 40:
        remainder = text[match.end():].strip(" ,-")
        if remainder:
            return remainder
    if bank_name:
        prefix = re.escape(bank_name)
        remainder = re.sub(rf"^{prefix}\s*", "", text, flags=re.IGNORECASE)
        if remainder != text:
            return remainder.strip(" ,-")
    return text


def _extract_branch(lines: list[str], bank_name: str) -> str:
    branch_re = re.compile(
        r"\bBRANCH(?:\s+NAME)?\s*[:\-]\s*(?P<value>.+)$", re.IGNORECASE
    )
    for line in lines:
        match = branch_re.search(line)
        if not match:
            continue
        value = _clean_branch(match.group("value"))
        if value:
            return value

    for line in lines:
        value = _cut_right(line)
        upper = value.upper()
        if not value or not re.search(r"\bBRANCH\b", upper):
            continue

        # "Jashore Branch Branch Code : 195"
        if "BRANCH CODE" in upper:
            match = re.match(r"(?P<name>.+?)\s+\bBRANCH\b", value, re.IGNORECASE)
            if match:
                return _clean_branch(match.group("name") + " Branch")

        if re.search(r"\bBRANCH\b\s*$", upper):
            value = _strip_bank_prefix(value, bank_name)
            return _clean_branch(value)

    # Dutch-Bangla mobile banking statements have no branch, but do print the
    # division that maintains the account.
    for line in lines:
        value = _cut_right(line)
        if re.search(r"\bMOBILE\s+BANKING\s+DIVISION\b", value, re.IGNORECASE):
            return _clean_branch(value)

    return ""


def _extract_info(first_page_text: str, filename: str) -> BankInfo:
    info = BankInfo()
    lines = [line.strip() for line in first_page_text.splitlines() if line.strip()]
    header_lines = _header_region(lines)

    info.bank_name = _detect_bank_name(header_lines)

    # Account number. Allow spaces / hyphens because many banks print
    # "0410 - 01 - 0002004" or "4021-111213-000".
    match = re.search(
        r"(?<!OLD\s)(?:ACCOUNT\s*(?:NUMBER|NO\.?)|A/C\s*(?:NO\.?)?)"
        r"\s*(?:/\s*CONTRACT\s*ID)?\s*[:\-]?\s*"
        r"([0-9][0-9\s\-]{7,30})",
        first_page_text,
        re.IGNORECASE,
    )
    if match:
        info.account_no = match.group(1).strip()
    else:
        match = re.search(r"(\d{10,20})", filename)
        if match:
            info.account_no = match.group(1)
    if info.account_no:
        digits = re.sub(r"\D", "", info.account_no)
        if digits:
            info.account_no = digits

    info.branch = _extract_branch(header_lines, info.bank_name)
    info.account_holder = _extract_holder(header_lines)

    # Period
    for line in header_lines:
        match = re.search(
            r"\b(?:PERIOD(?:\s+OF)?(?:\s+FROM)?|STATEMENT\s+FROM)\b"
            r"\s*[:\-]\s*(?P<value>.+)$",
            line,
            re.IGNORECASE,
        )
        if match:
            value = _cut_right(match.group("value"))
            value = re.sub(r"\s+", " ", value).strip(" :-")
            if value:
                info.period = value
                break

    return info


# ---------------------------------------------------------------------------
# Main parser
# ---------------------------------------------------------------------------


def parse_statement(
    source: str | BinaryIO,
    filename: str = "",
    learned_layout: dict | None = None,
) -> ParsedStatement:
    """Parse a bank statement PDF.

    ``source`` may be a path, a file-like object, or bytes.
    """
    if isinstance(source, (bytes, bytearray)):
        source = io.BytesIO(source)
    elif hasattr(source, "seek"):
        try:
            source.seek(0)
        except (OSError, ValueError):
            pass

    transactions: list[Transaction] = []
    warnings: list[str] = []
    info = BankInfo()
    columns: dict[str, tuple[float, float]] | None = None
    first_page_text = ""
    logo_bank_name = ""
    ocr_attempted = False

    with pdfplumber.open(source) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            words = page.extract_words(x_tolerance=2, y_tolerance=3)
            if not words:
                continue

            rows = _group_rows(words)
            if page_number == 1:
                first_page_text = page.extract_text() or " ".join(
                    row.text for row in rows
                )
                text_bank = _extract_info(first_page_text, filename).bank_name
                # Try the file name before OCR: several banks print the name
                # only as a logo image, and "City.pdf" is more reliable than a
                # failed OCR pass.
                if not text_bank and not _bank_from_filename(filename):
                    ocr_attempted = True
                    logo_bank_name = _detect_bank_from_logo(page)

            found = _find_columns(rows)
            header_top = 0.0
            geometry = (learned_layout or {}).get("transaction_geometry", {})
            learned_columns = geometry.get("columns", {})
            if learned_columns:
                learned_spans = {
                    key: (
                        float(span["x0_ratio"]) * page.width,
                        float(span["x1_ratio"]) * page.width,
                    )
                    for key, span in learned_columns.items()
                    if "x0_ratio" in span and "x1_ratio" in span
                }
                if learned_spans:
                    columns = learned_spans
                    header_top = float(
                        geometry.get("header_top_ratio", 0.0)
                    ) * page.height
            if found:
                detected_columns, detected_header_top = found
                if not learned_columns:
                    columns, header_top = detected_columns, detected_header_top
                elif not header_top:
                    header_top = detected_header_top
            elif columns is None:
                warnings.append(
                    "Could not detect the Debit/Credit/Balance columns; "
                    "used the fallback amount reader."
                )

            page_transactions: list[Transaction] = []
            pending: list[_Row] = []
            pending_opening: _Row | None = None
            footer_started = False

            for row in rows:
                text = row.text
                if not text:
                    continue
                if any(
                    exclusion.get("page") == page_number
                    and abs(
                        row.top / max(page.height, 1)
                        - float(exclusion.get("top_ratio", -1))
                    )
                    <= 0.02
                    and exclusion.get("fingerprint") == _row_fingerprint(text)
                    for exclusion in (learned_layout or {}).get("excluded_rows", [])
                ):
                    continue

                values, amount_indexes = _read_amounts(row, columns)
                date_value, date_index = _row_date(
                    row,
                    geometry.get("date_x_ratio"),
                    page.width,
                )
                if date_index is not None and _leading_date(text) is None:
                    amount_indexes.add(date_index)
                is_opening_row = _is_opening(text)
                has_balance = values["balance"] is not None
                has_movement = (
                    values["debit"] is not None or values["credit"] is not None
                )

                # An opening row may print its amount under the debit/credit
                # header instead of the balance header. Use the last money
                # token on the line as the opening balance in that case.
                if is_opening_row and not has_balance:
                    last_money: tuple[int, Decimal] | None = None
                    for index, word in enumerate(row.words):
                        value = _money(word["text"])
                        if value is not None:
                            last_money = (index, value)
                    if last_money is not None:
                        values["balance"] = last_money[1]
                        amount_indexes.add(last_money[0])
                        has_balance = True

                # Some layouts print "Opening Balance" on one line and the
                # amount on the next line (Jamuna). Remember the label and
                # combine it with the following money-only row.
                if is_opening_row and not has_balance and not has_movement:
                    if not page_transactions:
                        pending_opening = row
                        continue
                    pending_opening = None
                if pending_opening is not None:
                    if (
                        not date_value
                        and has_balance
                        and not has_movement
                        and row.top - pending_opening.top <= 20
                    ):
                        transaction = Transaction(
                            date="",
                            description="",
                            balance=values["balance"],
                            page=page_number,
                            is_opening=True,
                            top=pending_opening.top,
                        )
                        transaction.parts.append(
                            (pending_opening.top, _description(pending_opening, set()))
                        )
                        page_transactions.append(transaction)
                        pending_opening = None
                        continue
                    if row.top - pending_opening.top > 20:
                        pending_opening = None

                is_transaction = bool(
                    date_value and (has_balance or has_movement)
                )

                # Skip repeated page headers, but never skip a real dated
                # transaction or an opening row that sits above the header.
                if (
                    row.top <= header_top + 1
                    and not is_transaction
                    and not (is_opening_row and has_balance)
                ):
                    continue

                # A dated row with an amount is a real transaction. A
                # date-looking fragment with no amount is just a wrapped
                # description line.
                if is_transaction:
                    transaction = Transaction(
                        date=date_value or "",
                        description="",
                        debit=values["debit"],
                        credit=values["credit"],
                        balance=values["balance"],
                        page=page_number,
                        is_opening=is_opening_row,
                        top=row.top,
                    )
                    transaction.parts.append((row.top, _description(row, amount_indexes)))
                    page_transactions.append(transaction)
                    continue

                # An undated opening-balance row (for example "Brought Forward")
                # belongs to the top of the page. A line after the first
                # transaction is a footer summary, not an opening row.
                if is_opening_row and has_balance and not page_transactions:
                    transaction = Transaction(
                        date="",
                        description="",
                        balance=values["balance"],
                        page=page_number,
                        is_opening=True,
                        top=row.top,
                    )
                    transaction.parts.append((row.top, _description(row, amount_indexes)))
                    page_transactions.append(transaction)
                    continue

                # Skip page footers and summary lines only after real
                # transaction rows have been recognised.
                if _is_footer(text, row.top, page.height):
                    continue
                if SUMMARY_RE.match(text):
                    continue

                # Separator / decoration rows mark the end of the table.
                if re.fullmatch(r"[-=_*.·\s]+", text):
                    if page_transactions:
                        footer_started = True
                    continue
                # Ignore stray one-character fragments (City prints "e", "s").
                if len(re.sub(r"[^A-Za-z0-9]", "", text)) < 2:
                    continue
                # Table captions (Jamuna prints "Date No" / "Value Ref/Cheque",
                # Dhaka prints "DATE" / "DATE BRN" above the first row).
                if re.fullmatch(
                    r"(?i)(?:DATE|DATE\s+BRN|DATE\s+NO|VALUE\s+DATE"
                    r"|VALUE\s+REF/?\s*CHEQUE(?:\s+NO)?|TXN\s+DATE|TRANS\s+DATE"
                    r"|CHEQUE\s+NO|REF(?:ERENCE)?\s+NO|BRN)",
                    text.strip(),
                ):
                    continue

                # Totals below the closing separator are not narration.
                if footer_started:
                    continue

                pending.append(row)

            # Attach description lines to the closest transaction on the page.
            # This handles both "description below the amount row" (DBBL) and
            # "description wrapped above the amount row" (IBBL).
            for row in pending:
                if not page_transactions:
                    continue
                # A continuation line just below a transaction belongs to
                # that transaction. Otherwise use the closest transaction.
                previous = next(
                    (
                        transaction
                        for transaction in reversed(page_transactions)
                        if transaction.top <= row.top
                    ),
                    None,
                )
                if previous is not None and row.top - previous.top <= 15:
                    target = previous
                else:
                    target = min(
                        page_transactions,
                        key=lambda transaction: abs(transaction.top - row.top),
                    )
                    # Do not attach far-away page furniture (column captions
                    # such as "Date No" / "Value Ref/Cheque").
                    if abs(target.top - row.top) > 35:
                        continue
                target.parts.append((row.top, _description(row, set())))

            for transaction in page_transactions:
                transaction.description = _clean_description(
                    transaction.parts, keep_summary=transaction.is_opening
                )
                if not transaction.description:
                    transaction.description = (
                        "Opening balance" if transaction.is_opening else "—"
                    )
                transactions.append(transaction)

    info = _extract_info(first_page_text, filename) if first_page_text else BankInfo()
    if not info.bank_name:
        info.bank_name = _bank_from_filename(filename)
    if not info.bank_name and logo_bank_name:
        info.bank_name = logo_bank_name
    if not info.bank_name and ocr_attempted:
        if _tesseract_exe() is None:
            warnings.append(
                "The bank name is only inside the logo image and Tesseract OCR "
                "was not found, so it could not be read. Install Tesseract or "
                "set the TESSERACT_CMD environment variable."
            )
        else:
            warnings.append(
                "The bank name is only inside the logo image and OCR could not "
                "read it."
            )

    _finalize(transactions, warnings)

    return ParsedStatement(
        info=info,
        transactions=transactions,
        warnings=warnings,
        source_name=filename,
    )


def _finalize(transactions: list[Transaction], warnings: list[str]) -> None:
    """Confirm Debit/Credit against the running balance and fix disagreements."""
    previous_balance: Decimal | None = None
    opening_balance: Decimal | None = None

    for transaction in transactions:
        # A closing line may print Debit/Credit but leave the balance blank
        # (Jamuna closes the account this way). Derive it from the previous
        # running balance so the statement still balances.
        if (
            transaction.balance is None
            and not transaction.is_opening
            and previous_balance is not None
            and (transaction.debit is not None or transaction.credit is not None)
        ):
            transaction.balance = (
                previous_balance
                + (transaction.credit or Decimal("0"))
                - (transaction.debit or Decimal("0"))
            )

        if transaction.balance is None:
            continue

        if transaction.is_opening:
            transaction.debit = None
            transaction.credit = None
            previous_balance = transaction.balance
            if opening_balance is None:
                opening_balance = transaction.balance
            continue

        if previous_balance is None:
            debit = transaction.debit or Decimal("0")
            credit = transaction.credit or Decimal("0")
        else:
            movement = transaction.balance - previous_balance
            printed = (transaction.credit or Decimal("0")) - (
                transaction.debit or Decimal("0")
            )

            if abs(movement - printed) <= Decimal("0.01"):
                debit = transaction.debit or Decimal("0")
                credit = transaction.credit or Decimal("0")
            elif movement > Decimal("0.01"):
                # The balance moved up: treat it as a credit.
                debit, credit = Decimal("0"), movement
            elif movement < Decimal("-0.01"):
                # The balance moved down: treat it as a debit.
                debit, credit = -movement, Decimal("0")
            else:
                debit = transaction.debit or Decimal("0")
                credit = transaction.credit or Decimal("0")

        transaction.debit = debit if debit != 0 else None
        transaction.credit = credit if credit != 0 else None
        previous_balance = transaction.balance

    if len(transactions) < 2:
        return

    first_balance = next(
        (t.balance for t in transactions if t.balance is not None), None
    )
    last_balance = next(
        (t.balance for t in reversed(transactions) if t.balance is not None), None
    )
    if first_balance is None or last_balance is None:
        return

    total_debit = sum((t.debit or Decimal("0")) for t in transactions)
    total_credit = sum((t.credit or Decimal("0")) for t in transactions)
    movement = last_balance - (opening_balance if opening_balance is not None else Decimal("0"))

    if abs((total_credit - total_debit) - movement) > Decimal("1.00"):
        warnings.append(
            "Debit/Credit totals do not fully reconcile with the change in "
            f"balance (difference {total_credit - total_debit - movement:,.2f})."
        )
