"""Local, supervised statement-layout matching model.

Reviewed layout signatures and corrected statement labels are persisted
locally. Original PDF files are not retained.
"""

from __future__ import annotations

import hashlib
import io
import json
import math
import os
import re
import sqlite3
import statistics
from collections import Counter
from contextlib import closing
from datetime import datetime, timezone
from pathlib import Path
from typing import NamedTuple

import pdfplumber

MODEL_VERSION = "2"

_FIELD_LABELS = {
    "ACCOUNT HOLDER": re.compile(
        r"\b(?:ACCOUNT\s+HOLDER|ACCOUNT\s+TITLE|A/C\s+TITLE|"
        r"CUSTOMER\s+NAME|ACCOUNT\s+NAME)\b",
        re.IGNORECASE,
    ),
    "ACCOUNT NUMBER": re.compile(
        r"\b(?:ACCOUNT\s*(?:NO\.?|NUMBER)|A/C\s*(?:NO\.?|NUMBER))\b",
        re.IGNORECASE,
    ),
    "BRANCH": re.compile(r"\bBRANCH(?:\s+NAME)?\b", re.IGNORECASE),
    "PERIOD": re.compile(
        r"\b(?:STATEMENT\s+PERIOD|PERIOD\s+FROM|PERIOD|STATEMENT\s+FROM)\b",
        re.IGNORECASE,
    ),
}
_TABLE_WORDS = {
    "DATE",
    "DESCRIPTION",
    "PARTICULARS",
    "NARRATION",
    "TRANSACTION",
    "DEBIT",
    "DEBITS",
    "WITHDRAWAL",
    "CREDIT",
    "CREDITS",
    "DEPOSIT",
    "BALANCE",
    "AMOUNT",
    "VALUE",
    "REFERENCE",
    "CHEQUE",
    "DR",
    "CR",
}


class LayoutMatch(NamedTuple):
    layout_label: str
    similarity: float
    example_count: int
    layout_template: dict


def training_database_path() -> Path:
    app_data = os.environ.get("LOCALAPPDATA")
    root = Path(app_data) if app_data else Path.home() / ".local" / "share"
    return root / "BankStatementParser" / "layout_training.sqlite3"


def _connect(database_path: str | os.PathLike[str] | None = None) -> sqlite3.Connection:
    path = (
        Path(database_path)
        if database_path is not None
        else training_database_path()
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(path)
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS reviewed_layouts (
            fingerprint TEXT PRIMARY KEY,
            signature TEXT NOT NULL,
            layout_label TEXT NOT NULL,
            review_payload TEXT NOT NULL DEFAULT '{}',
            layout_template TEXT NOT NULL DEFAULT '{}',
            reviewed_at TEXT NOT NULL,
            model_version TEXT NOT NULL
        )
        """
    )
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS review_metrics (
            id INTEGER PRIMARY KEY,
            reviewed_at TEXT NOT NULL,
            field_correct INTEGER NOT NULL,
            field_total INTEGER NOT NULL,
            row_true_positive INTEGER NOT NULL,
            row_false_positive INTEGER NOT NULL,
            row_false_negative INTEGER NOT NULL,
            transaction_value_correct INTEGER NOT NULL,
            transaction_value_total INTEGER NOT NULL
        )
        """
    )
    columns = {
        row[1] for row in connection.execute("PRAGMA table_info(reviewed_layouts)")
    }
    if "review_payload" not in columns:
        connection.execute(
            "ALTER TABLE reviewed_layouts "
            "ADD COLUMN review_payload TEXT NOT NULL DEFAULT '{}'"
        )
    if "layout_template" not in columns:
        connection.execute(
            "ALTER TABLE reviewed_layouts "
            "ADD COLUMN layout_template TEXT NOT NULL DEFAULT '{}'"
        )
    return connection


def _line_signature(text: str, words: list[dict], page_width: float) -> str | None:
    normalized = " ".join(text.upper().split())
    field_match = next(
        (
            (label, pattern.search(normalized))
            for label, pattern in _FIELD_LABELS.items()
            if pattern.search(normalized)
        ),
        None,
    )
    tokens = re.findall(r"[A-Z]+(?:/[A-Z]+)?", normalized)
    table_tokens = [token for token in tokens if token in _TABLE_WORDS]
    if not field_match and len(table_tokens) < 2:
        return None

    positions = []
    for word in words:
        token = re.sub(r"[^A-Z/]", "", word["text"].upper())
        if token not in _TABLE_WORDS:
            continue
        bucket = min(9, max(0, int(10 * word["x0"] / max(page_width, 1))))
        positions.append(f"{token}@{bucket}")

    if field_match:
        label, match = field_match
        label_tokens = set(re.findall(r"[A-Z]+", match.group(0).upper()))
        label_word = next(
            (
                word
                for word in words
                if re.sub(r"[^A-Z]", "", word["text"].upper()) in label_tokens
            ),
            None,
        )
        bucket = (
            min(9, max(0, int(10 * label_word["x0"] / max(page_width, 1))))
            if label_word
            else 0
        )
        return f"FIELD:{label}@{bucket}"

    if len(positions) < 2:
        return None
    return "TABLE:" + " ".join(positions)


def extract_layout_signature(data: bytes) -> str:
    """Build a PII-minimized signature from header labels and table headings."""
    signature_lines: list[str] = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for page in pdf.pages[:2]:
            words = page.extract_words(x_tolerance=2, y_tolerance=3)
            rows: list[list[dict]] = []
            for word in sorted(words, key=lambda item: (item["top"], item["x0"])):
                if rows and abs(word["top"] - rows[-1][0]["top"]) <= 3:
                    rows[-1].append(word)
                else:
                    rows.append([word])
            for row in rows:
                row.sort(key=lambda item: item["x0"])
                text = " ".join(word["text"] for word in row)
                line = _line_signature(text, row, page.width)
                if line:
                    signature_lines.append(line)
    return "\n".join(signature_lines)


def extract_source_lines(data: bytes) -> list[dict]:
    """Return selectable text lines and normalized page coordinates for review."""
    lines: list[dict] = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            words = page.extract_words(x_tolerance=2, y_tolerance=3)
            rows: list[list[dict]] = []
            for word in sorted(words, key=lambda item: (item["top"], item["x0"])):
                if rows and abs(word["top"] - rows[-1][0]["top"]) <= 3:
                    rows[-1].append(word)
                else:
                    rows.append([word])
            for row in rows:
                row.sort(key=lambda item: item["x0"])
                text = " ".join(word["text"] for word in row).strip()
                if text:
                    lines.append({
                        "page": page_number,
                        "top": float(row[0]["top"]),
                        "page_height": float(page.height),
                        "page_width": float(page.width),
                        "text": text,
                    })
    return lines


def extract_field_value(field: str, line: str) -> str:
    """Extract a reviewed metadata field from its selected source line."""
    value = " ".join(line.split()).strip(" :-")
    patterns = {
        "bank_name": r"^\s*(?:BANK\s+NAME|BANK)\s*[:\-]\s*",
        "branch": r"^\s*(?:BRANCH(?:\s+NAME)?)\s*[:\-]\s*",
        "account_no": r"^\s*(?:ACCOUNT\s*(?:NO\.?|NUMBER)|A/C\s*(?:NO\.?|NUMBER))\s*[:\-]?\s*",
        "account_holder": (
            r"^\s*(?:ACCOUNT\s+HOLDER|ACCOUNT\s+TITLE|A/C\s+TITLE|CUSTOMER\s+NAME|"
            r"ACCOUNT\s+NAME)\s*(?:[:\-]\s*|\s+)"
        ),
        "period": r"^\s*(?:STATEMENT\s+PERIOD|PERIOD(?:\s+FROM)?|STATEMENT\s+FROM)\s*[:\-]\s*",
    }
    value = re.sub(patterns.get(field, r"(?!)"), "", value, flags=re.IGNORECASE)
    if field == "account_no":
        match = re.search(r"\d[\d\s-]{5,}\d", value)
        return re.sub(r"\D", "", match.group(0)) if match else value
    right_labels = (
        r"\s+(?:ACCOUNT\s*(?:NO\.?|NUMBER)|A/C\s*(?:NO\.?|NUMBER)|"
        r"PRINT(?:ING)?\s+DATE|CUSTOMER\s+ID|BRANCH\s+CODE)\b.*$"
    )
    value = re.sub(right_labels, "", value, flags=re.IGNORECASE)
    return value.strip(" :-,")


def suggest_metadata_fields(
    data: bytes, layout_template: dict
) -> dict[str, str]:
    """Use reviewed relative line positions to suggest corrected metadata."""
    anchors = layout_template.get("field_anchors", {})
    if not anchors:
        return {}
    source_lines = extract_source_lines(data)
    suggestions: dict[str, str] = {}
    for field, anchor in anchors.items():
        candidates = [
            line for line in source_lines if line["page"] == anchor.get("page", 1)
        ]
        if not candidates:
            continue
        target_top = float(anchor["top_ratio"])
        line = min(
            candidates,
            key=lambda item: abs(item["top"] / max(item["page_height"], 1) - target_top),
        )
        distance = abs(line["top"] / max(line["page_height"], 1) - target_top)
        if distance <= 0.06:
            suggestion = extract_field_value(field, line["text"])
            if suggestion:
                suggestions[field] = suggestion
    return suggestions


def _learn_layout_template(
    data: bytes, review_payload: dict
) -> dict:
    """Learn header/table geometry and reviewed source-line field anchors."""
    import bank_parser

    template: dict = {
        "field_anchors": {},
        "transaction_geometry": {},
        "excluded_rows": [],
    }
    field_sources = review_payload.get("field_sources", {})
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for field, source in field_sources.items():
            if not source:
                continue
            page_number = int(source["page"])
            if not 1 <= page_number <= len(pdf.pages):
                continue
            page = pdf.pages[page_number - 1]
            template["field_anchors"][field] = {
                "page": page_number,
                "top_ratio": float(source["top"]) / max(page.height, 1),
            }

        page = pdf.pages[0]
        words = page.extract_words(x_tolerance=2, y_tolerance=3)
        rows = bank_parser._group_rows(words)
        found = bank_parser._find_columns(rows)
        if found:
            columns, header_top = found
            template["transaction_geometry"]["header_top_ratio"] = (
                header_top / max(page.height, 1)
            )
            template["transaction_geometry"]["columns"] = {
                key: {
                    "x0_ratio": span[0] / max(page.width, 1),
                    "x1_ratio": span[1] / max(page.width, 1),
                }
                for key, span in columns.items()
            }
            header_row = min(rows, key=lambda row: abs(row.top - header_top))
            for word in header_row.words:
                if re.fullmatch(r"DATE", word["text"], flags=re.IGNORECASE):
                    template["transaction_geometry"]["date_x_ratio"] = (
                        (word["x0"] + word["x1"]) / (2 * max(page.width, 1))
                    )
                    break

        corrected_columns: dict[str, list[float]] = {
            "debit": [],
            "credit": [],
            "balance": [],
        }
        header_columns = template["transaction_geometry"].get("columns", {})
        corrected_dates: list[float] = []
        annotations = review_payload.get("transaction_sources", [])
        for annotation in annotations:
            page_number = int(annotation["page"])
            if not 1 <= page_number <= len(pdf.pages):
                continue
            page = pdf.pages[page_number - 1]
            page_words = page.extract_words(x_tolerance=2, y_tolerance=3)
            nearby = [
                word
                for word in page_words
                if abs(word["top"] - float(annotation["top"])) <= 3
            ]
            row_text = " ".join(
                word["text"] for word in sorted(nearby, key=lambda word: word["x0"])
            )
            expected_date = re.sub(r"\W", "", str(annotation.get("date", "")))
            if expected_date:
                date_word = next(
                    (
                        word
                        for word in nearby
                        if (
                            match := (
                                bank_parser.DATE_RE.match(word["text"])
                                or bank_parser.SERIAL_DATE_RE.match(word["text"])
                            )
                        )
                        and re.sub(r"\W", "", match.group("date"))
                        == expected_date
                    ),
                    None,
                )
                if date_word:
                    corrected_dates.append(
                        (date_word["x0"] + date_word["x1"]) / (2 * max(page.width, 1))
                    )
            for field in ("debit", "credit", "balance"):
                expected = annotation.get(field)
                if expected is None:
                    continue
                expected_value = bank_parser._money(str(expected))
                if expected_value is None:
                    continue
                candidates = [
                    item
                    for item in nearby
                    if bank_parser._money(item["text"]) == expected_value
                ]
                header_position = header_columns.get(field, {}).get("x1_ratio")
                word = (
                    min(
                        candidates,
                        key=lambda item: abs(
                            item["x1"] / max(page.width, 1) - header_position
                        ),
                    )
                    if candidates and header_position is not None
                    else candidates[0] if candidates else None
                )
                if word:
                    corrected_columns[field].append(word["x1"] / max(page.width, 1))

        for annotation in review_payload.get("excluded_transaction_sources", []):
            page_number = int(annotation["page"])
            if not 1 <= page_number <= len(pdf.pages):
                continue
            page = pdf.pages[page_number - 1]
            row_words = [
                word
                for word in page.extract_words(x_tolerance=2, y_tolerance=3)
                if abs(word["top"] - float(annotation["top"])) <= 3
            ]
            row_text = " ".join(
                word["text"] for word in sorted(row_words, key=lambda word: word["x0"])
            )
            fingerprint = _row_fingerprint(row_text)
            if fingerprint:
                template["excluded_rows"].append({
                    "page": page_number,
                    "top_ratio": float(annotation["top"]) / max(page.height, 1),
                    "fingerprint": fingerprint,
                })

        geometry = template["transaction_geometry"]
        if corrected_dates:
            geometry["date_x_ratio"] = statistics.median(corrected_dates)
        for field, positions in corrected_columns.items():
            if positions:
                x1_ratio = statistics.median(positions)
                geometry.setdefault("columns", {})[field] = {
                    "x0_ratio": max(0.0, x1_ratio - 0.08),
                    "x1_ratio": x1_ratio,
                }
    return template


def _row_fingerprint(text: str) -> str:
    tokens = re.findall(r"[A-Z]{2,}", text.upper())
    if not tokens:
        return ""
    return hashlib.sha256(" ".join(tokens).encode("utf-8")).hexdigest()


def training_example_count(
    database_path: str | os.PathLike[str] | None = None,
) -> int:
    with closing(_connect(database_path)) as connection:
        with connection:
            return int(
                connection.execute(
                    "SELECT COUNT(*) FROM reviewed_layouts WHERE model_version = ?",
                    (MODEL_VERSION,),
                ).fetchone()[0]
            )


def record_review_metrics(
    *,
    field_correct: int,
    field_total: int,
    row_true_positive: int,
    row_false_positive: int,
    row_false_negative: int,
    transaction_value_correct: int,
    transaction_value_total: int,
    database_path: str | os.PathLike[str] | None = None,
) -> None:
    """Persist aggregate pre-review prediction accuracy counts."""
    with closing(_connect(database_path)) as connection:
        with connection:
            connection.execute(
                """
                INSERT INTO review_metrics (
                    reviewed_at, field_correct, field_total, row_true_positive,
                    row_false_positive, row_false_negative,
                    transaction_value_correct, transaction_value_total
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    datetime.now(timezone.utc).isoformat(),
                    field_correct,
                    field_total,
                    row_true_positive,
                    row_false_positive,
                    row_false_negative,
                    transaction_value_correct,
                    transaction_value_total,
                ),
            )


def review_metrics(
    database_path: str | os.PathLike[str] | None = None,
) -> dict[str, int | float]:
    """Return exact field/value accuracy and transaction-row precision/recall."""
    with closing(_connect(database_path)) as connection:
        with connection:
            values = connection.execute(
                """
                SELECT
                    COALESCE(SUM(field_correct), 0),
                    COALESCE(SUM(field_total), 0),
                    COALESCE(SUM(row_true_positive), 0),
                    COALESCE(SUM(row_false_positive), 0),
                    COALESCE(SUM(row_false_negative), 0),
                    COALESCE(SUM(transaction_value_correct), 0),
                    COALESCE(SUM(transaction_value_total), 0)
                FROM review_metrics
                """
            ).fetchone()
    field_correct, field_total, row_tp, row_fp, row_fn, value_correct, value_total = values
    return {
        "field_correct": field_correct,
        "field_total": field_total,
        "field_accuracy": field_correct / field_total if field_total else 0.0,
        "row_true_positive": row_tp,
        "row_false_positive": row_fp,
        "row_false_negative": row_fn,
        "row_precision": row_tp / (row_tp + row_fp) if row_tp + row_fp else 0.0,
        "row_recall": row_tp / (row_tp + row_fn) if row_tp + row_fn else 0.0,
        "transaction_value_correct": value_correct,
        "transaction_value_total": value_total,
        "transaction_value_accuracy": (
            value_correct / value_total if value_total else 0.0
        ),
    }


def _character_ngrams(text: str) -> Counter[str]:
    normalized = " ".join(text.casefold().split())
    return Counter(
        normalized[index : index + size]
        for size in range(2, 6)
        for index in range(max(0, len(normalized) - size + 1))
    )


def _tfidf_vectors(
    documents: list[str],
) -> list[dict[str, float]]:
    counts = [_character_ngrams(document) for document in documents]
    document_frequency: Counter[str] = Counter()
    for document_counts in counts:
        document_frequency.update(document_counts.keys())

    document_total = len(counts)
    vectors: list[dict[str, float]] = []
    for document_counts in counts:
        vector = {
            ngram: (1 + math.log(count))
            * (math.log((1 + document_total) / (1 + document_frequency[ngram])) + 1)
            for ngram, count in document_counts.items()
        }
        norm = math.sqrt(sum(weight * weight for weight in vector.values()))
        if norm:
            vector = {ngram: weight / norm for ngram, weight in vector.items()}
        vectors.append(vector)
    return vectors


def save_reviewed_layout(
    data: bytes,
    layout_label: str,
    review_payload: dict,
    database_path: str | os.PathLike[str] | None = None,
) -> bool:
    """Persist a reviewed layout and its labeled fields; return True if new."""
    label = layout_label.strip()
    if not label:
        raise ValueError("A non-empty reviewed layout label is required.")

    signature = extract_layout_signature(data)
    if not signature:
        raise ValueError(
            "No recognizable header or transaction-table labels were found."
        )

    fingerprint = hashlib.sha256(
        MODEL_VERSION.encode("ascii") + b"\0" + data
    ).hexdigest()
    reviewed_at = datetime.now(timezone.utc).isoformat()
    payload_json = json.dumps(review_payload, ensure_ascii=False, allow_nan=False)
    template_json = json.dumps(
        _learn_layout_template(data, review_payload),
        ensure_ascii=False,
        allow_nan=False,
    )
    is_new_example = False
    with closing(_connect(database_path)) as connection:
        with connection:
            cursor = connection.execute(
                """
                INSERT OR IGNORE INTO reviewed_layouts
                    (fingerprint, signature, layout_label, review_payload,
                     layout_template, reviewed_at, model_version)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    fingerprint,
                    signature,
                    label,
                    payload_json,
                    template_json,
                    reviewed_at,
                    MODEL_VERSION,
                ),
            )
            if cursor.rowcount == 0:
                connection.execute(
                    """
                    UPDATE reviewed_layouts
                    SET signature = ?, layout_label = ?, review_payload = ?,
                        layout_template = ?, reviewed_at = ?, model_version = ?
                    WHERE fingerprint = ?
                    """,
                    (
                        signature,
                        label,
                        payload_json,
                        template_json,
                        reviewed_at,
                        MODEL_VERSION,
                        fingerprint,
                    ),
                )
            else:
                is_new_example = True
    return is_new_example


def find_similar_layout(
    data: bytes,
    database_path: str | os.PathLike[str] | None = None,
) -> LayoutMatch | None:
    """Return the closest reviewed layout using local TF-IDF cosine similarity."""
    signature = extract_layout_signature(data)
    if not signature:
        return None

    with closing(_connect(database_path)) as connection:
        with connection:
            examples = connection.execute(
                """
                SELECT signature, layout_label
                       , layout_template
                FROM reviewed_layouts
                WHERE model_version = ?
                ORDER BY reviewed_at DESC
                """,
                (MODEL_VERSION,),
            ).fetchall()
    if not examples:
        return None

    vectors = _tfidf_vectors([example[0] for example in examples] + [signature])
    query = vectors[-1]
    similarities = [
        sum(weight * vectors[index].get(ngram, 0.0) for ngram, weight in query.items())
        for index in range(len(examples))
    ]
    best_index = max(range(len(similarities)), key=similarities.__getitem__)
    return LayoutMatch(
        layout_label=examples[best_index][1],
        similarity=similarities[best_index],
        example_count=len(examples),
        layout_template=json.loads(examples[best_index][2]),
    )
