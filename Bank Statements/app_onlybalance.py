"""Bank statement PDF → Excel converter.

Supports the different layouts used by DBBL, IBBL, AB Bank and similar
statements. Amounts are read from the Debit / Credit / Balance columns of the
table, never from account numbers or reference numbers embedded in the
description.
"""

from __future__ import annotations

import hashlib
import importlib
import io
import json
import os
import sqlite3
import sys
from decimal import Decimal

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# When this file is run from the "Bank Statements" folder, bank_parser.py lives
# in the parent folder. Make sure it is importable either way.
_HERE = os.path.dirname(os.path.abspath(__file__))
for _candidate in (_HERE, os.path.dirname(_HERE)):
    if _candidate and _candidate not in sys.path:
        sys.path.insert(0, _candidate)

import bank_parser
import bank_layout_model


def _load_parser_module():
    """Return ``bank_parser``, reloading it when the file on disk changed.

    Streamlit keeps imported local modules in ``sys.modules`` between reruns,
    so after ``bank_parser.py`` is edited the running app can keep executing
    the old code (and even fail with an ImportError for a newly added name).
    Reloading on a changed modification time keeps the two in sync without
    requiring a manual server restart.
    """
    module = bank_parser
    try:
        stamp = os.path.getmtime(module.__file__)
    except (OSError, AttributeError, TypeError):
        stamp = 0.0

    previous = getattr(module, "_LOADED_STAMP", None)
    # ``previous is None`` means this is the first rerun after an upgrade: if
    # the already-imported module predates the current API, reload it too.
    if (previous is None and not hasattr(module, "PARSER_VERSION")) or (
        previous is not None and stamp and stamp != previous
    ):
        module = importlib.reload(module)

    module._LOADED_STAMP = stamp
    return module


_parser = _load_parser_module()
parse_statement = _parser.parse_statement
ParsedStatement = _parser.ParsedStatement
PARSER_VERSION = getattr(_parser, "PARSER_VERSION", "unknown")
Transaction = _parser.Transaction

st.set_page_config(
    page_title="Bank Statement Parser",
    page_icon=":material/account_balance:",
    layout="wide",
    initial_sidebar_state="expanded",
)

MONEY_FORMAT = "#,##0.00;[Red]-#,##0.00"
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TOTAL_FONT = Font(bold=True)


@st.cache_data(show_spinner=False)
def parse_pdf_bytes(
    data: bytes,
    filename: str,
    parser_version: str = PARSER_VERSION,
    learned_layout_json: str = "",
) -> ParsedStatement:
    """Parse one PDF and cache the result for the current session.

    ``parser_version`` is part of the cache key only; it guarantees that an
    upgraded parser never reuses a result produced by an older version.
    """
    del parser_version
    learned_layout = json.loads(learned_layout_json) if learned_layout_json else None
    return parse_statement(data, filename, learned_layout=learned_layout)


def _decimal(value: Decimal | None) -> float | None:
    return float(value) if value is not None else None


def _same_review_value(left, right) -> bool:
    if pd.isna(left) and pd.isna(right):
        return True
    if pd.isna(left) or pd.isna(right):
        return False
    if isinstance(left, (int, float, Decimal)) and isinstance(
        right, (int, float, Decimal)
    ):
        return Decimal(str(left)) == Decimal(str(right))
    return " ".join(str(left).split()).casefold() == " ".join(
        str(right).split()
    ).casefold()


def _to_frames(parsed_statements: list[ParsedStatement]):
    rows: list[dict] = []
    summary_rows: list[dict] = []

    for parsed in parsed_statements:
        info = parsed.info
        transactions = parsed.transactions

        debit_total = sum((t.debit or Decimal("0")) for t in transactions)
        credit_total = sum((t.credit or Decimal("0")) for t in transactions)
        balances = [t.balance for t in transactions if t.balance is not None]
        closing = balances[-1] if balances else None

        opening = next(
            (
                t.balance
                for t in transactions
                if t.is_opening and t.balance is not None
            ),
            None,
        )
        if opening is None and transactions and transactions[0].balance is not None:
            first = transactions[0]
            opening = first.balance - (
                (first.credit or Decimal("0")) - (first.debit or Decimal("0"))
            )

        net_change = (
            closing - (opening if opening is not None else Decimal("0"))
            if closing is not None
            else None
        )
        difference = (
            net_change - (credit_total - debit_total)
            if net_change is not None
            else None
        )

        for transaction in transactions:
            rows.append({
                "Bank Name": info.bank_name or "—",
                "Branch": info.branch or "—",
                "Account No": info.account_no or "—",
                "Account Holder": info.account_holder or "—",
                "Source File": parsed.source_name,
                "Transaction Date": transaction.date,
                "Description / Particulars": transaction.description,
                "Debit": _decimal(transaction.debit),
                "Credit": _decimal(transaction.credit),
                "Balance": _decimal(transaction.balance),
            })

        period = info.period
        if not period and transactions:
            period = f"{transactions[0].date} – {transactions[-1].date}"

        summary_rows.append({
            "Source File": parsed.source_name,
            "Bank Name": info.bank_name or "—",
            "Account No": info.account_no or "—",
            "Account Holder": info.account_holder or "—",
            "Period": period or "—",
            "Transactions": len(transactions),
            "Opening Balance": _decimal(opening),
            "Closing Balance": _decimal(closing),
            "Total Debit": float(debit_total),
            "Total Credit": float(credit_total),
            "Net Change": _decimal(net_change),
            "Difference": _decimal(difference),
        })

    transaction_columns = [
        "Bank Name",
        "Branch",
        "Account No",
        "Account Holder",
        "Source File",
        "Transaction Date",
        "Description / Particulars",
        "Debit",
        "Credit",
        "Balance",
    ]
    return (
        pd.DataFrame(rows, columns=transaction_columns),
        pd.DataFrame(summary_rows),
    )


def _add_total_row(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    total_row = {column: "" for column in df.columns}
    total_row["Description / Particulars"] = "Total Sum"
    total_row["Debit"] = round(float(df["Debit"].fillna(0).sum()), 2)
    total_row["Credit"] = round(float(df["Credit"].fillna(0).sum()), 2)
    return pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)


def _write_sheet(writer, df: pd.DataFrame, sheet_name: str, money_columns: list[str]):
    df.to_excel(writer, index=False, sheet_name=sheet_name)
    worksheet = writer.sheets[sheet_name]

    for cell in worksheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")

    headers = {cell.value: cell.column for cell in worksheet[1]}
    for column_name in money_columns:
        column_index = headers.get(column_name)
        if not column_index:
            continue
        for (cell,) in worksheet.iter_rows(
            min_col=column_index,
            max_col=column_index,
            min_row=2,
            max_row=worksheet.max_row,
        ):
            cell.number_format = MONEY_FORMAT

    for column_index, column_name in enumerate(df.columns, start=1):
        values = [str(column_name)] + [
            "" if pd.isna(value) else str(value) for value in df[column_name]
        ]
        width = max((len(value) for value in values), default=10) + 2
        worksheet.column_dimensions[get_column_letter(column_index)].width = min(
            width, 60
        )

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions


def build_excel(df: pd.DataFrame, summary_df: pd.DataFrame) -> bytes:
    excel_df = _add_total_row(df)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        _write_sheet(
            writer,
            excel_df,
            "Transactions",
            ["Debit", "Credit", "Balance"],
        )
        _write_sheet(
            writer,
            summary_df,
            "Summary",
            [
                "Opening Balance",
                "Closing Balance",
                "Total Debit",
                "Total Credit",
                "Net Change",
                "Difference",
            ],
        )

        transactions_sheet = writer.sheets["Transactions"]
        for cell in transactions_sheet[transactions_sheet.max_row]:
            cell.font = TOTAL_FONT

    return output.getvalue()


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

def _render_header() -> None:
    """Render the branded page header and support/version badges."""
    st.title(
        "Bank Statement Header & Transaction Extractor",
        icon=":material/account_balance:",
        anchor=False,
    )
    st.caption(
        "DBBL, IBBL, AB Bank and similar statement PDFs are parsed the same way. "
        "Amounts are taken only from the Debit / Credit / Balance columns."
    )
    st.markdown(
        ":blue-badge[:material/verified: DBBL] "
        ":blue-badge[:material/verified: IBBL] "
        ":blue-badge[:material/verified: AB Bank] "
        f":gray-badge[:material/memory: Parser v{PARSER_VERSION}]"
    )


def _render_empty_state() -> None:
    """Render onboarding guidance before any PDF has been uploaded."""
    with st.container(border=True):
        st.subheader(
            "Get Started",
            icon=":material/rocket_launch:",
            anchor=False,
        )
        st.write(
            "Upload one or more bank statement PDFs from the left sidebar. "
            "The system reads header information and transactions separately "
            "and produces a ready-to-use Excel file."
        )

    st.space("small")

    steps = [
        (
            ":material/upload_file:",
            "1. Upload PDF",
            "You can select multiple statements at once.",
        ),
        (
            ":material/document_scanner:",
            "2. Automatic Parsing",
            "The header, balances and every transaction are read automatically.",
        ),
        (
            ":material/download:",
            "3. Download Excel",
            "Save everything in one file with two sheets: Transactions and Summary.",
        ),
    ]
    for column, (icon, title, description) in zip(
        st.columns(3, border=True, gap="medium"), steps
    ):
        with column:
            st.subheader(title, icon=icon, anchor=False)
            st.caption(description)

    st.space("small")
    st.info(
        "Supported banks: DBBL, IBBL, AB Bank and similar "
        "table-based statements.",
        icon=":material/info:",
    )


_render_header()

with st.sidebar:
    st.subheader(
        "File Upload",
        icon=":material/upload_file:",
        anchor=False,
    )
    st.caption(
        "Select bank statement PDFs. Multiple files can be added at once."
    )
    uploaded_files = st.file_uploader(
        "Bank Statement PDF",
        type=["pdf"],
        accept_multiple_files=True,
        help="Upload multiple PDFs together and they will all be added to the same Excel file.",
        label_visibility="collapsed",
    )

    if uploaded_files:
        st.badge(
            f"{len(uploaded_files)} files selected",
            icon=":material/check_circle:",
            color="green",
        )

    st.divider()

    with st.container(border=True):
        st.markdown("**How It Works**")
        st.markdown(
            "1. Upload the PDF\n"
            "2. The header and transactions are read automatically\n"
            "3. Download the Excel file and verify"
        )

    st.space("small")
    st.caption(
        f"Parser v{PARSER_VERSION} · Column-based Debit / Credit / Balance"
    )
    st.caption(
        "Uploaded files are processed only in this session."
    )

if not uploaded_files:
    _render_empty_state()
    st.stop()

MODEL_MIN_SIMILARITY = 0.55
field_columns = {
    "bank_name": "Bank Name",
    "branch": "Branch",
    "account_no": "Account No",
    "account_holder": "Account Holder",
    "period": "Period",
}

parsed_statements: list[ParsedStatement] = []
layout_matches = []
layout_templates: list[dict] = []
source_lines_by_statement: list[list[dict]] = []
source_options_by_statement: list[list[str]] = []
source_lookup: dict[str, tuple[int, dict]] = {}

with st.status("Processing PDFs…", expanded=True) as parse_status:
    progress = st.progress(0.0, text="Starting to read PDFs…")
    for index, uploaded_file in enumerate(uploaded_files, start=1):
        progress.progress(
            (index - 1) / len(uploaded_files),
            text=f"Learning layout: {uploaded_file.name}",
        )
        data = uploaded_file.getvalue()
        try:
            match = bank_layout_model.find_similar_layout(data)
            source_lines = bank_layout_model.extract_source_lines(data)
        except (OSError, sqlite3.Error, ValueError) as error:
            st.error(f"Could not inspect {uploaded_file.name}: {error}")
            st.stop()

        learned_layout = (
            match.layout_template
            if match is not None and match.similarity >= MODEL_MIN_SIMILARITY
            else {}
        )
        try:
            parsed = parse_pdf_bytes(
                data,
                uploaded_file.name,
                PARSER_VERSION,
                json.dumps(learned_layout, sort_keys=True),
            )
            if learned_layout:
                suggestions = bank_layout_model.suggest_metadata_fields(
                    data, learned_layout
                )
                for field, value in suggestions.items():
                    if field in field_columns and value:
                        setattr(parsed.info, field, value)
        except (OSError, ValueError, json.JSONDecodeError) as error:
            st.error(f"Could not apply the learned layout for {uploaded_file.name}: {error}")
            st.stop()

        options: list[str] = []
        for line_index, line in enumerate(source_lines, start=1):
            option = (
                f"S{index}-L{line_index} · p{line['page']} y={line['top']:.1f} · "
                f"{line['text'][:90]}"
            )
            line["source_option"] = option
            options.append(option)
            source_lookup[option] = (index - 1, line)

        parsed_statements.append(parsed)
        layout_matches.append(match)
        layout_templates.append(learned_layout)
        source_lines_by_statement.append(source_lines)
        source_options_by_statement.append(options)
    progress.progress(1.0, text="Parsing complete")
    parse_status.update(
        label=f"{len(uploaded_files)} PDFs parsed successfully",
        state="complete",
        expanded=False,
    )

warnings = [
    (parsed.source_name, warning)
    for parsed in parsed_statements
    for warning in parsed.warnings
]
if warnings:
    with st.expander(
        f"{len(warnings)} warnings",
        icon=":material/warning:",
        expanded=False,
    ):
        for source_name, warning in warnings:
            st.warning(warning, title=f"Warning — {source_name}")

df, summary_df = _to_frames(parsed_statements)

try:
    example_count = bank_layout_model.training_example_count()
    quality_metrics = bank_layout_model.review_metrics()
except sqlite3.Error as error:
    st.error(f"Could not read local model data: {error}")
    st.stop()

summary_editor_rows = []
for statement_index, parsed in enumerate(parsed_statements):
    row = {"Source File": parsed.source_name}
    for field, column in field_columns.items():
        value = getattr(parsed.info, field)
        row[column] = value
        matching_line = next(
            (
                line["source_option"]
                for line in source_lines_by_statement[statement_index]
                if value and value.casefold() in line["text"].casefold()
            ),
            "",
        )
        row[f"{column} Source Line"] = matching_line
    summary_editor_rows.append(row)
summary_editor_df = pd.DataFrame(summary_editor_rows)

transaction_editor_rows = []
initial_transaction_predictions_by_statement: list[dict[str, dict]] = [
    {} for _ in parsed_statements
]
initial_source_options_by_statement: list[set[str]] = [
    set() for _ in parsed_statements
]
for statement_index, parsed in enumerate(parsed_statements):
    lines = source_lines_by_statement[statement_index]
    for transaction in parsed.transactions:
        nearest = min(
            (
                line
                for line in lines
                if line["page"] == transaction.page
            ),
            key=lambda line: abs(line["top"] - transaction.top),
            default=None,
        )
        transaction_editor_rows.append({
            "Source File": parsed.source_name,
            "Source Line": nearest["source_option"] if nearest else "",
            "Transaction Date": transaction.date,
            "Description / Particulars": transaction.description,
            "Debit": _decimal(transaction.debit),
            "Credit": _decimal(transaction.credit),
            "Balance": _decimal(transaction.balance),
        })
        if nearest:
            initial_source_options_by_statement[statement_index].add(
                nearest["source_option"]
            )
            initial_transaction_predictions_by_statement[statement_index][
                nearest["source_option"]
            ] = {
                "Transaction Date": transaction.date,
                "Description / Particulars": transaction.description,
                "Debit": _decimal(transaction.debit),
                "Credit": _decimal(transaction.credit),
                "Balance": _decimal(transaction.balance),
            }
transaction_editor_df = pd.DataFrame(
    transaction_editor_rows,
    columns=[
        "Source File",
        "Source Line",
        "Transaction Date",
        "Description / Particulars",
        "Debit",
        "Credit",
        "Balance",
    ],
)
all_source_options = [
    option for options in source_options_by_statement for option in options
]
predicted_metadata_by_statement = [
    {
        field: getattr(parsed.info, field)
        for field in field_columns
    }
    for parsed in parsed_statements
]
upload_fingerprint = hashlib.sha256()
for uploaded_file in uploaded_files:
    upload_fingerprint.update(uploaded_file.name.encode("utf-8"))
    upload_fingerprint.update(b"\0")
    upload_fingerprint.update(uploaded_file.getvalue())
review_key = upload_fingerprint.hexdigest()[:16]

with st.container(border=True):
    st.subheader("Local Supervised Layout Model", icon=":material/model_training:", anchor=False)
    st.caption(
        f"{example_count} reviewed training example(s). On a strong layout match, "
        "the learned field positions and transaction columns are applied before "
        "the existing parser runs. Correct the results, select the source PDF "
        "line for each field/transaction, and save the review. Added or removed "
        "transaction rows are included in training. The PDF itself is not saved; "
        "reviewed account and transaction data are stored locally without encryption."
    )
    if quality_metrics["field_total"]:
        st.caption(
            "Quality figures compare parser/model predictions with corrections "
            "you saved; they are not results from an independent test set."
        )
        metric_columns = st.columns(4)
        metric_columns[0].metric(
            "Reviewed field accuracy",
            f"{quality_metrics['field_accuracy']:.0%}",
            help="Exact field matches before correction across saved reviews, not held-out testing.",
        )
        metric_columns[1].metric(
            "Transaction row precision",
            f"{quality_metrics['row_precision']:.0%}",
        )
        metric_columns[2].metric(
            "Transaction row recall",
            f"{quality_metrics['row_recall']:.0%}",
        )
        metric_columns[3].metric(
            "Transaction value accuracy",
            f"{quality_metrics['transaction_value_accuracy']:.0%}",
            help="Exact values for rows matched to their reviewed PDF source line.",
        )
    for uploaded_file, match in zip(uploaded_files, layout_matches):
        if match is not None and match.similarity >= MODEL_MIN_SIMILARITY:
            st.info(
                f"{uploaded_file.name}: closest reviewed layout is "
                f"“{match.layout_label}” "
                f"(layout similarity {match.similarity:.0%}). Learned field "
                "and transaction geometry was applied; verify the result."
            )
        elif match is not None:
            st.warning(
                f"{uploaded_file.name}: closest reviewed layout is "
                f"“{match.layout_label}” ({match.similarity:.0%} similarity), "
                "below the application threshold. The baseline parser was used."
            )
        else:
            st.caption(
                f"{uploaded_file.name}: no matching training layout yet."
            )

    st.markdown("**1. Review statement fields and select their source lines**")
    st.markdown(
        "**2. Review transactions. Add missing rows or remove false rows, then "
        "select the source line for each transaction.**"
    )
    with st.form(f"review_and_train_layout_model_{review_key}"):
        edited_summary = st.data_editor(
            summary_editor_df,
            hide_index=True,
            width="stretch",
            disabled=["Source File"],
            column_config={
                **{
                    column: st.column_config.TextColumn()
                    for column in field_columns.values()
                },
                **{
                    f"{column} Source Line": st.column_config.SelectboxColumn(
                        options=["", *all_source_options],
                        help="Select the PDF line supporting this corrected field.",
                    )
                    for column in field_columns.values()
                },
            },
            key=f"reviewed_statement_fields_{review_key}",
        )
        edited_transactions = st.data_editor(
            transaction_editor_df,
            hide_index=True,
            width="stretch",
            disabled=["Source File"],
            num_rows="dynamic",
            column_order=[
                "Source File",
                "Source Line",
                "Transaction Date",
                "Description / Particulars",
                "Debit",
                "Credit",
                "Balance",
            ],
            column_config={
                "Source Line": st.column_config.SelectboxColumn(
                    options=["", *all_source_options],
                    help="For added rows, select the PDF line containing that transaction.",
                ),
                "Debit": st.column_config.NumberColumn(format="accounting"),
                "Credit": st.column_config.NumberColumn(format="accounting"),
                "Balance": st.column_config.NumberColumn(format="accounting"),
            },
            key=f"reviewed_transactions_{review_key}",
        )
        confirm_local_storage = st.checkbox(
            "I understand reviewed account and transaction data will be stored "
            "locally without encryption."
        )
        save_training_examples = st.form_submit_button(
            "Save reviewed examples and update layout model",
            type="primary",
        )

if save_training_examples and not confirm_local_storage:
    st.warning(
        "Confirm local storage before saving. No training data was written."
    )

if save_training_examples and confirm_local_storage:
    field_sources_by_statement: list[dict] = [
        {} for _ in parsed_statements
    ]
    field_correct = 0
    field_total = 0
    for statement_index, parsed in enumerate(parsed_statements):
        reviewed_info = edited_summary.iloc[statement_index]
        for field, column in field_columns.items():
            value = reviewed_info[column]
            reviewed_value = "" if pd.isna(value) else str(value).strip()
            predicted_value = predicted_metadata_by_statement[statement_index][field]
            if reviewed_value:
                field_total += 1
                field_correct += _same_review_value(
                    predicted_value, reviewed_value
                )
            setattr(parsed.info, field, "" if pd.isna(value) else str(value).strip())
            source_option = reviewed_info[f"{column} Source Line"]
            if pd.isna(source_option) or not source_option:
                continue
            source_statement_index, line = source_lookup[source_option]
            if source_statement_index != statement_index:
                st.error(
                    f"Select a source line from {parsed.source_name} for "
                    f"{column}; no examples were saved."
                )
                st.stop()
            field_sources_by_statement[statement_index][field] = {
                "page": line["page"],
                "top": line["top"],
            }

    reviewed_transactions: list[list[Transaction]] = [
        [] for _ in parsed_statements
    ]
    selected_source_options_by_statement: list[set[str]] = [
        set() for _ in parsed_statements
    ]
    transaction_sources_by_statement: list[list[dict]] = [
        [] for _ in parsed_statements
    ]
    reviewed_transaction_values_by_statement: list[dict[str, dict]] = [
        {} for _ in parsed_statements
    ]
    for row_index, (_, row) in enumerate(edited_transactions.iterrows()):
        source_option = row["Source Line"]
        if pd.isna(source_option) or not source_option:
            st.error(
                f"Select a source PDF line for transaction row {row_index + 1}; "
                "no examples were saved."
            )
            st.stop()
        statement_index, line = source_lookup[source_option]
        selected_source_options_by_statement[statement_index].add(source_option)
        date_value = (
            "" if pd.isna(row["Transaction Date"]) else str(row["Transaction Date"]).strip()
        )
        description = (
            ""
            if pd.isna(row["Description / Particulars"])
            else str(row["Description / Particulars"]).strip()
        )
        amounts = {}
        for field in ("debit", "credit", "balance"):
            value = row[field.title()]
            amounts[field] = (
                None
                if pd.isna(value) or value == ""
                else Decimal(str(value))
            )
        reviewed_transactions[statement_index].append(
            Transaction(
                date=date_value,
                description=description,
                debit=amounts["debit"],
                credit=amounts["credit"],
                balance=amounts["balance"],
                page=line["page"],
                top=line["top"],
                is_opening=bank_parser._is_opening(description),
            )
        )
        reviewed_transaction_values_by_statement[statement_index][source_option] = {
            "Transaction Date": date_value,
            "Description / Particulars": description,
            "Debit": _decimal(amounts["debit"]),
            "Credit": _decimal(amounts["credit"]),
            "Balance": _decimal(amounts["balance"]),
        }
        transaction_sources_by_statement[statement_index].append({
            "page": line["page"],
            "top": line["top"],
            "date": date_value,
            "debit": str(amounts["debit"]) if amounts["debit"] is not None else None,
            "credit": str(amounts["credit"]) if amounts["credit"] is not None else None,
            "balance": (
                str(amounts["balance"]) if amounts["balance"] is not None else None
            ),
        })

    for statement_index, parsed in enumerate(parsed_statements):
        parsed.transactions = reviewed_transactions[statement_index]

    excluded_transaction_sources_by_statement = []
    for statement_index, source_options in enumerate(
        initial_source_options_by_statement
    ):
        excluded_transaction_sources_by_statement.append([
            {
                "page": source_lookup[source_option][1]["page"],
                "top": source_lookup[source_option][1]["top"],
            }
            for source_option in source_options
            if source_option
            not in selected_source_options_by_statement[statement_index]
        ])

    row_true_positive = 0
    row_false_positive = 0
    row_false_negative = 0
    transaction_value_correct = 0
    transaction_value_total = 0
    transaction_value_fields = (
        "Transaction Date",
        "Description / Particulars",
        "Debit",
        "Credit",
        "Balance",
    )
    for statement_index in range(len(parsed_statements)):
        predicted_rows = initial_transaction_predictions_by_statement[
            statement_index
        ]
        reviewed_rows = reviewed_transaction_values_by_statement[statement_index]
        predicted_sources = set(predicted_rows)
        reviewed_sources = set(reviewed_rows)
        row_true_positive += len(predicted_sources & reviewed_sources)
        row_false_positive += len(predicted_sources - reviewed_sources)
        row_false_negative += len(reviewed_sources - predicted_sources)
        for source_option in predicted_sources & reviewed_sources:
            for field in transaction_value_fields:
                transaction_value_total += 1
                transaction_value_correct += _same_review_value(
                    predicted_rows[source_option][field],
                    reviewed_rows[source_option][field],
                )

    df, summary_df = _to_frames(parsed_statements)
    saved_count = 0
    try:
        for statement_index, (uploaded_file, parsed) in enumerate(
            zip(uploaded_files, parsed_statements)
        ):
            review_payload = {
                "metadata": {
                    "bank_name": parsed.info.bank_name,
                    "branch": parsed.info.branch,
                    "account_no": parsed.info.account_no,
                    "account_holder": parsed.info.account_holder,
                    "period": parsed.info.period,
                },
                "field_sources": field_sources_by_statement[statement_index],
                "transaction_sources": transaction_sources_by_statement[
                    statement_index
                ],
                "excluded_transaction_sources": (
                    excluded_transaction_sources_by_statement[statement_index]
                ),
                "transactions": [
                    {
                        "date": transaction.date,
                        "description": transaction.description,
                        "debit": (
                            str(transaction.debit)
                            if transaction.debit is not None
                            else None
                        ),
                        "credit": (
                            str(transaction.credit)
                            if transaction.credit is not None
                            else None
                        ),
                        "balance": (
                            str(transaction.balance)
                            if transaction.balance is not None
                            else None
                        ),
                    }
                    for transaction in parsed.transactions
                ],
            }
            if bank_layout_model.save_reviewed_layout(
                uploaded_file.getvalue(),
                parsed.info.bank_name or "Unlabeled bank",
                review_payload,
            ):
                saved_count += 1
        bank_layout_model.record_review_metrics(
            field_correct=field_correct,
            field_total=field_total,
            row_true_positive=row_true_positive,
            row_false_positive=row_false_positive,
            row_false_negative=row_false_negative,
            transaction_value_correct=transaction_value_correct,
            transaction_value_total=transaction_value_total,
        )
        st.success(
            f"Saved {len(uploaded_files)} reviewed statement(s) locally; "
            f"{saved_count} were new training examples. Their reviewed field "
            "positions and transaction geometry will guide future matching PDFs."
        )
        st.caption(
            "Training database: "
            f"{bank_layout_model.training_database_path()}"
        )
    except (OSError, sqlite3.Error, ValueError) as error:
        st.error(f"Could not save reviewed training examples: {error}")

total_debit = float(df["Debit"].fillna(0).sum())
total_credit = float(df["Credit"].fillna(0).sum())

st.subheader("At a Glance", icon=":material/insights:", anchor=False)
st.caption(
    f"{len(summary_df):,} files · {len(df):,} transactions · "
    f"{df['Bank Name'].nunique():,} banks"
)

metric_columns = st.columns(4, gap="medium")
metric_columns[0].metric(
    "Total Transactions",
    f"{len(df):,}",
    icon=":material/receipt_long:",
    border=True,
)
metric_columns[1].metric(
    "Total Debit",
    f"{total_debit:,.2f}",
    icon=":material/trending_down:",
    border=True,
    help="Sum of the Debit column across all transactions.",
)
metric_columns[2].metric(
    "Total Credit",
    f"{total_credit:,.2f}",
    icon=":material/trending_up:",
    border=True,
    help="Sum of the Credit column across all transactions.",
)
metric_columns[3].metric(
    "Net Change",
    f"{total_credit - total_debit:,.2f}",
    icon=":material/swap_vert:",
    border=True,
    help="Total Credit − Total Debit.",
)

st.space("small")

with st.container(border=True):
    st.subheader("Summary", icon=":material/summarize:", anchor=False)
    st.caption(
        "Opening and closing balances, total debit and credit, and "
        "reconciliation for each PDF. Account Holder values can be reviewed "
        "in the section above."
    )
    st.dataframe(
        summary_df,
        hide_index=True,
        width="stretch",
        alt="Summary and reconciliation for each PDF file",
        column_config={
            "Opening Balance": st.column_config.NumberColumn(format="accounting"),
            "Closing Balance": st.column_config.NumberColumn(format="accounting"),
            "Total Debit": st.column_config.NumberColumn(format="accounting"),
            "Total Credit": st.column_config.NumberColumn(format="accounting"),
            "Net Change": st.column_config.NumberColumn(format="accounting"),
            "Difference": st.column_config.NumberColumn(format="accounting"),
        },
    )

st.space("small")

excel_data = build_excel(df, summary_df)

with st.container(border=True):
    heading_column, action_column = st.columns(
        [3, 1], vertical_alignment="center", gap="medium"
    )
    with heading_column:
        st.subheader(
            "Transaction Details",
            icon=":material/table_rows:",
            anchor=False,
        )
        st.caption(
            f"{len(df):,} transactions · The Excel file will contain two "
            "sheets: Transactions and Summary."
        )
    with action_column:
        st.download_button(
            label="Download Excel",
            data=excel_data,
            file_name="Bank_Statement_Parsed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
            type="primary",
            width="stretch",
            help="Download an .xlsx file with all transactions and the summary.",
        )

    st.dataframe(
        df,
        hide_index=True,
        width="stretch",
        height=520,
        alt="List of parsed bank transactions",
        column_config={
            "Debit": st.column_config.NumberColumn(format="accounting"),
            "Credit": st.column_config.NumberColumn(format="accounting"),
            "Balance": st.column_config.NumberColumn(format="accounting"),
        },
    )

st.divider()
st.caption(
    ":material/verified_user: Uploaded files are processed only in this session "
    "and are not stored anywhere."
)
