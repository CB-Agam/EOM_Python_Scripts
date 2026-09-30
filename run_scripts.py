from __future__ import annotations

import argparse
import numbers
import re
import sys
from pathlib import Path
from typing import Any

import pandas as pd

from excel_to_snowflake import (
    DATE_COLUMN_PATTERN,
    ExcelToSnowflakeLoader,
    MISSING_DATE_MARKERS,
    dataframe_snowflake_type,
    prepare_temporal_columns,
    validate_columns,
)
from fourd_to_snowflake import SnowflakeTargetConfig


PROJECT_DIR = Path(__file__).resolve().parent
DEFAULT_ROOTS = (
    PROJECT_DIR / "2025 EOM Reports-20260923T222736Z-1-002" / "2025 EOM Reports",
    PROJECT_DIR / "2026 EOM Reports-20260923T224007Z-1-002" / "2026 EOM Reports",
)
ERROR_DESCRIPTION_COLUMN = "ERROR_DESCRIPTION"
TEXT_COLUMN_PATTERN = re.compile(
    r"phone|telephone|mobile|contact|name|address|description|email|\btype\b",
    re.IGNORECASE,
)


def table_name_from_file(path: Path) -> str:
    """Convert a workbook stem to a Snowflake-safe underscore table name."""
    name = re.sub(r"\W+", "_", path.stem, flags=re.ASCII).strip("_")
    name = re.sub(r"_+", "_", name)
    if not name:
        raise ValueError(f"Cannot derive a table name from {path.name!r}")
    if len(name) > 255:
        raise ValueError(f"Derived table name exceeds 255 characters: {name}")
    return name


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Load each Excel workbook in the EOM folders to its own Snowflake table."
    )
    input_group = parser.add_mutually_exclusive_group()
    input_group.add_argument(
        "--root",
        type=Path,
        action="append",
        help="folder to scan recursively; repeat to add folders (defaults to both year folders)",
    )
    input_group.add_argument(
        "--file",
        type=Path,
        help="process only this Excel workbook instead of scanning the year folders",
    )
    parser.add_argument(
        "--errors-dir",
        type=Path,
        default=PROJECT_DIR / "load_errors",
        help="folder for rejected-row and file-error workbooks",
    )
    parser.add_argument(
        "--if-exists",
        choices=("fail", "replace"),
        default="fail",
        help="fail safely if a table exists, or replace it (default: fail)",
    )
    parser.add_argument(
        "--chunk-size",
        type=int,
        default=10_000,
        help="rows per Snowflake upload chunk",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="validate workbooks and create reject reports without connecting to Snowflake",
    )
    return parser


def _date_errors_by_row(dataframe: pd.DataFrame) -> dict[int, list[str]]:
    errors: dict[int, list[str]] = {}
    for column in dataframe.columns:
        series = dataframe[column]
        if not pd.api.types.is_object_dtype(series.dtype):
            continue
        if not DATE_COLUMN_PATTERN.search(str(column)):
            continue

        strings = series.map(lambda value: value.strip() if isinstance(value, str) else value)
        missing_marker = strings.map(
            lambda value: isinstance(value, str) and value in MISSING_DATE_MARKERS
        )
        cleaned = strings.mask(missing_marker)
        parsed = pd.to_datetime(cleaned, errors="coerce", format="mixed")
        invalid = cleaned.notna() & parsed.isna()
        for row_index in invalid[invalid].index:
            errors.setdefault(int(row_index), []).append(
                f"Date-like column {column!r} has unparseable value "
                f"{cleaned.loc[row_index]!r}"
            )
    return errors


def _numeric_conversions(
    dataframe: pd.DataFrame,
    errors_by_row: dict[int, list[str]],
) -> dict[str, pd.Series]:
    conversions = {}
    for column in dataframe.columns:
        if TEXT_COLUMN_PATTERN.search(str(column)):
            continue
        series = dataframe[column]
        numeric_cells = series.map(
            lambda value: isinstance(value, numbers.Number)
            and not isinstance(value, bool)
        )
        if not numeric_cells.any():
            continue

        parsed = pd.to_numeric(series, errors="coerce")
        invalid = series.notna() & parsed.isna()
        non_null_count = int(series.notna().sum())
        invalid_ratio = invalid.sum() / max(non_null_count, 1)
        if invalid_ratio > 0.25:
            continue

        conversions[column] = parsed
        for row_index in invalid[invalid].index:
            errors_by_row.setdefault(int(row_index), []).append(
                f"Numeric column {column!r} contains non-numeric value "
                f"{series.loc[row_index]!r}"
            )
    return conversions


def read_workbook_rows(path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Read the first worksheet and separate locally invalid rows."""
    source_dataframe = validate_columns(pd.read_excel(path, sheet_name=0, header=0))
    dataframe = source_dataframe.copy()
    errors_by_row = _date_errors_by_row(dataframe)
    for column, converted in _numeric_conversions(dataframe, errors_by_row).items():
        dataframe[column] = converted

    rejected_rows = []
    for row_index, messages in errors_by_row.items():
        rejected = source_dataframe.iloc[row_index].to_dict()
        rejected[ERROR_DESCRIPTION_COLUMN] = "; ".join(messages)
        rejected_rows.append(rejected)

    valid_indices = [index for index in range(len(dataframe)) if index not in errors_by_row]
    valid_data = dataframe.iloc[valid_indices].copy()
    prepared = prepare_temporal_columns(valid_data) if valid_indices else valid_data
    for column in prepared.columns:
        if dataframe_snowflake_type(prepared[column]) == "TEXT":
            prepared[column] = prepared[column].astype("string")
    return prepared.reset_index(drop=True), pd.DataFrame(rejected_rows)


def error_report_path(errors_dir: Path, source: Path, roots: list[Path]) -> Path:
    root = next((candidate for candidate in roots if source.is_relative_to(candidate)), None)
    if root is None:
        relative = source.stem
    else:
        relative = str(source.relative_to(root).with_suffix(""))
    safe_name = re.sub(r"\W+", "_", relative, flags=re.ASCII).strip("_")
    return errors_dir / f"{safe_name}_errors.xlsx"


def write_error_report(
    output_path: Path,
    rejected_rows: pd.DataFrame,
    file_error: str | None = None,
) -> None:
    if rejected_rows.empty and file_error:
        rejected_rows = pd.DataFrame([{ERROR_DESCRIPTION_COLUMN: file_error}])
    elif not rejected_rows.empty:
        rejected_rows = rejected_rows.copy()
        if ERROR_DESCRIPTION_COLUMN not in rejected_rows.columns:
            rejected_rows[ERROR_DESCRIPTION_COLUMN] = "Row validation failed"
        if file_error:
            summary = dict.fromkeys(rejected_rows.columns)
            summary[ERROR_DESCRIPTION_COLUMN] = file_error
            rejected_rows = pd.concat(
                [rejected_rows, pd.DataFrame([summary])], ignore_index=True
            )
        error_column = rejected_rows.pop(ERROR_DESCRIPTION_COLUMN)
        rejected_rows[ERROR_DESCRIPTION_COLUMN] = error_column

    output_path.parent.mkdir(parents=True, exist_ok=True)
    rejected_rows.to_excel(output_path, index=False)


def process_workbook(
    path: Path,
    roots: list[Path],
    errors_dir: Path,
    loader: ExcelToSnowflakeLoader | None,
    if_exists: str,
) -> tuple[int, int, bool]:
    report_path = error_report_path(errors_dir, path, roots)
    rejected_rows = pd.DataFrame()
    try:
        dataframe, rejected_rows = read_workbook_rows(path)
    except Exception as exc:
        write_error_report(report_path, rejected_rows, str(exc))
        print(f"FAILED {path.name}: {exc}", file=sys.stderr)
        return 0, 1, True

    table_name = table_name_from_file(path)
    if not rejected_rows.empty:
        write_error_report(report_path, rejected_rows)
        print(f"REJECTED {len(rejected_rows):,} row(s): {report_path}")

    if loader is None:
        print(
            f"VALID {path.name}: table={table_name}, "
            f"loadable_rows={len(dataframe):,}, rejected_rows={len(rejected_rows):,}"
        )
        return len(dataframe), len(rejected_rows), False

    try:
        rows_loaded = loader.load(dataframe, table_name, mode=if_exists)
        print(f"LOADED {path.name}: {rows_loaded:,} row(s) -> {table_name}")
        return rows_loaded, len(rejected_rows), False
    except Exception as exc:
        write_error_report(report_path, pd.DataFrame(), str(exc))
        print(f"FAILED {path.name}: {exc}; details: {report_path}", file=sys.stderr)
        return 0, len(rejected_rows) + 1, True


def select_workbooks(args: argparse.Namespace) -> tuple[list[Path], list[Path]] | None:
    if args.file:
        workbook = args.file.resolve()
        if not workbook.is_file():
            print(f"Excel file not found: {workbook}", file=sys.stderr)
            return None
        if workbook.suffix.casefold() != ".xlsx":
            print("Input must be an .xlsx file", file=sys.stderr)
            return None
        return [workbook.parent], [workbook]

    roots = [path.resolve() for path in (args.root or list(DEFAULT_ROOTS))]
    missing_roots = [root for root in roots if not root.is_dir()]
    if missing_roots:
        for root in missing_roots:
            print(f"Input folder not found: {root}", file=sys.stderr)
        return None

    workbooks = sorted(
        path
        for root in roots
        for path in root.rglob("*.xlsx")
        if not path.name.startswith("~$")
    )
    return roots, workbooks


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.chunk_size <= 0:
        print("--chunk-size must be greater than zero", file=sys.stderr)
        return 2

    selection = select_workbooks(args)
    if selection is None:
        return 2
    roots, workbooks = selection

    if not workbooks:
        print("No .xlsx workbooks found.", file=sys.stderr)
        return 1

    loader = None
    connection = None
    if not args.dry_run:
        config = SnowflakeTargetConfig.from_env()
        connection = config.connect()
        loader = ExcelToSnowflakeLoader(connection, config, chunk_size=args.chunk_size)

    total_loaded = 0
    total_rejected = 0
    failed_files = 0
    try:
        for workbook in workbooks:
            loaded, rejected, failed = process_workbook(
                workbook,
                roots,
                args.errors_dir,
                loader,
                args.if_exists,
            )
            total_loaded += loaded
            total_rejected += rejected
            failed_files += int(failed)
    finally:
        if connection is not None:
            connection.close()

    print(
        f"Finished: {len(workbooks)} workbook(s), {total_loaded:,} row(s) loaded/valid, "
        f"{total_rejected:,} rejected/error item(s), {failed_files} failed workbook(s)."
    )
    return 1 if failed_files else 0


if __name__ == "__main__":
    raise SystemExit(main())
