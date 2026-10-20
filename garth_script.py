"""Convert a tab-delimited text file to Excel and load it into Snowflake."""

from __future__ import annotations

import argparse
import csv
import re
import sys
from pathlib import Path

import pandas as pd

from excel_to_snowflake import (
    ExcelToSnowflakeLoader,
    prepare_temporal_columns,
    read_excel_sheet,
    validate_columns,
)
from fourd_to_snowflake import SnowflakeTargetConfig


SUPPORTED_EXCEL_SUFFIXES = {".xls", ".xlsm", ".xlsx"}
SOURCE_FILE_COLUMN = "SOURCE_FILE_NAME"


def table_name_from_path(path: Path) -> str:
    """Create a Snowflake-safe table name from a file or folder name."""
    source_name = path.name if path.is_dir() else path.stem
    name = re.sub(r"\W+", "_", source_name, flags=re.ASCII).strip("_")
    name = re.sub(r"_+", "_", name)
    if not name:
        raise ValueError(f"Cannot derive a table name from {path.name!r}")
    if len(name) > 255:
        raise ValueError(f"Derived table name exceeds 255 characters: {name}")
    return name


def resolve_input(input_path: Path) -> tuple[list[Path], str]:
    """Resolve a file or the supported files directly inside a folder."""
    path = input_path.expanduser().resolve()
    if path.is_dir():
        source_files = sorted(
            candidate
            for candidate in path.iterdir()
            if candidate.is_file()
            and (
                candidate.suffix.casefold() == ".txt"
                or candidate.suffix.casefold() in SUPPORTED_EXCEL_SUFFIXES
            )
            and not candidate.name.startswith("~$")
        )
        text_stems = {
            candidate.stem.casefold()
            for candidate in source_files
            if candidate.suffix.casefold() == ".txt"
        }
        source_files = [
            candidate
            for candidate in source_files
            if not (
                candidate.suffix.casefold() in SUPPORTED_EXCEL_SUFFIXES
                and candidate.stem.casefold() in text_stems
            )
        ]
        if not source_files:
            raise ValueError(f"No supported .txt or Excel files found directly inside {path}")
        return source_files, table_name_from_path(path)

    if not path.is_file():
        raise FileNotFoundError(f"Input file or folder not found: {path}")

    suffix = path.suffix.casefold()
    if suffix != ".txt" and suffix not in SUPPORTED_EXCEL_SUFFIXES:
        raise ValueError("Input must be a .txt, .xlsx, .xlsm, or .xls file")
    return [path], table_name_from_path(path)


def add_source_file_column(dataframe: pd.DataFrame, source: Path) -> pd.DataFrame:
    """Add the source filename (without extension) to each input row."""
    if SOURCE_FILE_COLUMN in dataframe.columns:
        raise ValueError(
            f"Input already contains the reserved {SOURCE_FILE_COLUMN!r} column"
        )
    prepared = dataframe.copy()
    prepared[SOURCE_FILE_COLUMN] = source.stem
    return prepared


def resolve_text_encoding(path: Path, encoding: str) -> str:
    """Resolve an explicit encoding or detect common BOM/Windows encodings."""
    if encoding.casefold() != "auto":
        return encoding

    with path.open("rb") as source_file:
        prefix = source_file.read(4)
    if prefix.startswith(b"\x00\x00\xfe\xff") or prefix.startswith(b"\xff\xfe\x00\x00"):
        return "utf-32"
    if prefix.startswith(b"\xef\xbb\xbf"):
        return "utf-8-sig"
    if prefix.startswith(b"\xff\xfe") or prefix.startswith(b"\xfe\xff"):
        return "utf-16"

    for candidate in ("utf-8-sig", "cp1252"):
        try:
            with path.open("r", encoding=candidate) as source_file:
                for _ in source_file:
                    pass
        except UnicodeDecodeError:
            continue
        return candidate

    raise ValueError(
        f"Could not detect the text encoding for {path}; pass it with --encoding"
    )


def read_tab_delimited_text(path: Path, encoding: str) -> pd.DataFrame:
    """Read a tab-delimited text file whose first row contains column names."""
    actual_encoding = resolve_text_encoding(path, encoding)
    try:
        with path.open("r", encoding=actual_encoding, newline="") as text_file:
            header = next(csv.reader(text_file, delimiter="\t"), None)
        if not header:
            raise ValueError(f"The text file is empty or has no header row: {path}")

        columns = [column.strip() for column in header]
        if any(not column for column in columns):
            raise ValueError("The text header contains a blank column name")
        duplicates = sorted({column for column in columns if columns.count(column) > 1})
        if duplicates:
            raise ValueError("Duplicate text header name(s): " + ", ".join(duplicates))

        dataframe = pd.read_csv(
            path,
            sep="\t",
            header=0,
            dtype=str,
            keep_default_na=False,
            na_values=[""],
            encoding=actual_encoding,
        )
    except UnicodeDecodeError as exc:
        raise ValueError(
            f"Could not decode {path} using {actual_encoding!r}; set --encoding "
            "to the file's encoding"
        ) from exc

    if len(dataframe.columns) != len(columns):
        raise ValueError("The text header and parsed data have different column counts")
    dataframe.columns = columns
    return validate_columns(dataframe)


def convert_text_to_excel(
    source: Path,
    output: Path,
    encoding: str,
    overwrite: bool,
) -> pd.DataFrame:
    """Write a tab-delimited text file as an Excel workbook and return its data."""
    if output.exists() and not overwrite:
        raise FileExistsError(
            f"Converted workbook already exists: {output}. "
            "Pass --overwrite to replace it."
        )

    dataframe = add_source_file_column(read_tab_delimited_text(source, encoding), source)
    try:
        dataframe.to_excel(output, index=False, header=True)
    except ImportError as exc:
        raise RuntimeError(
            "The Excel writer is missing; run 'pip install -r requirements.txt'"
        ) from exc
    return dataframe


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Load one Excel workbook, one tab-delimited .txt file, or combine "
            "supported files directly inside a folder and load them to Snowflake."
        )
    )
    parser.add_argument(
        "path",
        type=Path,
        help="Excel/TXT file path or folder of supported files",
    )
    parser.add_argument(
        "--if-exists",
        choices=("fail", "append", "replace"),
        default="fail",
        help="behavior when the Snowflake table exists (default: fail)",
    )
    parser.add_argument(
        "--chunk-size",
        type=int,
        default=10_000,
        help="rows per Snowflake upload chunk",
    )
    parser.add_argument(
        "--encoding",
        default="auto",
        help="TXT encoding, or 'auto' to detect it (default: auto)",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="replace an existing converted .xlsx file",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="convert/validate the input without connecting to Snowflake",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.chunk_size <= 0:
        parser.error("--chunk-size must be greater than zero")

    sources, table_name = resolve_input(args.path)
    dataframes = []
    for source in sources:
        if source.suffix.casefold() == ".txt":
            if args.dry_run:
                dataframe = add_source_file_column(
                    read_tab_delimited_text(source, args.encoding),
                    source,
                )
            else:
                output = source.with_suffix(".xlsx")
                dataframe = convert_text_to_excel(
                    source,
                    output,
                    encoding=args.encoding,
                    overwrite=args.overwrite,
                )
                print(f"CONVERTED {source.name} -> {output}")
        else:
            dataframe = add_source_file_column(read_excel_sheet(source), source)
        dataframes.append(dataframe)

    dataframe = pd.concat(dataframes, ignore_index=True, sort=False)
    dataframe = prepare_temporal_columns(validate_columns(dataframe))
    if args.dry_run:
        print(
            f"VALID {len(sources)} file(s): table={table_name}, "
            f"rows={len(dataframe):,}, columns={len(dataframe.columns):,}; "
            f"source files: {', '.join(source.name for source in sources)}"
        )
        return 0

    config = SnowflakeTargetConfig.from_env(
        schema_env_var="SNOWFLAKE_GARTH_SCHEMA"
    )
    connection = config.connect()
    try:
        loader = ExcelToSnowflakeLoader(
            connection,
            config,
            chunk_size=args.chunk_size,
        )
        rows_loaded = loader.load(
            dataframe,
            table_name,
            mode=args.if_exists,
            add_missing_columns=True,
        )
    finally:
        connection.close()

    print(f"LOADED {len(sources)} file(s): {rows_loaded:,} row(s) -> {table_name}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (
        FileExistsError,
        FileNotFoundError,
        ImportError,
        OSError,
        RuntimeError,
        ValueError,
    ) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc
