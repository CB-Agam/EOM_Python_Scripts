"""Convert a tab-delimited text file to Excel and load it into Snowflake."""

from __future__ import annotations

import argparse
import csv
import io
import os
import re
import sys
from pathlib import Path
from typing import Any

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
    return table_name_from_name(source_name)


def table_name_from_name(source_name: str) -> str:
    """Create a Snowflake-safe table name from a Drive file or folder label."""
    name = re.sub(r"\W+", "_", source_name, flags=re.ASCII).strip("_")
    name = re.sub(r"_+", "_", name)
    if not name:
        raise ValueError(f"Cannot derive a table name from {source_name!r}")
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


def decode_text_content(content: bytes, encoding: str, source_name: str) -> str:
    """Decode TXT content with explicit encoding or common encoding detection."""
    if encoding.casefold() != "auto":
        candidates = (encoding,)
    elif content.startswith((b"\x00\x00\xfe\xff", b"\xff\xfe\x00\x00")):
        candidates = ("utf-32",)
    elif content.startswith(b"\xef\xbb\xbf"):
        candidates = ("utf-8-sig",)
    elif content.startswith((b"\xff\xfe", b"\xfe\xff")):
        candidates = ("utf-16",)
    else:
        candidates = ("utf-8-sig", "cp1252")

    for candidate in candidates:
        try:
            return content.decode(candidate)
        except UnicodeDecodeError:
            continue

    raise ValueError(
        f"Could not decode {source_name}; pass its encoding with --encoding"
    )


def read_tab_delimited_content(
    content: bytes,
    encoding: str,
    source_name: str,
) -> pd.DataFrame:
    """Read tab-delimited TXT bytes with column names in the first row."""
    text = decode_text_content(content, encoding, source_name)
    header = next(csv.reader(io.StringIO(text), delimiter="\t"), None)
    if not header:
        raise ValueError(f"The text file is empty or has no header row: {source_name}")

    columns = [column.strip() for column in header]
    if any(not column for column in columns):
        raise ValueError("The text header contains a blank column name")
    keep_indices = []
    seen_columns = set()
    dropped_columns = []
    for index, column in enumerate(columns):
        if column in seen_columns:
            dropped_columns.append(column)
        else:
            seen_columns.add(column)
            keep_indices.append(index)

    if dropped_columns:
        print(
            f"WARNING {source_name}: ignoring duplicate column(s), keeping the "
            f"first occurrence: {', '.join(dropped_columns)}",
            file=sys.stderr,
        )

    dataframe = pd.read_csv(
        io.StringIO(text),
        sep="\t",
        header=0,
        names=[f"__garth_column_{index}" for index in range(len(columns))],
        dtype=str,
        keep_default_na=False,
        na_values=[""],
    )
    if len(dataframe.columns) != len(columns):
        raise ValueError("The text header and parsed data have different column counts")
    dataframe = dataframe.iloc[:, keep_indices].copy()
    dataframe.columns = [columns[index] for index in keep_indices]
    return validate_columns(dataframe)


def read_tab_delimited_text(path: Path, encoding: str) -> pd.DataFrame:
    """Read a tab-delimited text file whose first row contains column names."""
    return read_tab_delimited_content(path.read_bytes(), encoding, str(path))


def detect_excel_header_row(
    content: bytes,
    source_name: str,
    sheet_name: str | int,
) -> int:
    """Find the first row with at least three non-empty text header cells."""
    try:
        preview = pd.read_excel(
            io.BytesIO(content),
            sheet_name=sheet_name,
            header=None,
            nrows=50,
        )
    except ImportError as exc:
        raise RuntimeError(
            "An Excel reader dependency is missing; run 'pip install -r requirements.txt'"
        ) from exc

    for row_index, row in preview.iterrows():
        values = [value for value in row.tolist() if pd.notna(value)]
        text_values = [
            value for value in values if isinstance(value, str) and value.strip()
        ]
        if len(text_values) >= 3 and len(text_values) == len(values):
            return int(row_index)

    raise ValueError(
        f"Could not detect a header row in the first 50 rows of {source_name}; "
        "expected a row containing at least three text column names."
    )


def read_excel_content(
    content: bytes,
    source_name: str = "<Excel file>",
    sheet_name: str | int = 0,
) -> pd.DataFrame:
    """Read a worksheet, detecting and skipping title/preamble rows."""
    try:
        workbook = pd.ExcelFile(io.BytesIO(content))
    except ImportError as exc:
        raise RuntimeError(
            "An Excel reader dependency is missing; run 'pip install -r requirements.txt'"
        ) from exc

    if isinstance(sheet_name, str) and sheet_name not in workbook.sheet_names:
        raise ValueError(
            f"Required worksheet {sheet_name!r} was not found in {source_name}; "
            f"available worksheets: {workbook.sheet_names}"
        )

    selected_sheet = sheet_name
    header_row = detect_excel_header_row(content, source_name, selected_sheet)
    dataframe = pd.read_excel(
        io.BytesIO(content),
        sheet_name=selected_sheet,
        header=header_row,
    )
    selected_sheet_name = (
        workbook.sheet_names[selected_sheet]
        if isinstance(selected_sheet, int)
        else selected_sheet
    )
    print(
        f"INFO {source_name}: using worksheet {selected_sheet_name!r}, "
        f"Excel row {header_row + 1} as column headers"
    )
    return prepare_temporal_columns(
        validate_columns(dataframe),
        source_name=source_name,
    )


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


def required_environment_setting(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise ValueError(f"Missing required environment setting {name}")
    return value


def build_drive_service() -> Any:
    """Authenticate to Google Drive using a local service-account JSON file."""
    try:
        from dotenv import load_dotenv
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
    except ImportError as exc:
        raise RuntimeError(
            "Google Drive dependencies are missing; run 'pip install -r requirements.txt'"
        ) from exc

    load_dotenv()
    credentials_path = Path(
        required_environment_setting("GOOGLE_SERVICE_ACCOUNT_FILE")
    ).expanduser()
    credentials = service_account.Credentials.from_service_account_file(
        str(credentials_path),
        scopes=["https://www.googleapis.com/auth/drive"],
    )
    return build("drive", "v3", credentials=credentials, cache_discovery=False)


def drive_children(service: Any, folder_id: str) -> list[dict[str, Any]]:
    """List direct children of a Google Drive folder, including shared-drive items."""
    files: list[dict[str, Any]] = []
    page_token = None
    while True:
        response = (
            service.files()
            .list(
                q=f"'{folder_id}' in parents and trashed = false",
                pageSize=1000,
                pageToken=page_token,
                fields="nextPageToken,files(id,name,mimeType,parents)",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            )
            .execute()
        )
        files.extend(response.get("files", []))
        page_token = response.get("nextPageToken")
        if not page_token:
            return files


def get_drive_root(service: Any, root_id: str) -> dict[str, Any]:
    """Verify the configured root is a Drive folder visible to this account."""
    try:
        root = (
            service.files()
            .get(
                fileId=root_id,
                fields="id,name,mimeType",
                supportsAllDrives=True,
            )
            .execute()
        )
    except Exception as exc:
        status = getattr(getattr(exc, "resp", None), "status", None)
        if status in {403, 404}:
            raise PermissionError(
                "Google Drive cannot access GOOGLE_DRIVE_ROOT_FOLDER_ID. "
                "Verify it is the folder ID for 'Claims Dashboard Source', and "
                "share that folder with the service account email from the JSON "
                "file configured by GOOGLE_SERVICE_ACCOUNT_FILE as an Editor."
            ) from exc
        raise

    if root.get("mimeType") != "application/vnd.google-apps.folder":
        raise ValueError(
            "GOOGLE_DRIVE_ROOT_FOLDER_ID must point to the 'Claims Dashboard "
            "Source' folder, not a file."
        )
    return root


def find_drive_folder(
    service: Any,
    parent_id: str,
    folder_name: str,
) -> dict[str, Any]:
    matching = [
        item
        for item in drive_children(service, parent_id)
        if item.get("mimeType") == "application/vnd.google-apps.folder"
        and item.get("name", "").casefold() == folder_name.casefold()
    ]
    if not matching:
        raise FileNotFoundError(
            f"Google Drive folder {folder_name!r} was not found under the configured root"
        )
    if len(matching) > 1:
        raise ValueError(
            f"More than one Google Drive folder named {folder_name!r} exists under the configured root"
        )
    return matching[0]


def is_supported_drive_file(item: dict[str, Any]) -> bool:
    name = str(item.get("name", ""))
    return (
        item.get("mimeType") != "application/vnd.google-apps.folder"
        and Path(name).suffix.casefold() in SUPPORTED_EXCEL_SUFFIXES | {".txt"}
        and not name.startswith("~$")
    )


def drive_load_groups(
    service: Any,
    root_id: str,
    branch_name: str,
) -> tuple[list[tuple[str, list[dict[str, Any]]]], str]:
    """Group branch files by file name or immediate child-folder name."""
    branch = find_drive_folder(service, root_id, branch_name)
    children = drive_children(service, branch["id"])
    groups: list[tuple[str, list[dict[str, Any]]]] = []

    for item in children:
        if is_supported_drive_file(item):
            groups.append((table_name_from_name(Path(item["name"]).stem), [item]))

    child_folders = [
        item
        for item in children
        if item.get("mimeType") == "application/vnd.google-apps.folder"
    ]
    for folder in child_folders:
        folder_files = [
            item
            for item in drive_children(service, folder["id"])
            if is_supported_drive_file(item)
        ]
        if folder_files:
            groups.append((table_name_from_name(folder["name"]), folder_files))

    if not groups:
        raise ValueError(
            f"No supported TXT or Excel files found in Google Drive {branch_name} "
            "or its immediate child folders"
        )
    return groups, branch["id"]


def download_drive_file(service: Any, file_id: str) -> bytes:
    try:
        from googleapiclient.http import MediaIoBaseDownload
    except ImportError as exc:
        raise RuntimeError(
            "Google Drive dependencies are missing; run 'pip install -r requirements.txt'"
        ) from exc

    request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buffer.getvalue()


def read_drive_dataframe(
    service: Any,
    item: dict[str, Any],
    encoding: str,
) -> pd.DataFrame:
    content = download_drive_file(service, item["id"])
    name = str(item["name"])
    if Path(name).suffix.casefold() == ".txt":
        dataframe = read_tab_delimited_content(content, encoding, name)
    else:
        sheet_name: str | int = (
            "Data" if "reported claims" in Path(name).stem.casefold() else 0
        )
        dataframe = read_excel_content(content, name, sheet_name=sheet_name)
    return add_source_file_column(dataframe, Path(name))


def archive_drive_files(
    service: Any,
    files: list[dict[str, Any]],
    archive_folder_id: str,
) -> None:
    """Move successfully loaded Drive files into the root archive folder."""
    for item in files:
        service.files().update(
            fileId=item["id"],
            addParents=archive_folder_id,
            removeParents=",".join(item.get("parents", [])),
            fields="id,parents",
            supportsAllDrives=True,
        ).execute()


def process_drive_branch(args: argparse.Namespace) -> int:
    from dotenv import load_dotenv

    load_dotenv()
    root_id = required_environment_setting("GOOGLE_DRIVE_ROOT_FOLDER_ID")
    service = build_drive_service()
    get_drive_root(service, root_id)
    groups, _ = drive_load_groups(service, root_id, args.drive_branch)
    archive_folder = find_drive_folder(service, root_id, "archive")

    config = None
    connection = None
    loader = None
    if not args.dry_run:
        config = SnowflakeTargetConfig.from_env(
            schema_env_var="SNOWFLAKE_GARTH_SCHEMA"
        )
        connection = config.connect()
        loader = ExcelToSnowflakeLoader(
            connection,
            config,
            chunk_size=args.chunk_size,
        )

    failed_files = 0
    total_rows = 0
    try:
        for table_name, files in groups:
            for item in files:
                file_name = str(item.get("name", item.get("id", "<unknown>")))
                try:
                    dataframe = prepare_temporal_columns(
                        validate_columns(
                            read_drive_dataframe(service, item, args.encoding)
                        ),
                        source_name=file_name,
                    )
                    if args.dry_run:
                        print(
                            f"VALID {args.drive_branch}: table={table_name}, "
                            f"file={file_name}, rows={len(dataframe):,}, "
                            f"columns={len(dataframe.columns):,}"
                        )
                        continue

                    if loader is None:
                        raise RuntimeError(
                            "Snowflake loader is unavailable outside dry-run mode"
                        )
                    rows_loaded = loader.load(
                        dataframe,
                        table_name,
                        mode=args.if_exists,
                        add_missing_columns=True,
                        widen_text_columns=True,
                    )
                    total_rows += rows_loaded
                    print(
                        f"LOADED {args.drive_branch}: table={table_name}, "
                        f"file={file_name}, rows={rows_loaded:,}"
                    )
                    archive_drive_files(service, [item], archive_folder["id"])
                    print(f"ARCHIVED file={file_name} -> archive")
                except Exception as exc:
                    failed_files += 1
                    print(
                        f"FAILED {args.drive_branch} table={table_name} "
                        f"file={file_name}: {exc}",
                        file=sys.stderr,
                    )
    finally:
        if connection is not None:
            connection.close()

    if args.dry_run:
        print(
            f"Dry run complete: {len(groups)} table group(s), "
            f"{sum(len(files) for _, files in groups)} file(s), "
            f"{failed_files} failed file(s); nothing was uploaded or archived."
        )
        return 1 if failed_files else 0

    print(
        f"Finished: {len(groups)} table group(s), {total_rows:,} row(s) loaded, "
        f"{failed_files} failed file(s)."
    )
    return 1 if failed_files else 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Load local Excel/TXT files or Google Drive EIS/UW branches to Snowflake. "
            "Drive branch files load as individual tables; files directly inside "
            "each child folder combine into a table named after that folder."
        )
    )
    input_group = parser.add_mutually_exclusive_group()
    input_group.add_argument(
        "path",
        nargs="?",
        type=Path,
        help="local Excel/TXT file path or folder of supported files",
    )
    input_group.add_argument(
        "--drive-branch",
        choices=("EIS", "UW"),
        help="load Google Drive files from the selected root branch",
    )
    parser.add_argument(
        "--if-exists",
        choices=("fail", "append", "replace"),
        default="append",
        help="behavior when the Snowflake table exists (default: append)",
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
    if not args.path and not args.drive_branch:
        parser.error("provide a local path or --drive-branch")
    if args.chunk_size <= 0:
        parser.error("--chunk-size must be greater than zero")

    if args.drive_branch:
        return process_drive_branch(args)

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
