"""Reusable Excel-to-Snowflake loading helpers."""

from __future__ import annotations

import datetime as dt
import decimal
import re
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

import pandas as pd
from pandas.api import types as pandas_types
from fourd_to_snowflake import (
    SnowflakeTargetConfig,
    quote_snowflake_identifier,
)


WriteDataFrame = Callable[..., tuple[bool, int, int, Any]]

DATE_COLUMN_PATTERN = re.compile(
    r"(?:^|[ _])(date|on)(?:$|[ _])|"
    r"(?:inception|effective|expiration|reviewed|closed)",
    re.IGNORECASE,
)
MISSING_DATE_MARKERS = {"00/00/00", "00/00/0000"}


def default_table_name(path: Path) -> str:
    """Use the Excel base file name with whitespace replaced by underscores."""
    name = re.sub(r"\s+", "_", path.stem.strip())
    if not name:
        raise ValueError("Could not derive a table name from the Excel file name")
    if len(name) > 255:
        raise ValueError("Snowflake table names cannot exceed 255 characters")
    return name


def validate_columns(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Trim and validate column labels while preserving their spelling and case."""
    if len(dataframe.columns) == 0:
        raise ValueError("The worksheet does not contain a header row")

    columns = [str(column).strip() for column in dataframe.columns]
    blank = [index + 1 for index, column in enumerate(columns) if not column]
    if blank:
        raise ValueError(f"Blank column name(s) at position(s): {blank}")

    duplicates = sorted({column for column in columns if columns.count(column) > 1})
    if duplicates:
        raise ValueError("Duplicate column name(s): " + ", ".join(duplicates))

    prepared = dataframe.copy()
    prepared.columns = columns
    return prepared


def dataframe_snowflake_type(series: pd.Series) -> str:
    """Infer a conservative Snowflake type for a pandas Series."""
    dtype = series.dtype
    if pandas_types.is_bool_dtype(dtype):
        return "BOOLEAN"
    if pandas_types.is_integer_dtype(dtype):
        return "NUMBER(38,0)"
    if pandas_types.is_float_dtype(dtype):
        return "FLOAT"
    if pandas_types.is_datetime64_any_dtype(dtype):
        values = series.dropna()
        if values.empty or values.dt.normalize().equals(values):
            return "DATE"
        return "TIMESTAMP_NTZ"
    if pandas_types.is_timedelta64_dtype(dtype):
        return "TIME"

    values = series.dropna()
    if not values.empty:
        sample = values.iloc[:1000]
        if sample.map(lambda value: isinstance(value, bool)).all():
            return "BOOLEAN"
        if sample.map(
            lambda value: isinstance(value, dt.datetime)
        ).all():
            return "TIMESTAMP_NTZ"
        if sample.map(
            lambda value: isinstance(value, dt.date) and not isinstance(value, dt.datetime)
        ).all():
            return "DATE"
        if sample.map(lambda value: isinstance(value, dt.time)).all():
            return "TIME"
        if sample.map(lambda value: isinstance(value, decimal.Decimal)).all():
            return "NUMBER(38,9)"
    return "TEXT"


def prepare_temporal_columns(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Validate date-like columns and prepare logical date/timestamp values."""
    prepared = dataframe.copy()
    for column in prepared.columns:
        series = prepared[column]
        if pandas_types.is_datetime64_any_dtype(series.dtype):
            values = series.dropna()
            if values.empty or values.dt.normalize().equals(values):
                prepared[column] = series.dt.date
            continue

        if not pandas_types.is_object_dtype(series.dtype):
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
        if invalid.any():
            samples = cleaned[invalid].astype(str).drop_duplicates().head(5).tolist()
            raise ValueError(
                f"Date-like column {column!r} contains {int(invalid.sum())} "
                f"unparseable value(s); examples: {samples}"
            )
        if parsed.notna().any():
            values = parsed.dropna()
            prepared[column] = (
                parsed.dt.date if values.dt.normalize().equals(values) else parsed
            )
    return prepared


def read_excel_sheet(
    path: Path,
    sheet: str | int = 0,
    header_row: int = 1,
) -> pd.DataFrame:
    """Read and validate one worksheet from an Excel workbook."""
    if not path.is_file():
        raise FileNotFoundError(f"Excel file not found: {path}")
    if path.suffix.casefold() not in {".xlsx", ".xlsm", ".xls"}:
        raise ValueError("Input must be an .xlsx, .xlsm, or .xls file")
    if header_row <= 0:
        raise ValueError("header_row must be at least 1")

    try:
        dataframe = pd.read_excel(path, sheet_name=sheet, header=header_row - 1)
    except ImportError as exc:
        raise RuntimeError(
            "The Excel reader is missing; run 'pip install -r requirements.txt'"
        ) from exc
    return prepare_temporal_columns(validate_columns(dataframe))


@dataclass(frozen=True)
class ExcelLoadResult:
    file: str
    sheet: str
    target_table: str
    rows_loaded: int
    mode: str


class ExcelToSnowflakeLoader:
    """Load a prepared DataFrame into a Snowflake table."""

    def __init__(
        self,
        connection: Any,
        target_config: SnowflakeTargetConfig,
        chunk_size: int = 10_000,
        write_dataframe: WriteDataFrame | None = None,
    ) -> None:
        if chunk_size <= 0:
            raise ValueError("chunk_size must be greater than zero")
        self.connection = connection
        self.target_config = target_config
        self.chunk_size = chunk_size
        self._write_dataframe = write_dataframe

    @property
    def target_namespace(self) -> str:
        return ".".join(
            (
                quote_snowflake_identifier(self.target_config.database),
                quote_snowflake_identifier(self.target_config.schema),
            )
        )

    def _qualified_table(self, table_name: str) -> str:
        return f"{self.target_namespace}.{quote_snowflake_identifier(table_name)}"

    def ensure_target_schema(self) -> None:
        cursor = self.connection.cursor()
        try:
            cursor.execute(f"CREATE SCHEMA IF NOT EXISTS {self.target_namespace}")
        finally:
            cursor.close()

    def table_exists(self, table_name: str) -> bool:
        database = quote_snowflake_identifier(self.target_config.database)
        cursor = self.connection.cursor()
        try:
            cursor.execute(
                f"SELECT 1 FROM {database}.INFORMATION_SCHEMA.TABLES "
                "WHERE TABLE_SCHEMA = %s AND TABLE_NAME = %s",
                (self.target_config.schema, table_name),
            )
            return cursor.fetchone() is not None
        finally:
            cursor.close()

    def target_columns(self, table_name: str) -> list[str]:
        database = quote_snowflake_identifier(self.target_config.database)
        cursor = self.connection.cursor()
        try:
            cursor.execute(
                f"SELECT COLUMN_NAME FROM {database}.INFORMATION_SCHEMA.COLUMNS "
                "WHERE TABLE_SCHEMA = %s AND TABLE_NAME = %s "
                "ORDER BY ORDINAL_POSITION",
                (self.target_config.schema, table_name),
            )
            return [str(row[0]) for row in cursor.fetchall()]
        finally:
            cursor.close()

    def _create_table(self, table_name: str, dataframe: pd.DataFrame) -> None:
        definitions = [
            f"{quote_snowflake_identifier(column)} "
            f"{dataframe_snowflake_type(dataframe[column])}"
            for column in dataframe.columns
        ]
        cursor = self.connection.cursor()
        try:
            # Snowflake cannot clone a transient object into the permanent target.
            # Use a short-lived permanent staging table and always drop it in load().
            cursor.execute(
                f"CREATE TABLE {self._qualified_table(table_name)} "
                f"({', '.join(definitions)})"
            )
        finally:
            cursor.close()

    def _drop_table(self, table_name: str) -> None:
        cursor = self.connection.cursor()
        try:
            cursor.execute(f"DROP TABLE IF EXISTS {self._qualified_table(table_name)}")
        finally:
            cursor.close()

    def _publish(self, staging_table: str, target_table: str) -> None:
        cursor = self.connection.cursor()
        try:
            cursor.execute(
                f"CREATE OR REPLACE TABLE {self._qualified_table(target_table)} "
                f"CLONE {self._qualified_table(staging_table)}"
            )
        finally:
            cursor.close()

    def _write(self, dataframe: pd.DataFrame, table_name: str) -> int:
        if dataframe.empty:
            return 0
        if self._write_dataframe is None:
            from snowflake.connector.pandas_tools import write_pandas

            self._write_dataframe = write_pandas

        success, _, rows_written, output = self._write_dataframe(
            self.connection,
            dataframe,
            table_name,
            database=self.target_config.database,
            schema=self.target_config.schema,
            quote_identifiers=True,
            chunk_size=self.chunk_size,
            use_logical_type=True,
        )
        if not success or int(rows_written) != len(dataframe):
            raise RuntimeError(
                f"Snowflake wrote {rows_written} of {len(dataframe)} rows: {output}"
            )
        return int(rows_written)

    def load(self, dataframe: pd.DataFrame, table_name: str, mode: str = "fail") -> int:
        """Load a DataFrame using fail, append, or atomic replace semantics."""
        if mode not in {"fail", "append", "replace"}:
            raise ValueError("mode must be one of: fail, append, replace")
        if not table_name:
            raise ValueError("table_name cannot be empty")

        dataframe = prepare_temporal_columns(validate_columns(dataframe))
        self.ensure_target_schema()
        exists = self.table_exists(table_name)

        if mode == "fail" and exists:
            raise FileExistsError(
                f"Target table {self._qualified_table(table_name)} already exists"
            )

        if mode == "append" and exists:
            existing_columns = self.target_columns(table_name)
            supplied_columns = list(dataframe.columns)
            missing = [name for name in existing_columns if name not in supplied_columns]
            extra = [name for name in supplied_columns if name not in existing_columns]
            if missing or extra:
                details = []
                if missing:
                    details.append("missing: " + ", ".join(missing))
                if extra:
                    details.append("extra: " + ", ".join(extra))
                raise ValueError("Excel columns do not match target table (" + "; ".join(details) + ")")
            return self._write(dataframe[existing_columns], table_name)

        staging_table = f"_EXCEL_LOAD_{uuid.uuid4().hex[:16].upper()}"
        created = False
        try:
            self._create_table(staging_table, dataframe)
            created = True
            rows_loaded = self._write(dataframe, staging_table)
            self._publish(staging_table, table_name)
            return rows_loaded
        finally:
            if created:
                self._drop_table(staging_table)


#env details
#SNOWFLAKE_USER=VENGATESHD@BCICJAMAICA.COM
#SNOWFLAKE_ACCOUNT=zfb92242.us-east-1
#SNOWFLAKE_WAREHOUSE=ANA_DEV_WH
#SNOWFLAKE_ROLE=ANA_DEV_ROLE
#SNOWFLAKE_DATABASE=DW_AGAM_DEV
#SNOWFLAKE_SCHEMA=UWTR_ALL_DBO

#to execute
#cd "C:\Users\Vengatesh\Documents\snowflake_conn\snowflake_conn"

#$env:SNOWFLAKE_SCHEMA = "4D_Report"

#& "..\.venv\Scripts\python.exe" ".\load_excel_to_snowflake.py"   "C:\Users\Vengatesh\Downloads\2025 EOM Reports\2025 EOM Reports\2025-10 EOM Reports\All Claims By YOL_202510.xlsx" `