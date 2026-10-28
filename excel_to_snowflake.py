"""Reusable Excel-to-Snowflake loading helpers."""

from __future__ import annotations

import datetime as dt
import decimal
import re
import uuid
import warnings
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

import pandas as pd
from dateutil.parser import parse as parse_datetime
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
TEXT_ONLY_COLUMNS = {
    "current end date",
    "drivers date of birth",
    "policy #",
    "policy no",
    "policy number",
}
SNOWFLAKE_TEXT_TYPES = {"CHAR", "CHARACTER", "STRING", "TEXT", "VARCHAR"}


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


def normalize_text_columns(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Convert text and mixed text/object columns to Arrow-compatible strings."""
    prepared = dataframe.copy()
    for column in prepared.columns:
        series = prepared[column]
        if str(column).strip().casefold() in TEXT_ONLY_COLUMNS:
            prepared[column] = series.astype("string")
            continue
        if not pandas_types.is_object_dtype(series.dtype):
            continue

        contains_text = any(
            isinstance(value, str)
            for value in series.array
            if not pd.isna(value)
        )
        if contains_text or dataframe_snowflake_type(series) == "TEXT":
            prepared[column] = series.astype("string")
    return prepared


def coerce_series_to_snowflake_type(
    series: pd.Series,
    snowflake_type: str,
    column_name: str,
) -> pd.Series:
    """Convert incoming values to the type declared by the target table."""
    target_type = snowflake_type.upper().split("(", 1)[0]
    if target_type in SNOWFLAKE_TEXT_TYPES:
        return series.astype("string")

    nonblank = series.map(
        lambda value: not pd.isna(value)
        and not (isinstance(value, str) and not value.strip())
    )
    values = series.where(nonblank, None)

    if target_type in {
        "NUMBER",
        "DECIMAL",
        "NUMERIC",
        "FIXED",
        "INT",
        "INTEGER",
        "BIGINT",
        "SMALLINT",
        "TINYINT",
        "BYTEINT",
        "FLOAT",
        "REAL",
        "DOUBLE",
    }:
        converted = pd.to_numeric(values, errors="coerce")
        invalid = nonblank & converted.isna()
    elif target_type == "BOOLEAN":
        boolean_values = {
            "true": True,
            "false": False,
            "1": True,
            "0": False,
        }

        def parse_boolean(value: Any) -> bool | None:
            if pd.isna(value):
                return None
            if isinstance(value, bool):
                return value
            if isinstance(value, (int, float)) and value in (0, 1):
                return bool(value)
            if isinstance(value, str) and value.strip().casefold() in boolean_values:
                return boolean_values[value.strip().casefold()]
            return None

        converted = values.map(parse_boolean)
        invalid = nonblank & converted.isna()
    elif target_type in {"DATE", "DATETIME"} or target_type.startswith("TIMESTAMP"):
        timestamp_values: list[dt.datetime | None] = []
        invalid_values: list[str] = []
        for value, should_convert in zip(values.tolist(), nonblank.tolist()):
            if not should_convert:
                timestamp_values.append(None)
                continue
            try:
                if isinstance(value, dt.datetime):
                    parsed_value = value
                elif isinstance(value, dt.date):
                    parsed_value = dt.datetime.combine(value, dt.time())
                else:
                    parsed_value = pd.to_datetime(
                        value,
                        errors="coerce",
                        format="mixed",
                    )
                    if pd.isna(parsed_value):
                        parsed_value = parse_datetime(str(value))
                    else:
                        parsed_value = parsed_value.to_pydatetime()
                timestamp_values.append(parsed_value)
            except (OverflowError, TypeError, ValueError):
                timestamp_values.append(None)
                invalid_values.append(str(value))

        if invalid_values:
            examples = list(dict.fromkeys(invalid_values))[:5]
            raise ValueError(
                f"Column {column_name!r} has {len(invalid_values)} value(s) "
                f"that cannot be converted to Snowflake {snowflake_type}: {examples}"
            )
        if target_type == "DATE":
            converted = pd.Series(
                [
                    value.date() if value is not None else None
                    for value in timestamp_values
                ],
                index=series.index,
                dtype=object,
            )
        else:
            converted = pd.Series(
                timestamp_values,
                index=series.index,
                dtype=object,
            )
        invalid = pd.Series(False, index=series.index)
    elif target_type == "TIME":
        converted_values: list[dt.time | None] = []
        invalid_values = []
        for value, should_convert in zip(values.tolist(), nonblank.tolist()):
            if not should_convert:
                converted_values.append(None)
                continue
            try:
                if isinstance(value, dt.time):
                    converted_values.append(value)
                elif isinstance(value, dt.datetime):
                    converted_values.append(value.time())
                elif isinstance(value, str):
                    converted_values.append(parse_datetime(value).time())
                else:
                    raise ValueError("unsupported time value")
            except (OverflowError, TypeError, ValueError):
                converted_values.append(None)
                invalid_values.append(str(value))
        if invalid_values:
            examples = list(dict.fromkeys(invalid_values))[:5]
            raise ValueError(
                f"Column {column_name!r} has {len(invalid_values)} value(s) "
                f"that cannot be converted to Snowflake {snowflake_type}: {examples}"
            )
        converted = pd.Series(
            converted_values,
            index=series.index,
            dtype=object,
        )
        invalid = pd.Series(False, index=series.index)
    else:
        return series

    if invalid.any():
        examples = series.loc[invalid].astype(str).drop_duplicates().head(5).tolist()
        raise ValueError(
            f"Column {column_name!r} has {int(invalid.sum())} value(s) "
            f"that cannot be converted to Snowflake {snowflake_type}: {examples}"
        )
    return converted


def prepare_temporal_columns(
    dataframe: pd.DataFrame,
    source_name: str | None = None,
) -> pd.DataFrame:
    """Validate date-like columns and prepare logical date/timestamp values."""
    prepared = dataframe.copy()
    for column in prepared.columns:
        series = prepared[column]
        if str(column).strip().casefold() in TEXT_ONLY_COLUMNS:
            prepared[column] = series.astype("string")
            continue

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
        parsed_values: list[dt.datetime | None] = []
        invalid_values: list[str] = []
        for value, parsed_value in zip(cleaned.tolist(), parsed.tolist()):
            if pd.isna(value):
                parsed_values.append(None)
                continue

            if pd.isna(parsed_value):
                try:
                    parsed_values.append(parse_datetime(str(value)))
                except (OverflowError, TypeError, ValueError):
                    parsed_values.append(None)
                    invalid_values.append(str(value))
            else:
                parsed_values.append(parsed_value.to_pydatetime())

        if invalid_values:
            location = f" in {source_name}" if source_name else ""
            examples = list(dict.fromkeys(invalid_values))[:5]
            warnings.warn(
                f"Date-like column {column!r}{location} contains "
                f"{len(invalid_values)} invalid value(s); loading those cells "
                f"as NULL. Examples: {examples}",
                RuntimeWarning,
                stacklevel=2,
            )

        valid_values = [value for value in parsed_values if value is not None]
        if valid_values:
            has_time = any(value.time() != dt.time() for value in valid_values)
            if has_time:
                prepared[column] = pd.Series(
                    parsed_values,
                    index=prepared.index,
                    dtype=object,
                )
            else:
                prepared[column] = pd.Series(
                    [
                        value.date() if value is not None else None
                        for value in parsed_values
                    ],
                    index=prepared.index,
                    dtype=object,
                )
        elif invalid_values or missing_marker.any():
            prepared[column] = pd.Series(
                pd.NaT,
                index=prepared.index,
                dtype="datetime64[ns]",
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

    def target_column_types(self, table_name: str) -> dict[str, str]:
        database = quote_snowflake_identifier(self.target_config.database)
        cursor = self.connection.cursor()
        try:
            cursor.execute(
                f"SELECT COLUMN_NAME, DATA_TYPE FROM {database}.INFORMATION_SCHEMA.COLUMNS "
                "WHERE TABLE_SCHEMA = %s AND TABLE_NAME = %s "
                "ORDER BY ORDINAL_POSITION",
                (self.target_config.schema, table_name),
            )
            return {str(row[0]): str(row[1]).upper() for row in cursor.fetchall()}
        finally:
            cursor.close()

    def _widen_column_to_varchar(self, table_name: str, column: str) -> None:
        table = self._qualified_table(table_name)
        suffix = uuid.uuid4().hex[:12].upper()
        replacement_column = f"_GARTH_TEXT_{suffix}"
        backup_column = f"_GARTH_OLD_{suffix}"
        quoted_column = quote_snowflake_identifier(column)
        quoted_replacement = quote_snowflake_identifier(replacement_column)
        quoted_backup = quote_snowflake_identifier(backup_column)

        cursor = self.connection.cursor()
        try:
            cursor.execute(
                f"ALTER TABLE {table} ADD COLUMN {quoted_replacement} VARCHAR"
            )
            cursor.execute(
                f"UPDATE {table} SET {quoted_replacement} = "
                f"TO_VARCHAR({quoted_column})"
            )
            cursor.execute(
                f"ALTER TABLE {table} RENAME COLUMN {quoted_column} TO {quoted_backup}"
            )
            try:
                cursor.execute(
                    f"ALTER TABLE {table} RENAME COLUMN "
                    f"{quoted_replacement} TO {quoted_column}"
                )
            except Exception as rename_error:
                try:
                    cursor.execute(
                        f"ALTER TABLE {table} RENAME COLUMN "
                        f"{quoted_backup} TO {quoted_column}"
                    )
                except Exception as restore_error:
                    raise RuntimeError(
                        f"Could not finish the VARCHAR migration for {column!r} "
                        f"or restore its original name. The original data remains "
                        f"in {backup_column!r}; replacement data remains in "
                        f"{replacement_column!r}. Restore error: {restore_error}"
                    ) from rename_error
                raise
            cursor.execute(f"ALTER TABLE {table} DROP COLUMN {quoted_backup}")
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

        dataframe = normalize_text_columns(dataframe)

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

    def _replace_source_file_rows(
        self,
        dataframe: pd.DataFrame,
        table_name: str,
        source_file_name: str,
        source_file_column: str,
    ) -> int:
        if dataframe.empty:
            raise ValueError(
                f"Cannot replace existing rows for {source_file_name!r} with "
                "an empty dataframe"
            )
        if source_file_column not in dataframe.columns:
            raise ValueError(
                f"Dataframe is missing required source file column "
                f"{source_file_column!r}"
            )

        source_values = dataframe[source_file_column]
        source_names = source_values.dropna().astype(str).unique()
        if (
            source_values.isna().any()
            or len(source_names) != 1
            or source_names[0] != source_file_name
        ):
            raise ValueError(
                f"Expected every row to have {source_file_column}="
                f"{source_file_name!r}; found {source_names[:5].tolist()}"
            )

        staging_table = f"_GARTH_RELOAD_{uuid.uuid4().hex[:16].upper()}"
        created = False
        try:
            self._create_table(staging_table, dataframe)
            created = True
            rows_staged = self._write(dataframe, staging_table)

            columns = ", ".join(
                quote_snowflake_identifier(column) for column in dataframe.columns
            )
            target = self._qualified_table(table_name)
            staging = self._qualified_table(staging_table)
            source_column = quote_snowflake_identifier(source_file_column)
            cursor = self.connection.cursor()
            try:
                cursor.execute("BEGIN")
                cursor.execute(
                    f"DELETE FROM {target} WHERE {source_column} = %s",
                    (source_file_name,),
                )
                cursor.execute(
                    f"INSERT INTO {target} ({columns}) "
                    f"SELECT {columns} FROM {staging}"
                )
                cursor.execute("COMMIT")
            except Exception as load_error:
                try:
                    cursor.execute("ROLLBACK")
                except Exception as rollback_error:
                    raise RuntimeError(
                        f"Failed to replace rows for {source_file_name!r} in "
                        f"{table_name!r}, and rollback also failed: {rollback_error}"
                    ) from load_error
                raise
            finally:
                cursor.close()
            return rows_staged
        finally:
            if created:
                self._drop_table(staging_table)

    def load(
        self,
        dataframe: pd.DataFrame,
        table_name: str,
        mode: str = "fail",
        *,
        add_missing_columns: bool = False,
        widen_text_columns: bool = False,
        replace_source_file: bool = False,
        source_file_column: str = "SOURCE_FILE_NAME",
    ) -> int:
        """Load a DataFrame using fail, append, or atomic replace semantics."""
        if mode not in {"fail", "append", "replace"}:
            raise ValueError("mode must be one of: fail, append, replace")
        if not table_name:
            raise ValueError("table_name cannot be empty")

        dataframe = normalize_text_columns(
            prepare_temporal_columns(validate_columns(dataframe))
        )
        self.ensure_target_schema()
        exists = self.table_exists(table_name)

        if mode == "fail" and exists:
            raise FileExistsError(
                f"Target table {self._qualified_table(table_name)} already exists"
            )

        if mode == "append" and exists:
            existing_types = self.target_column_types(table_name)
            existing_columns = list(existing_types)
            supplied_columns = list(dataframe.columns)
            extra = [name for name in supplied_columns if name not in existing_columns]
            if extra and add_missing_columns:
                cursor = self.connection.cursor()
                try:
                    for column in extra:
                        column_type = dataframe_snowflake_type(dataframe[column])
                        cursor.execute(
                            f"ALTER TABLE {self._qualified_table(table_name)} "
                            f"ADD COLUMN {quote_snowflake_identifier(column)} "
                            f"{column_type}"
                        )
                finally:
                    cursor.close()
                existing_types = self.target_column_types(table_name)
                existing_columns = list(existing_types)
                extra = [name for name in supplied_columns if name not in existing_columns]

            if extra:
                raise ValueError(
                    "Excel columns do not match target table (extra: "
                    + ", ".join(extra)
                    + ")"
                )

            columns_to_widen = []
            for column in supplied_columns:
                target_type = existing_types[column].upper()
                if target_type in SNOWFLAKE_TEXT_TYPES:
                    continue
                if column.strip().casefold() in TEXT_ONLY_COLUMNS:
                    if not widen_text_columns:
                        raise ValueError(
                            f"Column {column!r} must be VARCHAR, but the target "
                            f"table defines it as {target_type}"
                        )
                    columns_to_widen.append(column)
                    continue
                try:
                    coerce_series_to_snowflake_type(
                        dataframe[column],
                        target_type,
                        column,
                    )
                except ValueError:
                    if not widen_text_columns:
                        raise
                    columns_to_widen.append(column)
            for column in columns_to_widen:
                self._widen_column_to_varchar(table_name, column)
            if columns_to_widen:
                existing_types = self.target_column_types(table_name)
                existing_columns = list(existing_types)

            missing = [name for name in existing_columns if name not in supplied_columns]
            if missing and not add_missing_columns:
                details = []
                details.append("missing: " + ", ".join(missing))
                raise ValueError(
                    "Excel columns do not match target table ("
                    + "; ".join(details)
                    + ")"
                )
            aligned = dataframe.reindex(columns=existing_columns)
            for column, target_type in existing_types.items():
                aligned[column] = coerce_series_to_snowflake_type(
                    aligned[column],
                    target_type,
                    column,
                )
            if replace_source_file:
                if source_file_column not in aligned.columns:
                    raise ValueError(
                        f"Dataframe is missing required source file column "
                        f"{source_file_column!r}"
                    )
                source_names = aligned[source_file_column].dropna().astype(str).unique()
                if len(source_names) != 1:
                    raise ValueError(
                        "Replacing source-file rows requires exactly one non-empty "
                        f"{source_file_column} value; found {source_names[:5].tolist()}"
                    )
                return self._replace_source_file_rows(
                    aligned,
                    table_name,
                    source_names[0],
                    source_file_column,
                )
            return self._write(aligned, table_name)

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
