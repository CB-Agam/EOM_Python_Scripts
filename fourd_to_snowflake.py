"""One-time, sequential full loads from 4D into Snowflake."""

from __future__ import annotations

import datetime as dt
import decimal
import hashlib
import os
import time
import uuid
from dataclasses import dataclass
from typing import Any, Callable, Iterable


DataFrameFactory = Callable[..., Any]
WriteDataFrame = Callable[..., tuple[bool, int, int, Any]]
ProgressCallback = Callable[[str], None]


def quote_snowflake_identifier(identifier: str) -> str:
    """Return a Snowflake identifier with safe ANSI quoting."""
    if not identifier:
        raise ValueError("Snowflake identifiers cannot be empty")
    return '"' + identifier.replace('"', '""') + '"'


def quote_4d_identifier(identifier: str) -> str:
    """Return a bracket-quoted 4D SQL identifier."""
    if not identifier:
        raise ValueError("4D identifiers cannot be empty")
    return "[" + identifier.replace("]", "]]") + "]"


def snowflake_type(description: tuple[Any, ...]) -> str:
    """Map a DB-API cursor-description entry to a Snowflake column type."""
    type_code = description[1] if len(description) > 1 else None
    precision = description[4] if len(description) > 4 else None
    scale = description[5] if len(description) > 5 else None

    if type_code is bool:
        return "BOOLEAN"
    if type_code is int:
        return "NUMBER(38,0)"
    if type_code is float:
        return "FLOAT"
    if type_code is decimal.Decimal:
        numeric_precision = min(max(int(precision or 38), 1), 38)
        numeric_scale = min(max(int(scale or 0), 0), numeric_precision)
        return f"NUMBER({numeric_precision},{numeric_scale})"
    if type_code is dt.datetime:
        return "TIMESTAMP_NTZ"
    if type_code is dt.date:
        return "DATE"
    if type_code is dt.time:
        return "TIME"
    if type_code in {bytes, bytearray, memoryview}:
        return "BINARY"
    return "TEXT"


def is_sensitive_table(table_name: str) -> bool:
    """Identify tables that should not be copied by an implicit all-table run."""
    normalized = table_name.casefold()
    return any(word in normalized for word in ("password", "credential", "secret"))


@dataclass(frozen=True)
class SnowflakeTargetConfig:
    user: str
    account: str
    warehouse: str
    role: str
    database: str
    schema: str
    authenticator: str = "externalbrowser"
    password: str | None = None
    private_key_file: str | None = None
    private_key_file_pwd: str | None = None

    @classmethod
    def from_env(
        cls,
        schema_env_var: str = "SNOWFLAKE_SCHEMA",
    ) -> "SnowflakeTargetConfig":
        try:
            from dotenv import load_dotenv
        except ImportError as exc:
            raise RuntimeError(
                "python-dotenv is not installed; run 'pip install -r requirements.txt'"
            ) from exc

        load_dotenv()

        def required(name: str) -> str:
            value = os.getenv(name)
            if not value:
                raise ValueError(f"Missing required environment setting {name}")
            return value

        return cls(
            user=required("SNOWFLAKE_USER"),
            account=required("SNOWFLAKE_ACCOUNT"),
            warehouse=required("SNOWFLAKE_WAREHOUSE"),
            role=required("SNOWFLAKE_ROLE"),
            database=required("SNOWFLAKE_DATABASE"),
            schema=required(schema_env_var),
            authenticator=os.getenv("SNOWFLAKE_AUTHENTICATOR", "externalbrowser"),
            password=os.getenv("SNOWFLAKE_PASSWORD") or None,
            private_key_file=os.getenv("SNOWFLAKE_PRIVATE_KEY_FILE") or None,
            private_key_file_pwd=os.getenv("SNOWFLAKE_PRIVATE_KEY_FILE_PWD") or None,
        )

    def connect(self) -> Any:
        try:
            import snowflake.connector
        except ImportError as exc:
            raise RuntimeError(
                "Snowflake connector is not installed; run "
                "'pip install -r requirements.txt'"
            ) from exc

        parameters: dict[str, Any] = {
            "user": self.user,
            "account": self.account,
            "warehouse": self.warehouse,
            "role": self.role,
            "database": self.database,
            "authenticator": self.authenticator,
        }
        if self.password:
            parameters["password"] = self.password
        if self.private_key_file:
            parameters["private_key_file"] = self.private_key_file
        if self.private_key_file_pwd:
            parameters["private_key_file_pwd"] = self.private_key_file_pwd

        # The schema is deliberately omitted because it may not exist yet.
        return snowflake.connector.connect(**parameters)


@dataclass(frozen=True)
class TableLoadResult:
    source_table: str
    target_table: str
    status: str
    rows_loaded: int = 0
    error: str | None = None


class FourDToSnowflakeLoader:
    """Load complete 4D tables to Snowflake without parallel source queries."""

    def __init__(
        self,
        source_connection: Any,
        target_connection: Any,
        target_config: SnowflakeTargetConfig,
        batch_size: int = 5_000,
        pause_seconds: float = 0.0,
        dataframe_factory: DataFrameFactory | None = None,
        write_dataframe: WriteDataFrame | None = None,
        progress: ProgressCallback | None = None,
    ) -> None:
        if batch_size <= 0:
            raise ValueError("batch_size must be greater than zero")
        if pause_seconds < 0:
            raise ValueError("pause_seconds cannot be negative")

        self.source_connection = source_connection
        self.target_connection = target_connection
        self.target_config = target_config
        self.batch_size = batch_size
        self.pause_seconds = pause_seconds
        self._dataframe_factory = dataframe_factory
        self._write_dataframe = write_dataframe
        self.progress = progress or (lambda message: None)

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
        cursor = self.target_connection.cursor()
        try:
            cursor.execute(f"CREATE SCHEMA IF NOT EXISTS {self.target_namespace}")
        finally:
            cursor.close()

    def target_table_exists(self, table_name: str) -> bool:
        database = quote_snowflake_identifier(self.target_config.database)
        cursor = self.target_connection.cursor()
        try:
            cursor.execute(
                f"SELECT 1 FROM {database}.INFORMATION_SCHEMA.TABLES "
                "WHERE TABLE_SCHEMA = %s AND TABLE_NAME = %s",
                (self.target_config.schema, table_name),
            )
            return cursor.fetchone() is not None
        finally:
            cursor.close()

    def _staging_table_name(self, source_table: str) -> str:
        digest = hashlib.sha256(source_table.encode("utf-8")).hexdigest()[:10]
        run_suffix = uuid.uuid4().hex[:8]
        return f"_4D_LOAD_{digest}_{run_suffix}"

    def _create_staging_table(
        self,
        staging_table: str,
        description: Iterable[tuple[Any, ...]],
    ) -> None:
        columns = [
            f"{quote_snowflake_identifier(str(column[0]))} {snowflake_type(column)}"
            for column in description
        ]
        if not columns:
            raise RuntimeError("4D returned no columns for the source table")

        cursor = self.target_connection.cursor()
        try:
            cursor.execute(
                f"CREATE TRANSIENT TABLE {self._qualified_table(staging_table)} "
                f"({', '.join(columns)})"
            )
        finally:
            cursor.close()

    def _drop_staging_table(self, staging_table: str) -> None:
        cursor = self.target_connection.cursor()
        try:
            cursor.execute(
                f"DROP TABLE IF EXISTS {self._qualified_table(staging_table)}"
            )
        finally:
            cursor.close()

    def _publish_staging_table(self, staging_table: str, target_table: str) -> None:
        cursor = self.target_connection.cursor()
        try:
            cursor.execute(
                f"CREATE OR REPLACE TABLE {self._qualified_table(target_table)} "
                f"CLONE {self._qualified_table(staging_table)}"
            )
        finally:
            cursor.close()

    def _make_dataframe(self, rows: list[Any], columns: list[str]) -> Any:
        if self._dataframe_factory is None:
            try:
                import pandas as pd
            except ImportError as exc:
                raise RuntimeError(
                    "pandas is not installed; run 'pip install -r requirements.txt'"
                ) from exc
            self._dataframe_factory = pd.DataFrame
        return self._dataframe_factory.from_records(rows, columns=columns)

    def _write_batch(self, dataframe: Any, staging_table: str) -> int:
        if self._write_dataframe is None:
            try:
                from snowflake.connector.pandas_tools import write_pandas
            except ImportError as exc:
                raise RuntimeError(
                    "Snowflake pandas support is not installed; run "
                    "'pip install -r requirements.txt'"
                ) from exc
            self._write_dataframe = write_pandas

        success, _, rows_written, output = self._write_dataframe(
            self.target_connection,
            dataframe,
            staging_table,
            database=self.target_config.database,
            schema=self.target_config.schema,
            quote_identifiers=True,
            chunk_size=self.batch_size,
        )
        if not success or rows_written != len(dataframe):
            raise RuntimeError(
                f"Snowflake wrote {rows_written} of {len(dataframe)} rows: {output}"
            )
        return int(rows_written)

    def load_table(self, source_table: str, replace_existing: bool = False) -> int:
        """Fully load one source table and atomically publish its target copy."""
        if self.target_table_exists(source_table) and not replace_existing:
            raise FileExistsError(
                f"Target table {self._qualified_table(source_table)} already exists"
            )

        source_cursor = self.source_connection.cursor()
        staging_table = self._staging_table_name(source_table)
        staging_created = False
        try:
            source_cursor.execute(
                f"SELECT * FROM {quote_4d_identifier(source_table)}"
            )
            description = list(source_cursor.description or [])
            columns = [str(column[0]) for column in description]
            if len(columns) != len(set(columns)):
                raise RuntimeError(
                    f"Source table {source_table!r} contains duplicate column names"
                )

            self._create_staging_table(staging_table, description)
            staging_created = True
            rows_loaded = 0

            while True:
                rows = source_cursor.fetchmany(self.batch_size)
                if not rows:
                    break
                dataframe = self._make_dataframe(rows, columns)
                rows_loaded += self._write_batch(dataframe, staging_table)
                self.progress(
                    f"{source_table}: {rows_loaded:,} row(s) transferred"
                )

            self._publish_staging_table(staging_table, source_table)
            return rows_loaded
        finally:
            source_cursor.close()
            if staging_created:
                self._drop_staging_table(staging_table)

    def load_tables(
        self,
        table_names: Iterable[str],
        if_exists: str = "fail",
    ) -> list[TableLoadResult]:
        """Load tables sequentially and retain a result for every requested table."""
        if if_exists not in {"fail", "replace", "skip"}:
            raise ValueError("if_exists must be one of: fail, replace, skip")

        names = list(dict.fromkeys(table_names))
        self.ensure_target_schema()

        if if_exists == "fail":
            conflicts = [name for name in names if self.target_table_exists(name)]
            if conflicts:
                rendered = ", ".join(repr(name) for name in conflicts[:10])
                if len(conflicts) > 10:
                    rendered += f", and {len(conflicts) - 10} more"
                raise FileExistsError(
                    "No tables were loaded because target tables already exist: "
                    + rendered
                )

        results: list[TableLoadResult] = []
        for index, table_name in enumerate(names, start=1):
            self.progress(f"[{index}/{len(names)}] Loading {table_name}")
            if if_exists == "skip" and self.target_table_exists(table_name):
                results.append(
                    TableLoadResult(table_name, table_name, "SKIPPED")
                )
                self.progress(f"{table_name}: skipped because target exists")
                continue

            try:
                row_count = self.load_table(
                    table_name,
                    replace_existing=if_exists == "replace",
                )
                results.append(
                    TableLoadResult(
                        table_name,
                        table_name,
                        "LOADED",
                        rows_loaded=row_count,
                    )
                )
                self.progress(f"{table_name}: completed with {row_count:,} row(s)")
            except Exception as exc:
                results.append(
                    TableLoadResult(
                        table_name,
                        table_name,
                        "FAILED",
                        error=str(exc),
                    )
                )
                self.progress(f"{table_name}: FAILED - {exc}")

            if self.pause_seconds and index < len(names):
                time.sleep(self.pause_seconds)

        return results