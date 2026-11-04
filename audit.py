from __future__ import annotations

import logging
import traceback
from collections.abc import Callable
from datetime import datetime
from typing import TypeVar
from uuid import uuid4


_T = TypeVar("_T")
logger = logging.getLogger(__name__)


def _quote_snowflake_identifier(identifier: str) -> str:
    if not identifier.strip():
        raise ValueError("Snowflake database and schema names cannot be blank")
    return '"' + identifier.replace('"', '""') + '"'


class SnowflakeAuditLogger:
    def __init__(self, connection, database: str, schema: str):
        self.connection = connection
        self.run_id = str(uuid4())
        self._step_number = 0
        qualified_table = ".".join(
            (
                _quote_snowflake_identifier(database),
                _quote_snowflake_identifier(schema),
                '"ETL_PIPELINE_LOG"',
            )
        )
        self.qualified_table = qualified_table

    def ensure_table(self) -> None:
        sql = f"""
            CREATE TABLE IF NOT EXISTS {self.qualified_table} (
                RUN_ID VARCHAR(36) NOT NULL,
                RUN_STEP_NO NUMBER(38, 0) NOT NULL,
                TABLE_NAME VARCHAR(255),
                STEP VARCHAR(100) NOT NULL,
                STATUS VARCHAR(20) NOT NULL,
                MESSAGE VARCHAR,
                ERROR_MESSAGE VARCHAR,
                START_TIME TIMESTAMP_NTZ,
                END_TIME TIMESTAMP_NTZ,
                TIME_STAMP TIMESTAMP_NTZ DEFAULT CURRENT_TIMESTAMP()
            )
        """
        cursor = self.connection.cursor()
        try:
            cursor.execute(sql)
            cursor.execute(
                f"ALTER TABLE {self.qualified_table} "
                "ADD COLUMN IF NOT EXISTS RUN_STEP_NO NUMBER(38, 0)"
            )
            cursor.execute(
                f"ALTER TABLE {self.qualified_table} DROP COLUMN IF EXISTS S_NO"
            )
            cursor.execute(
                f"ALTER TABLE {self.qualified_table} "
                "ADD COLUMN IF NOT EXISTS START_TIME TIMESTAMP_NTZ"
            )
            cursor.execute(
                f"ALTER TABLE {self.qualified_table} "
                "ADD COLUMN IF NOT EXISTS END_TIME TIMESTAMP_NTZ"
            )
        finally:
            cursor.close()

    def record(
        self,
        step: str,
        status: str,
        table_name: str | None = None,
        message: str | None = None,
        error_message: str | None = None,
        start_time: datetime | None = None,
    ) -> int:
        self._step_number += 1
        step_number = self._step_number
        sql = f"""
            INSERT INTO {self.qualified_table}
                (RUN_ID, RUN_STEP_NO, TABLE_NAME, STEP, STATUS, MESSAGE,
                 ERROR_MESSAGE, START_TIME, END_TIME)
            VALUES (%s, %s, %s, %s, %s, %s, %s,
                    COALESCE(%s, CURRENT_TIMESTAMP()),
                    CASE WHEN %s = 'RUNNING' THEN NULL ELSE CURRENT_TIMESTAMP() END)
        """
        cursor = self.connection.cursor()
        try:
            cursor.execute(
                sql,
                (
                    self.run_id,
                    self._step_number,
                    table_name,
                    step,
                    status,
                    message,
                    error_message,
                    start_time,
                    status,
                ),
            )
        finally:
            cursor.close()
        return step_number

    def finish_step(
        self,
        step_number: int,
        status: str,
        message: str | None = None,
        error_message: str | None = None,
    ) -> None:
        sql = f"""
            UPDATE {self.qualified_table}
            SET STATUS = %s,
                MESSAGE = %s,
                ERROR_MESSAGE = %s,
                END_TIME = CURRENT_TIMESTAMP()
            WHERE RUN_ID = %s AND RUN_STEP_NO = %s
        """
        cursor = self.connection.cursor()
        try:
            cursor.execute(
                sql,
                (status, message, error_message, self.run_id, step_number),
            )
        finally:
            cursor.close()

    def run_step(
        self,
        step: str,
        action: Callable[[], _T],
        table_name: str | None = None,
    ) -> _T:
        step_number = self.record(step, "RUNNING", table_name, f"{step} started")
        try:
            result = action()
        except KeyboardInterrupt:
            try:
                self.finish_step(
                    step_number,
                    "INTERRUPTED",
                    f"{step} interrupted",
                )
            except Exception:
                logger.exception("Could not store interrupted step %s in Snowflake audit table", step)
            raise
        except Exception as step_error:
            error_message = traceback.format_exc()
            try:
                self.finish_step(
                    step_number,
                    "FAILED",
                    f"{step} failed: {step_error}",
                    error_message,
                )
            except Exception as audit_error:
                logger.exception("Could not store failed step %s in Snowflake audit table", step)
                raise step_error from audit_error
            raise

        self.finish_step(
            step_number,
            "SUCCESS",
            f"{step} completed successfully",
        )
        return result
