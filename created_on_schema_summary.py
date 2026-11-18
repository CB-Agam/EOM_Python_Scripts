"""Export CREATED_ON null/total/max summaries for Snowflake tables in a schema."""

from __future__ import annotations

import argparse
import csv
import os
import sys
from pathlib import Path
from typing import Any

from fourd_to_snowflake import SnowflakeTargetConfig, quote_snowflake_identifier


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Run the CREATED_ON null-count, total-count, and maximum-date summary "
            "for every base table in a Snowflake schema."
        )
    )
    parser.add_argument(
        "--database",
        help="Snowflake database (defaults to SNOWFLAKE_DATABASE)",
    )
    parser.add_argument(
        "--schema",
        help="Snowflake schema (defaults to SNOWFLAKE_SCHEMA)",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("created_on_summary.csv"),
        help="CSV output path (default: ./created_on_summary.csv)",
    )
    return parser


def list_base_tables(cursor: Any, database: str, schema: str) -> list[str]:
    database_name = quote_snowflake_identifier(database)
    cursor.execute(
        f"""
        SELECT TABLE_NAME
        FROM {database_name}.INFORMATION_SCHEMA.TABLES
        WHERE TABLE_SCHEMA = %s
          AND TABLE_TYPE = 'BASE TABLE'
        ORDER BY TABLE_NAME
        """,
        (schema,),
    )
    return [row[0] for row in cursor.fetchall()]


def list_created_on_tables(
    cursor: Any,
    database: str,
    schema: str,
) -> set[str]:
    database_name = quote_snowflake_identifier(database)
    cursor.execute(
        f"""
        SELECT DISTINCT TABLE_NAME
        FROM {database_name}.INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = %s
          AND COLUMN_NAME = 'CREATED_ON'
        """,
        (schema,),
    )
    return {row[0] for row in cursor.fetchall()}


def summarize_table(
    cursor: Any,
    database: str,
    schema: str,
    table: str,
) -> tuple[int, int, str | None]:
    qualified_table = ".".join(
        quote_snowflake_identifier(identifier)
        for identifier in (database, schema, table)
    )
    created_on = quote_snowflake_identifier("CREATED_ON")
    cursor.execute(
        f"""
        SELECT COALESCE(COUNT_IF({created_on} IS NULL), 0),
               COUNT(*),
               MAX({created_on})::VARCHAR
        FROM {qualified_table}
        """
    )
    row = cursor.fetchone()
    if row is None:
        raise RuntimeError(f"Snowflake returned no summary row for {schema}.{table}")
    return int(row[0]), int(row[1]), row[2]


def write_summary(output_path: Path, rows: list[dict[str, object]]) -> None:
    with output_path.open("w", newline="", encoding="utf-8") as output_file:
        writer = csv.DictWriter(
            output_file,
            fieldnames=("TABLE_NAME", "METRIC", "COUNT", "VALUE"),
        )
        writer.writeheader()
        writer.writerows(rows)


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.database:
        os.environ["SNOWFLAKE_DATABASE"] = args.database
    if args.schema:
        os.environ["SNOWFLAKE_SCHEMA"] = args.schema

    config = SnowflakeTargetConfig.from_env()
    database = args.database or config.database
    schema = args.schema or config.schema

    connection = config.connect()
    try:
        from snowflake.connector.errors import Error as SnowflakeError

        cursor = connection.cursor()
        try:
            base_tables = list_base_tables(cursor, database, schema)
            created_on_tables = list_created_on_tables(cursor, database, schema)
            tables = [table for table in base_tables if table in created_on_tables]
            skipped_count = len(base_tables) - len(tables)
            if skipped_count:
                print(
                    f"Skipping {skipped_count} base table(s) without CREATED_ON.",
                    file=sys.stderr,
                )

            result_rows: list[dict[str, object]] = []
            failed_tables = 0
            for table in tables:
                try:
                    null_count, total_count, created_max = summarize_table(
                        cursor,
                        database,
                        schema,
                        table,
                    )
                except SnowflakeError as exc:
                    failed_tables += 1
                    print(
                        f"FAILED {database}.{schema}.{table}: {exc}",
                        file=sys.stderr,
                    )
                    continue

                result_rows.extend(
                    (
                        {
                            "TABLE_NAME": table,
                            "METRIC": "Total_null",
                            "COUNT": null_count,
                            "VALUE": "dt",
                        },
                        {
                            "TABLE_NAME": table,
                            "METRIC": "Total",
                            "COUNT": total_count,
                            "VALUE": "dt",
                        },
                        {
                            "TABLE_NAME": table,
                            "METRIC": "Created_max",
                            "COUNT": 0,
                            "VALUE": created_max,
                        },
                    )
                )
        finally:
            cursor.close()
    finally:
        connection.close()

    write_summary(args.output, result_rows)
    print(
        f"Finished: {len(tables) - failed_tables} table(s) summarized, "
        f"{failed_tables} failed. CSV: {args.output.resolve()}"
    )
    return 1 if failed_tables else 0


if __name__ == "__main__":
    raise SystemExit(main())
