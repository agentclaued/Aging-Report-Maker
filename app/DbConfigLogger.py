import json
from dataclasses import dataclass, field
from datetime import datetime
from functools import lru_cache
from pathlib import Path
from contextlib import suppress
from typing import Iterable

import pandas as pd
import pyodbc

STEP_01_RAW_ROWS_TABLE = "step_01_raw_rows"
STEP_01_RAW_ROWS_TABLE_LEGACY = "raw_rows"
STEP_01_INQUIRY_TABLE = "step_01_inquiry_rows"
STEP_01_INQUIRY_TABLE_LEGACY_CURRENT = "step_01_inquiry_rows"   #inquiry_rows  EDIT STEP 01
STEP_01_INQUIRY_TABLE_LEGACY_OLD = "step_01_inquiry"
STEP_01A_CR_TABLE = "step_01a_cr_sheet"
STEP_01A_CR_TABLE_LEGACY = "rpt_02a_cr_sheet"
STEP_01A_IN_TABLE = "step_01a_in_sheet"
STEP_01A_IN_TABLE_LEGACY = "rpt_02b_in_sheet"
STEP_01A_DB_TABLE = "step_01a_db_sheet"
STEP_01A_DB_TABLE_LEGACY = "rpt_02c_db_sheet"
STEP_01_NORMALIZED_FILE_KIND = "inquiry_normalized_file"
STEP_01_COMPLETED_STATUS = "inquiry_upload_completed"


@dataclass(frozen=True)
class Settings:
    app_title: str = "WalMartSummary SQL FastAPI"
    sql_server: str = r"localhost\SQLEXPRESS"
    sql_database: str = "WallMartSummary"
    sql_driver: str = "ODBC Driver 17 for SQL Server"
    trusted_connection: str = "yes"
    default_company_code: str = "wallmartsummary"

    @property
    def base_dir(self) -> Path:
        return Path(__file__).resolve().parent.parent

    @property
    def output_dir(self) -> Path:
        return self.base_dir / "output"

    @property
    def templates_dir(self) -> Path:
        return self.base_dir / "templates"


def _parse_env_file(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    if not path.exists():
        return values

    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue

        key, value = line.split("=", 1)
        values[key.strip().upper()] = value.strip()

    return values


@lru_cache(maxsize=1)
def get_settings() -> Settings:
    base_dir = Path(__file__).resolve().parent.parent
    env_values = _parse_env_file(base_dir / "app_runtime.env")
    return Settings(
        app_title=env_values.get("APP_TITLE", Settings.app_title),
        sql_server=env_values.get("SQL_SERVER", Settings.sql_server),
        sql_database=env_values.get("SQL_DATABASE", Settings.sql_database),
        sql_driver=env_values.get("SQL_DRIVER", Settings.sql_driver),
        trusted_connection=env_values.get("TRUSTED_CONNECTION", Settings.trusted_connection),
        default_company_code=env_values.get("DEFAULT_COMPANY_CODE", Settings.default_company_code),
    )


@dataclass
class Logger:
    entries: list[dict[str, str]] = field(default_factory=list)
    progress_token: str = ""

    def log(self, message: str, level: str = "INFO") -> None:
        self.entries.append(
            {
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "level": level,
                "message": message,
            }
        )
        if self.progress_token:
            from app import progress

            progress.bump(self.progress_token, message)


class SQLServerRepository:
    def __init__(self, settings: Settings, logger: Logger) -> None:
        self.settings = settings
        self.logger = logger

    def ensure_database(self) -> None:
        self.logger.log(f"Ensuring database exists: {self.settings.sql_database}")

        with self._connect("master", autocommit=True) as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                IF DB_ID(?) IS NULL
                BEGIN
                    DECLARE @sql NVARCHAR(MAX) = N'CREATE DATABASE [' + REPLACE(?, ']', ']]') + N']';
                    EXEC(@sql);
                END
                """,
                self.settings.sql_database,
                self.settings.sql_database,
            )

        with self._connect() as conn:
            cursor = conn.cursor()
            for statement in self._schema_statements():
                cursor.execute(statement)
            conn.commit()

        self.logger.log("Database ready", "SUCCESS")

    def create_batch(
        self,
        raw_filename: str,
        combine_file_count: int,
        status: str,
        notes: str = "",
    ) -> int:
        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                INSERT INTO dbo.upload_batches (raw_filename, combine_file_count, status, notes)
                OUTPUT INSERTED.id
                VALUES (?, ?, ?, ?)
                """,
                raw_filename,
                combine_file_count,
                status,
                notes,
            )
            batch_id = int(cursor.fetchone()[0])
            conn.commit()

        self.logger.log(f"Created batch id {batch_id}")
        return batch_id

    def update_batch_status(self, batch_id: int, status: str, notes: str = "") -> None:
        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "UPDATE dbo.upload_batches SET status = ?, notes = ? WHERE id = ?",
                status,
                notes,
                batch_id,
            )
            conn.commit()

    def store_raw_rows(self, batch_id: int, raw_df: pd.DataFrame, ref_col: str | None = None) -> None:
        rows = [
            (
                batch_id,
                row_number,
                record.get(ref_col) if ref_col else None,
                json.dumps(record, ensure_ascii=True, default=str),
            )
            for row_number, record in enumerate(self._records(raw_df), start=1)
        ]
        self._executemany(
            f"INSERT INTO dbo.{STEP_01_RAW_ROWS_TABLE} (batch_id, row_number, ref_value, payload_json) VALUES (?, ?, ?, ?)",
            rows,
        )
        self.logger.log(f"Uploaded {len(rows)} rows to SQL table {STEP_01_RAW_ROWS_TABLE}", "SUCCESS")


    def store_generated_file(self, batch_id: int, kind: str, file_name: str, file_path: str) -> None:
        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO dbo.generated_files (batch_id, kind, file_name, file_path) VALUES (?, ?, ?, ?)",
                batch_id,
                kind,
                file_name,
                file_path,
            )
            conn.commit()
        self.logger.log(f"Uploaded generated file record to SQL table generated_files: {file_name}", "SUCCESS")



    def replace_dynamic_table_rows(self, table_name: str, data_frame: pd.DataFrame, batch_id: int = 0) -> None:
        columns = [str(column) for column in data_frame.columns]
        self._ensure_dynamic_table(table_name, columns)

        with self._connect() as conn:
            cursor = conn.cursor()
            try:
                cursor.execute(f"TRUNCATE TABLE dbo.{self._quote_identifier(table_name)}")
                self.logger.log(f"SQL table {table_name} truncated before refresh")
            except Exception:
                conn.rollback()
                cursor.execute(f"DELETE FROM dbo.{self._quote_identifier(table_name)}")
                self.logger.log(f"SQL table {table_name} cleared with DELETE before refresh", "WARNING")
            conn.commit()

        row_count = self._store_dynamic_table_rows(table_name, batch_id, data_frame)
        self.logger.log(f"Replaced SQL table {table_name} with {row_count} refreshed rows", "SUCCESS")

    def log_prepared_columns(self, table_name: str, columns: list[str], source_step: str) -> None:
        if not columns:
            self.logger.log(f"{source_step}: no early prepared columns for {table_name}")
            return
        self.logger.log(
            f"{source_step}: early prepared columns for {table_name}: {', '.join(columns)}",
            "SUCCESS",
        )

    def ensure_step_table(self, table_name: str, data_frame: pd.DataFrame) -> None:
        columns = [str(column) for column in data_frame.columns]
        table_exists = self._table_exists(table_name)
        if not table_exists:
            self.logger.log(f"SQL table {table_name} does not exist. Creating table.", "INFO")
        self._ensure_dynamic_table(table_name, columns)
        self._expand_dynamic_text_columns(table_name, columns)
        if not table_exists:
            self.logger.log(f"SQL table {table_name} created.", "SUCCESS")


    def execute_sql_script(self, sql_script: str) -> None:
        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute(sql_script)
            conn.commit()
        self.logger.log("SQL script executed successfully", "SUCCESS")

    def execute_sql_script_with_messages(self, sql_script: str) -> list[str]:
        messages: list[str] = []
        seen_messages: set[str] = set()

        def collect_messages(source: object) -> None:
            raw_messages = getattr(source, "messages", None)
            if not raw_messages:
                return

            for raw_message in raw_messages:
                message_text = ""
                if isinstance(raw_message, tuple) and len(raw_message) >= 2:
                    message_text = str(raw_message[1]).strip()
                else:
                    message_text = str(raw_message).strip()

                if not message_text or message_text in seen_messages:
                    continue

                seen_messages.add(message_text)
                messages.append(message_text)

        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute(sql_script)
            collect_messages(conn)
            collect_messages(cursor)

            while cursor.nextset():
                collect_messages(conn)
                collect_messages(cursor)

            conn.commit()
            collect_messages(conn)
            collect_messages(cursor)

        self.logger.log("SQL script executed successfully", "SUCCESS")
        return messages

    def apply_upload_data_types(self, sql_script: str, description: str) -> None:
        self.logger.log(f"Applying SQL data types after upload: {description}")
        sql_messages = self.execute_sql_script_with_messages(sql_script)
        for message_text in sql_messages:
            normalized_message = str(message_text).strip()
            if "Altered column:" in normalized_message or "Skipped missing column:" in normalized_message:
                self.logger.log(f"Upload type SQL message: {normalized_message}", "NOTICE")
            else:
                self.logger.log(f"Upload type SQL message: {normalized_message}")
        self.logger.log(f"SQL data types applied after upload: {description}", "SUCCESS")

    def read_table(self, table_name: str) -> pd.DataFrame:
        if not self._table_exists(table_name):
            return pd.DataFrame()

        with self._connect() as conn:
            return pd.read_sql(f"SELECT * FROM dbo.{self._quote_identifier(table_name)}", conn)

    def list_user_tables(self) -> list[str]:
        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                SELECT t.name
                FROM sys.tables AS t
                INNER JOIN sys.schemas AS s ON s.schema_id = t.schema_id
                WHERE s.name = 'dbo'
                ORDER BY t.name
                """
            )
            return [str(row[0]) for row in cursor.fetchall()]

    def get_latest_batch_id(self) -> int | None:
        if not self._table_exists("upload_batches"):
            return None

        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT MAX(id) FROM dbo.upload_batches")
            row = cursor.fetchone()

        if not row or row[0] is None:
            return None

        return int(row[0])


    def list_uploaded_batch_filenames(self) -> set[str]:
        if not self._table_exists("upload_batches"):
            return set()
        if not self._column_exists("upload_batches", "raw_filename"):
            return set()

        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                SELECT DISTINCT raw_filename
                FROM dbo.upload_batches
                WHERE raw_filename IS NOT NULL
                  AND LTRIM(RTRIM(raw_filename)) <> ''
                """
            )
            return {str(row[0]).strip() for row in cursor.fetchall() if row[0] is not None}


    def healthcheck(self) -> dict[str, str]:
        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT DB_NAME()")
            db_name = cursor.fetchone()[0]
        return {"server": self.settings.sql_server, "database": db_name, "status": "ok"}

    def table_has_rows(self, table_name: str) -> bool:
        if not self._table_exists(table_name):
            return False

        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                SELECT CASE
                    WHEN EXISTS (SELECT TOP 1 1 FROM dbo.""" + self._quote_identifier(table_name) + """) THEN 1
                    ELSE 0
                END
                """
            )
            return bool(cursor.fetchone()[0])




    def _connect(self, database: str | None = None, autocommit: bool = False) -> pyodbc.Connection:
        target_db = database or self.settings.sql_database
        connection_string = (
            f"DRIVER={{{self.settings.sql_driver}}};"
            f"SERVER={self.settings.sql_server};"
            f"DATABASE={target_db};"
            f"Trusted_Connection={self.settings.trusted_connection};"
        )
        return pyodbc.connect(connection_string, autocommit=autocommit)

    def _executemany(self, sql: str, rows: Iterable[tuple]) -> None:
        row_list = list(rows)
        if not row_list:
            return
        conn = self._connect()
        try:
            cursor = conn.cursor()
            cursor.fast_executemany = True
            cursor.executemany(sql, row_list)
            conn.commit()
        except Exception:
            # Some SQL Server failures abort the transaction server-side.
            # A follow-up rollback can then raise 3971/3926 and hide the real error.
            with suppress(Exception):
                conn.rollback()
            raise
        finally:
            with suppress(Exception):
                conn.close()

    def _schema_statements(self) -> list[str]:
        return [
            f"""
            IF OBJECT_ID('dbo.{STEP_01_RAW_ROWS_TABLE_LEGACY}', 'U') IS NOT NULL
                AND OBJECT_ID('dbo.{STEP_01_RAW_ROWS_TABLE}', 'U') IS NULL
            BEGIN
                EXEC sp_rename 'dbo.{STEP_01_RAW_ROWS_TABLE_LEGACY}', '{STEP_01_RAW_ROWS_TABLE}';
            END
            """,
            f"""
            IF OBJECT_ID('dbo.{STEP_01_INQUIRY_TABLE_LEGACY_CURRENT}', 'U') IS NOT NULL
                AND OBJECT_ID('dbo.{STEP_01_INQUIRY_TABLE}', 'U') IS NULL
            BEGIN
                EXEC sp_rename 'dbo.{STEP_01_INQUIRY_TABLE_LEGACY_CURRENT}', '{STEP_01_INQUIRY_TABLE}';
            END
            """,
            f"""
            IF OBJECT_ID('dbo.{STEP_01_INQUIRY_TABLE_LEGACY_OLD}', 'U') IS NOT NULL
                AND OBJECT_ID('dbo.{STEP_01_INQUIRY_TABLE}', 'U') IS NULL
            BEGIN
                EXEC sp_rename 'dbo.{STEP_01_INQUIRY_TABLE_LEGACY_OLD}', '{STEP_01_INQUIRY_TABLE}';
            END
            """,
            f"""
            IF OBJECT_ID('dbo.{STEP_01A_CR_TABLE_LEGACY}', 'U') IS NOT NULL
                AND OBJECT_ID('dbo.{STEP_01A_CR_TABLE}', 'U') IS NULL
            BEGIN
                EXEC sp_rename 'dbo.{STEP_01A_CR_TABLE_LEGACY}', '{STEP_01A_CR_TABLE}';
            END
            """,
            f"""
            IF OBJECT_ID('dbo.{STEP_01A_IN_TABLE_LEGACY}', 'U') IS NOT NULL
                AND OBJECT_ID('dbo.{STEP_01A_IN_TABLE}', 'U') IS NULL
            BEGIN
                EXEC sp_rename 'dbo.{STEP_01A_IN_TABLE_LEGACY}', '{STEP_01A_IN_TABLE}';
            END
            """,
            f"""
            IF OBJECT_ID('dbo.{STEP_01A_DB_TABLE_LEGACY}', 'U') IS NOT NULL
                AND OBJECT_ID('dbo.{STEP_01A_DB_TABLE}', 'U') IS NULL
            BEGIN
                EXEC sp_rename 'dbo.{STEP_01A_DB_TABLE_LEGACY}', '{STEP_01A_DB_TABLE}';
            END
            """,
            """
            IF OBJECT_ID('dbo.upload_batches', 'U') IS NOT NULL
                AND NOT EXISTS (
                    SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
                    WHERE TABLE_NAME = 'upload_batches' AND COLUMN_NAME = 'raw_filename'
                )
            DROP TABLE dbo.upload_batches
            """,
            """
            IF OBJECT_ID('dbo.upload_batches', 'U') IS NOT NULL
                AND EXISTS (
                    SELECT 1
                    FROM INFORMATION_SCHEMA.COLUMNS
                    WHERE TABLE_NAME = 'upload_batches'
                      AND COLUMN_NAME = 'raw_filename'
                      AND CHARACTER_MAXIMUM_LENGTH <> -1
                )
            ALTER TABLE dbo.upload_batches ALTER COLUMN raw_filename NVARCHAR(MAX) NOT NULL
            """,
            """
            IF OBJECT_ID('dbo.upload_batches', 'U') IS NULL
            CREATE TABLE dbo.upload_batches (
                id                  INT IDENTITY(1,1) PRIMARY KEY,
                created_at          DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME(),
                raw_filename        NVARCHAR(MAX) NOT NULL,
                combine_file_count  INT NOT NULL,
                status              NVARCHAR(50) NOT NULL,
                notes               NVARCHAR(MAX) NULL
            )
            """,
            f"""
            IF OBJECT_ID('dbo.{STEP_01_RAW_ROWS_TABLE}', 'U') IS NULL
            CREATE TABLE dbo.{STEP_01_RAW_ROWS_TABLE} (
                id              INT IDENTITY(1,1) PRIMARY KEY,
                batch_id        INT NOT NULL,
                row_number      INT NOT NULL,
                ref_value       NVARCHAR(255) NULL,
                payload_json    NVARCHAR(MAX) NOT NULL,
                created_at      DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME()
            )
            """,
            """
            IF OBJECT_ID('dbo.combine_rows', 'U') IS NULL
            CREATE TABLE dbo.combine_rows (
                id              INT IDENTITY(1,1) PRIMARY KEY,
                batch_id        INT NOT NULL,
                source_kind     NVARCHAR(100) NOT NULL DEFAULT 'legacy',
                source_file     NVARCHAR(255) NOT NULL,
                row_number      INT NOT NULL,
                payload_json    NVARCHAR(MAX) NOT NULL,
                created_at      DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME()
            )
            """,
            """
            IF OBJECT_ID('dbo.combine_rows', 'U') IS NOT NULL
                AND NOT EXISTS (
                    SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
                    WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'combine_rows' AND COLUMN_NAME = 'source_kind'
                )
            ALTER TABLE dbo.combine_rows
            ADD source_kind NVARCHAR(100) NOT NULL CONSTRAINT DF_combine_rows_source_kind DEFAULT 'legacy'
            """,
            """
            IF OBJECT_ID('dbo.generated_files', 'U') IS NULL
            CREATE TABLE dbo.generated_files (
                id              INT IDENTITY(1,1) PRIMARY KEY,
                batch_id        INT NOT NULL,
                kind            NVARCHAR(50) NOT NULL,
                file_name       NVARCHAR(255) NOT NULL,
                file_path       NVARCHAR(1000) NOT NULL,
                created_at      DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME()
            )
            """,
        ]

    def _store_dynamic_table_rows(self, table_name: str, batch_id: int, data_frame: pd.DataFrame) -> int:
        columns = [str(column) for column in data_frame.columns]
        self._ensure_dynamic_table(table_name, columns)
        self._expand_dynamic_text_columns(table_name, columns)

        column_sql = ", ".join(self._quote_identifier(column) for column in columns)
        insert_sql = (
            f"INSERT INTO dbo.{self._quote_identifier(table_name)} "
            f"([batch_id], [row_number], {column_sql}) "
            f"VALUES (?, ?, {', '.join('?' for _ in columns)})"
        )

        rows = []
        for row_number, record in enumerate(self._records(data_frame), start=1):
            row_values = [None if record.get(column, "") == "" else str(record.get(column, "")) for column in columns]
            rows.append((batch_id, row_number, *row_values))

        try:
            self._executemany(insert_sql, rows)
        except Exception as exc:
            error_text = str(exc)
            if "22001" not in error_text and "right truncation" not in error_text.lower():
                raise

            self.logger.log(
                f"Detected SQL truncation while uploading {table_name}; rebuilding the dynamic table and retrying once.",
                "WARNING",
            )
            self._rebuild_dynamic_table(table_name, columns)
            self._executemany(insert_sql, rows)
        return len(rows)

    def _ensure_dynamic_table(self, table_name: str, columns: list[str]) -> None:
        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                IF OBJECT_ID(?, 'U') IS NULL
                BEGIN
                    EXEC(
                        'CREATE TABLE ' + ? + ' (
                            [id] INT IDENTITY(1,1) PRIMARY KEY,
                            [batch_id] INT NOT NULL,
                            [row_number] INT NOT NULL
                        )'
                    )
                END
                """,
                f"dbo.{table_name}",
                f"dbo.{self._quote_identifier(table_name)}",
            )

            existing_columns = {
                row[0].lower()
                for row in cursor.execute(
                    """
                    SELECT COLUMN_NAME
                    FROM INFORMATION_SCHEMA.COLUMNS
                    WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = ?
                    """,
                    table_name,
                ).fetchall()
            }

            for column in columns:
                if column.lower() in existing_columns:
                    continue
                cursor.execute(
                    f"ALTER TABLE dbo.{self._quote_identifier(table_name)} "
                    f"ADD {self._quote_identifier(column)} NVARCHAR(MAX) NULL"
                )

            conn.commit()

    def _expand_dynamic_text_columns(
        self,
        table_name: str,
        columns: list[str],
        include_legacy_types: bool = False,
    ) -> None:
        if not columns:
            return

        protected_columns = {"id", "batch_id", "row_number"}
        target_columns = [column for column in columns if column.lower() not in protected_columns]
        if not target_columns:
            return

        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH
                FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = ?
                """,
                table_name,
            )
            column_info = {
                str(row[0]).lower(): (str(row[1]).lower(), row[2])
                for row in cursor.fetchall()
            }

            altered_columns: list[str] = []
            for column in target_columns:
                info = column_info.get(column.lower())
                if info is None:
                    continue

                data_type, max_length = info
                allowed_types = {"varchar", "nvarchar"}
                if include_legacy_types:
                    allowed_types |= {"char", "nchar", "text", "ntext"}
                if data_type not in allowed_types:
                    continue
                if max_length == -1:
                    continue

                cursor.execute(
                    f"ALTER TABLE dbo.{self._quote_identifier(table_name)} "
                    f"ALTER COLUMN {self._quote_identifier(column)} NVARCHAR(MAX) NULL"
                )
                altered_columns.append(column)

            conn.commit()

        if altered_columns:
            self.logger.log(
                f"Expanded SQL text columns to NVARCHAR(MAX) in {table_name}: {', '.join(altered_columns)}",
                "NOTICE",
            )

    def _rebuild_dynamic_table(self, table_name: str, columns: list[str]) -> None:
        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"DROP TABLE IF EXISTS dbo.{self._quote_identifier(table_name)}"
            )
            cursor.execute(
                f"""
                CREATE TABLE dbo.{self._quote_identifier(table_name)} (
                    [id] INT IDENTITY(1,1) PRIMARY KEY,
                    [batch_id] INT NOT NULL,
                    [row_number] INT NOT NULL
                )
                """
            )
            for column in columns:
                cursor.execute(
                    f"ALTER TABLE dbo.{self._quote_identifier(table_name)} "
                    f"ADD {self._quote_identifier(column)} NVARCHAR(MAX) NULL"
                )
            conn.commit()

        self.logger.log(
            f"Rebuilt SQL dynamic table {table_name} from the current upload columns.",
            "NOTICE",
        )

    def _table_exists(self, table_name: str) -> bool:
        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                SELECT CASE WHEN OBJECT_ID(?, 'U') IS NULL THEN 0 ELSE 1 END
                """,
                f"dbo.{table_name}",
            )
            return bool(cursor.fetchone()[0])

    def _column_exists(self, table_name: str, column_name: str) -> bool:
        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                SELECT CASE WHEN EXISTS (
                    SELECT 1
                    FROM INFORMATION_SCHEMA.COLUMNS
                    WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = ? AND COLUMN_NAME = ?
                ) THEN 1 ELSE 0 END
                """,
                table_name,
                column_name,
            )
            return bool(cursor.fetchone()[0])

    @staticmethod
    def _quote_identifier(name: str) -> str:
        return f"[{name.replace(']', ']]')}]"

    @staticmethod
    def _records(df: pd.DataFrame) -> list[dict]:
        return df.fillna("").to_dict(orient="records")

