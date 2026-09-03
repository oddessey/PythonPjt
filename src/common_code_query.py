"""공통코드 컬럼의 중복 없는 값을 조회하여 Excel 결과를 생성한다."""

from __future__ import annotations

import importlib
import os
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


REQUIRED_PACKAGES: Dict[str, str] = {
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "psycopg2": "psycopg2-binary",
}

INPUT_EXCEL_PATH: Path = Path(r"D:\ai\work\공통코드조사.xlsx")
OUTPUT_EXCEL_PATH: Path = Path(r"D:\ai\work\공통코드조사_결과.xlsx")
DB_PASSWORD: str = r"lQt}_cB6E&4N"
EXCEL_CELL_MAX_LENGTH: int = 32767
DISTINCT_VALUE_SEPARATOR: str = ", "

SHEET_NAMES: Tuple[str, ...] = ("stastics", "member")

OUTPUT_COLUMNS: Tuple[str, ...] = (
    "데이터베이스",
    "스키마명",
    "테이블물리명",
    "테이블논리명",
    "컬럼물리명",
    "컬럼논리명",
    "공통코드데이터",
)

SOURCE_COLUMNS: Tuple[str, ...] = (
    "데이터베이스",
    "스키마명",
    "테이블물리명",
    "테이블논리명",
    "컬럼물리명",
    "컬럼논리명",
    "데이터타입",
)

DB_CONFIG_BY_SHEET: Dict[str, Dict[str, Any]] = {
    "stastics": {
        "host": "localhost",
        "port": 5532,
        "dbname": "stastics",
        "user": "devreadonly",
        "connect_timeout": 15,
        "sslmode": "require",
        "ssl_min_protocol_version": "TLSv1.2",
    },
    "member": {
        "host": "localhost",
        "port": 5532,
        "dbname": "member",
        "user": "devreadonly",
        "connect_timeout": 15,
        "sslmode": "require",
        "ssl_min_protocol_version": "TLSv1.2",
    },
}


def install_required_packages() -> None:
    """필요한 Python 패키지를 자동으로 설치한다."""
    missing_packages: list[str] = []

    for import_name, package_name in REQUIRED_PACKAGES.items():
        try:
            importlib.import_module(import_name)
        except ImportError:
            missing_packages.append(package_name)

    if not missing_packages:
        return

    print(f"필요한 패키지 설치 중: {', '.join(missing_packages)}")
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", *missing_packages],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE,
        )
    except subprocess.CalledProcessError as error:
        stderr_text: str = error.stderr.decode("utf-8", errors="replace")
        raise RuntimeError(
            f"패키지 설치 실패: {', '.join(missing_packages)}\n{stderr_text}"
        ) from error


install_required_packages()

import pandas as pd
import psycopg2
from psycopg2 import sql
from psycopg2.extensions import connection as PgConnection


@dataclass(frozen=True)
class ColumnQueryKey:
    """스키마/테이블/컬럼 조회 결과 매핑 키."""

    schema_name: str
    table_name: str
    column_name: str


@dataclass(frozen=True)
class SheetJob:
    """시트별 조사 작업 정보."""

    sheet_name: str
    source_dataframe: pd.DataFrame


def connect_postgresql(sheet_name: str) -> PgConnection:
    """시트에 해당하는 PostgreSQL 데이터베이스에 연결한다."""
    if sheet_name not in DB_CONFIG_BY_SHEET:
        raise ValueError(f"지원하지 않는 시트명입니다: {sheet_name}")

    db_config: Dict[str, Any] = dict(DB_CONFIG_BY_SHEET[sheet_name])

    try:
        db_connection: PgConnection = psycopg2.connect(
            **db_config,
            password=DB_PASSWORD,
        )
    except psycopg2.Error as error:
        try:
            db_config.pop("ssl_min_protocol_version", None)
            db_connection = psycopg2.connect(
                **db_config,
                password=DB_PASSWORD,
            )
        except psycopg2.Error as retry_error:
            raise ConnectionError(
                f"PostgreSQL 연결 실패 ({sheet_name}): {retry_error}"
            ) from retry_error

        print(
            "경고: ssl_min_protocol_version 파라미터를 제외하고 재연결했습니다."
        )

    return db_connection


def read_source_sheet(sheet_name: str) -> pd.DataFrame:
    """입력 Excel의 지정 시트 데이터를 읽는다."""
    if not INPUT_EXCEL_PATH.exists():
        raise FileNotFoundError(
            f"입력 파일을 찾을 수 없습니다: {INPUT_EXCEL_PATH}"
        )

    try:
        raw_dataframe: pd.DataFrame = pd.read_excel(
            INPUT_EXCEL_PATH,
            sheet_name=sheet_name,
            header=None,
            skiprows=2,
            engine="openpyxl",
        )
    except ValueError as error:
        raise ValueError(
            f"'{sheet_name}' 시트를 읽을 수 없습니다: {error}"
        ) from error
    except Exception as error:
        raise RuntimeError(
            f"Excel 파일 읽기 실패: {INPUT_EXCEL_PATH}\n{error}"
        ) from error

    if raw_dataframe.empty:
        raise ValueError(f"'{sheet_name}' 시트에 처리할 데이터가 없습니다.")

    if raw_dataframe.shape[1] < len(SOURCE_COLUMNS):
        raise ValueError(
            f"'{sheet_name}' 시트의 컬럼 수가 부족합니다. "
            f"필요: {len(SOURCE_COLUMNS)}, 실제: {raw_dataframe.shape[1]}"
        )

    raw_dataframe = raw_dataframe.iloc[:, : len(SOURCE_COLUMNS)].copy()
    raw_dataframe.columns = list(SOURCE_COLUMNS)
    raw_dataframe = raw_dataframe.dropna(
        subset=["테이블물리명", "컬럼물리명"],
        how="any",
    )
    raw_dataframe = raw_dataframe.reset_index(drop=True)

    if raw_dataframe.empty:
        raise ValueError(
            f"'{sheet_name}' 시트에 유효한 테이블/컬럼 데이터가 없습니다."
        )

    return raw_dataframe


def load_sheet_jobs() -> List[SheetJob]:
    """stastics 시트와 member 시트를 나누어 읽는다."""
    sheet_jobs: List[SheetJob] = []

    for sheet_name in SHEET_NAMES:
        source_dataframe: pd.DataFrame = read_source_sheet(sheet_name)
        sheet_jobs.append(
            SheetJob(
                sheet_name=sheet_name,
                source_dataframe=source_dataframe,
            )
        )
        print(
            f"'{sheet_name}' 시트 입력 건수: {len(source_dataframe)}"
        )

    return sheet_jobs


def build_distinct_query(
    schema_name: str,
    table_name: str,
    column_name: str,
) -> sql.Composed:
    """컬럼의 중복 없는 값을 조회하는 SQL을 생성한다."""
    column_identifier = sql.Identifier(column_name)
    schema_identifier = sql.Identifier(schema_name)
    table_identifier = sql.Identifier(table_name)

    query: sql.Composed = sql.SQL(
        """
        SELECT DISTINCT {column_name}::text AS code_value
        FROM {schema_name}.{table_name}
        WHERE {column_name} IS NOT NULL
        ORDER BY 1
        """
    ).format(
        column_name=column_identifier,
        schema_name=schema_identifier,
        table_name=table_identifier,
    )

    return query


def fetch_distinct_values(
    db_connection: PgConnection,
    schema_name: str,
    table_name: str,
    column_name: str,
) -> str:
    """지정된 스키마/테이블/컬럼의 중복 없는 값을 조회한다."""
    query: sql.Composed = build_distinct_query(
        schema_name=schema_name,
        table_name=table_name,
        column_name=column_name,
    )

    try:
        with db_connection.cursor() as cursor:
            cursor.execute(query)
            rows: List[Tuple[Any, ...]] = cursor.fetchall()
    except psycopg2.Error as error:
        db_connection.rollback()
        raise RuntimeError(
            "쿼리 실행 실패: "
            f"{schema_name}.{table_name}.{column_name}\n{error}"
        ) from error

    distinct_values: List[str] = [
        str(row[0]).strip()
        for row in rows
        if row and row[0] is not None and str(row[0]).strip() != ""
    ]

    if not distinct_values:
        return ""

    combined_text: str = DISTINCT_VALUE_SEPARATOR.join(dict.fromkeys(distinct_values))
    if len(combined_text) <= EXCEL_CELL_MAX_LENGTH:
        return combined_text

    truncated_text: str = combined_text[: EXCEL_CELL_MAX_LENGTH - 20]
    return f"{truncated_text}...(truncated)"


def collect_query_results(
    db_connection: PgConnection,
    source_dataframe: pd.DataFrame,
) -> Dict[ColumnQueryKey, str]:
    """테이블물리명/컬럼물리명 기준 중복 없는 값을 수집한다."""
    query_cache: Dict[ColumnQueryKey, str] = {}
    unique_rows: pd.DataFrame = source_dataframe.drop_duplicates(
        subset=["스키마명", "테이블물리명", "컬럼물리명"]
    )

    total_count: int = len(unique_rows)
    for index, row in enumerate(unique_rows.itertuples(index=False), start=1):
        schema_name: str = str(row.스키마명).strip()
        table_name: str = str(row.테이블물리명).strip()
        column_name: str = str(row.컬럼물리명).strip()
        cache_key: ColumnQueryKey = ColumnQueryKey(
            schema_name=schema_name,
            table_name=table_name,
            column_name=column_name,
        )

        if cache_key in query_cache:
            continue

        print(
            f"[{index}/{total_count}] 조회 중: "
            f"{schema_name}.{table_name}.{column_name}"
        )

        try:
            query_cache[cache_key] = fetch_distinct_values(
                db_connection=db_connection,
                schema_name=schema_name,
                table_name=table_name,
                column_name=column_name,
            )
        except RuntimeError as error:
            print(f"  경고: {error}")
            query_cache[cache_key] = ""

    return query_cache


def build_output_dataframe(
    source_dataframe: pd.DataFrame,
    query_results: Dict[ColumnQueryKey, str],
) -> pd.DataFrame:
    """결과 Excel에 저장할 DataFrame을 생성한다."""
    output_rows: list[Dict[str, Any]] = []

    for row in source_dataframe.itertuples(index=False):
        cache_key: ColumnQueryKey = ColumnQueryKey(
            schema_name=str(row.스키마명).strip(),
            table_name=str(row.테이블물리명).strip(),
            column_name=str(row.컬럼물리명).strip(),
        )
        common_code_data: str = query_results.get(cache_key, "")

        output_rows.append(
            {
                "데이터베이스": row.데이터베이스,
                "스키마명": row.스키마명,
                "테이블물리명": row.테이블물리명,
                "테이블논리명": row.테이블논리명,
                "컬럼물리명": row.컬럼물리명,
                "컬럼논리명": row.컬럼논리명,
                "공통코드데이터": common_code_data,
            }
        )

    return pd.DataFrame(output_rows, columns=list(OUTPUT_COLUMNS))


def normalize_excel_values(output_dataframe: pd.DataFrame) -> pd.DataFrame:
    """Excel이 지원하는 기본 Python 값으로 셀 값을 변환한다."""
    normalized_dataframe: pd.DataFrame = output_dataframe.copy()

    def normalize_value(value: Any) -> Any:
        if value is None or value is pd.NA:
            return None

        if isinstance(value, float) and pd.isna(value):
            return None

        if hasattr(value, "item"):
            value = value.item()

        if value is None:
            return None
        if isinstance(value, pd.Timestamp):
            return value.to_pydatetime()
        if isinstance(value, (str, int, float, bool, datetime)):
            return value
        return str(value)

    values: List[List[Any]] = [
        [normalize_value(value) for value in row]
        for row in normalized_dataframe.itertuples(index=False, name=None)
    ]
    return pd.DataFrame(
        values,
        columns=normalized_dataframe.columns,
        dtype=object,
    )


def write_output_excel(output_by_sheet: Dict[str, pd.DataFrame]) -> Path:
    """조회 결과를 새 Excel 파일로 저장한다."""
    temporary_path: Path = OUTPUT_EXCEL_PATH.with_name(
        f"{OUTPUT_EXCEL_PATH.stem}.tmp.xlsx"
    )

    try:
        with pd.ExcelWriter(temporary_path, engine="openpyxl") as writer:
            for sheet_name, output_dataframe in output_by_sheet.items():
                safe_dataframe: pd.DataFrame = normalize_excel_values(
                    output_dataframe
                )
                safe_dataframe.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    index=False,
                )
        os.replace(temporary_path, OUTPUT_EXCEL_PATH)
    except Exception as error:
        if temporary_path.exists():
            temporary_path.unlink()
        raise RuntimeError(
            f"결과 Excel 저장 실패: {OUTPUT_EXCEL_PATH}\n{error}"
        ) from error

    return OUTPUT_EXCEL_PATH


def process_sheet_job(sheet_job: SheetJob) -> pd.DataFrame:
    """시트 단위로 DB 조회 후 결과 DataFrame을 만든다."""
    db_connection: Optional[PgConnection] = None

    try:
        db_connection = connect_postgresql(sheet_job.sheet_name)
        print(f"PostgreSQL 연결 성공: {sheet_job.sheet_name}")

        query_results: Dict[ColumnQueryKey, str] = collect_query_results(
            db_connection=db_connection,
            source_dataframe=sheet_job.source_dataframe,
        )
        return build_output_dataframe(
            source_dataframe=sheet_job.source_dataframe,
            query_results=query_results,
        )
    finally:
        if db_connection is not None:
            try:
                db_connection.close()
            except psycopg2.Error as error:
                print(f"DB 연결 종료 중 오류: {error}")


def main() -> None:
    """프로그램 실행 진입점."""
    started_at: datetime = datetime.now()
    print("=" * 60)
    print("공통코드 중복 없는 데이터 조회 프로그램")
    print(f"시작 시각: {started_at.strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    try:
        sheet_jobs: List[SheetJob] = load_sheet_jobs()
        output_by_sheet: Dict[str, pd.DataFrame] = {}

        for sheet_job in sheet_jobs:
            print("-" * 60)
            print(f"시트 처리 시작: {sheet_job.sheet_name}")
            output_by_sheet[sheet_job.sheet_name] = process_sheet_job(sheet_job)

        output_path: Path = write_output_excel(output_by_sheet)
        total_rows: int = sum(
            len(dataframe) for dataframe in output_by_sheet.values()
        )

        print("-" * 60)
        print(f"결과 파일 생성 완료: {output_path}")
        print(f"출력 데이터 건수: {total_rows}")

    except (FileNotFoundError, ValueError, ConnectionError, RuntimeError) as error:
        print(f"오류 발생: {error}")
        sys.exit(1)
    except Exception as error:
        print(f"예상치 못한 오류 발생: {error}")
        sys.exit(1)

    finished_at: datetime = datetime.now()
    elapsed_seconds: float = (finished_at - started_at).total_seconds()
    print(f"종료 시각: {finished_at.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"소요 시간: {elapsed_seconds:.2f}초")
    print("=" * 60)


if __name__ == "__main__":
    main()
