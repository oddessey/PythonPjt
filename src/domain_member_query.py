"""PostgreSQL 컬럼 최대 길이 조회 및 Excel 결과 생성 프로그램."""

from __future__ import annotations

import importlib
import os
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional, Tuple


REQUIRED_PACKAGES: Dict[str, str] = {
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "psycopg2": "psycopg2-binary",
}

INPUT_EXCEL_PATH: Path = Path(r"D:\ai\work\도메인중복데이터.xlsx")
OUTPUT_EXCEL_PATH: Path = Path(
    r"D:\ai\work\도메인중복데이터_결과_mnember.xlsx"
)
SHEET_NAME: str = "member"

DB_CONFIG: Dict[str, Any] = {
    "host": "localhost",
    "port": 5532,
    "dbname": "member",
    "user": "devreadonly",
    "connect_timeout": 15,
    "sslmode": "require",
}

EXCEL_COLUMNS: Tuple[str, ...] = (
    "DATABASE",
    "스키마",
    "테이블명",
    "테이블컬럼명",
    "데이터타입",
    "데이터최대",
    "데이터",
)


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
    """테이블/컬럼 조회 결과 매핑 키."""

    table_name: str
    column_name: str


@dataclass(frozen=True)
class ColumnQueryResult:
    """컬럼 최대 길이 조회 결과."""

    data_value: Optional[str]
    data_max: Optional[int]


def connect_postgresql() -> PgConnection:
    """PostgreSQL 데이터베이스에 연결한다."""
    try:
        db_connection: PgConnection = psycopg2.connect(
            **DB_CONFIG,
            password="lQt}_cB6E&4N",
        )
    except psycopg2.Error as error:
        raise ConnectionError(
            f"PostgreSQL 연결 실패: {error}"
        ) from error

    return db_connection


def read_statistics_sheet() -> pd.DataFrame:
    """입력 Excel의 member 시트 데이터를 읽는다."""
    if not INPUT_EXCEL_PATH.exists():
        raise FileNotFoundError(
            f"입력 파일을 찾을 수 없습니다: {INPUT_EXCEL_PATH}"
        )

    try:
        raw_dataframe: pd.DataFrame = pd.read_excel(
            INPUT_EXCEL_PATH,
            sheet_name=SHEET_NAME,
            header=None,
            skiprows=2,
            engine="openpyxl",
        )
    except ValueError as error:
        raise ValueError(
            f"'{SHEET_NAME}' 시트를 읽을 수 없습니다: {error}"
        ) from error
    except Exception as error:
        raise RuntimeError(
            f"Excel 파일 읽기 실패: {INPUT_EXCEL_PATH}\n{error}"
        ) from error

    if raw_dataframe.empty:
        raise ValueError("member 시트에 처리할 데이터가 없습니다.")

    raw_dataframe.columns = [
        "DATABASE",
        "schema",
        "tbNm",
        "tbColNm",
        "data_type",
        "data_max",
        "data_value",
    ]

    return raw_dataframe


def build_max_length_query(
    schema: str,
    table_name: str,
    column_name: str,
) -> sql.Composed:
    """컬럼 최대 길이 조회 SQL을 생성한다."""
    column_identifier = sql.Identifier(column_name)
    schema_identifier = sql.Identifier(schema)
    table_identifier = sql.Identifier(table_name)

    query: sql.Composed = sql.SQL(
        """
        SELECT {column_name}, data_size
        FROM (
            SELECT
                {column_name},
                MAX(LENGTH({column_name}::text)) AS data_size
            FROM {schema_name}.{table_name}
            WHERE {column_name} IS NOT NULL
            GROUP BY {column_name}
        ) AS grouped_data
        ORDER BY data_size DESC
        LIMIT 1
        """
    ).format(
        column_name=column_identifier,
        schema_name=schema_identifier,
        table_name=table_identifier,
    )

    return query


def fetch_column_max_length(
    db_connection: PgConnection,
    schema: str,
    table_name: str,
    column_name: str,
) -> ColumnQueryResult:
    """지정된 스키마/테이블/컬럼의 최대 길이 값을 조회한다."""
    query: sql.Composed = build_max_length_query(schema, table_name, column_name)

    try:
        with db_connection.cursor() as cursor:
            cursor.execute(query)
            row: Optional[Tuple[Any, Any]] = cursor.fetchone()
    except psycopg2.Error as error:
        db_connection.rollback()
        raise RuntimeError(
            "쿼리 실행 실패: "
            f"{schema}.{table_name}.{column_name}\n{error}"
        ) from error

    if row is None:
        return ColumnQueryResult(data_value=None, data_max=None)

    data_value: Optional[str] = None if row[0] is None else str(row[0])
    data_max: Optional[int] = None if row[1] is None else int(row[1])

    return ColumnQueryResult(data_value=data_value, data_max=data_max)


def collect_query_results(
    db_connection: PgConnection,
    source_dataframe: pd.DataFrame,
) -> Dict[ColumnQueryKey, ColumnQueryResult]:
    """테이블명/테이블컬럼명 기준 조회 결과를 수집한다."""
    query_cache: Dict[ColumnQueryKey, ColumnQueryResult] = {}
    unique_rows: pd.DataFrame = source_dataframe.drop_duplicates(
        subset=["schema", "tbNm", "tbColNm"]
    )

    total_count: int = len(unique_rows)
    for index, row in enumerate(unique_rows.itertuples(index=False), start=1):
        schema: str = str(row.schema).strip()
        tb_nm: str = str(row.tbNm).strip()
        tb_col_nm: str = str(row.tbColNm).strip()
        cache_key: ColumnQueryKey = ColumnQueryKey(
            table_name=tb_nm,
            column_name=tb_col_nm,
        )

        if cache_key in query_cache:
            continue

        print(
            f"[{index}/{total_count}] 조회 중: "
            f"{schema}.{tb_nm}.{tb_col_nm}"
        )

        try:
            query_cache[cache_key] = fetch_column_max_length(
                db_connection=db_connection,
                schema=schema,
                table_name=tb_nm,
                column_name=tb_col_nm,
            )
        except RuntimeError as error:
            print(f"  경고: {error}")
            query_cache[cache_key] = ColumnQueryResult(
                data_value=None,
                data_max=None,
            )

    return query_cache


def build_output_dataframe(
    source_dataframe: pd.DataFrame,
    query_results: Dict[ColumnQueryKey, ColumnQueryResult],
) -> pd.DataFrame:
    """결과 Excel에 저장할 DataFrame을 생성한다."""
    output_rows: list[Dict[str, Any]] = []

    for row in source_dataframe.itertuples(index=False):
        cache_key: ColumnQueryKey = ColumnQueryKey(
            table_name=str(row.tbNm).strip(),
            column_name=str(row.tbColNm).strip(),
        )
        query_result: ColumnQueryResult = query_results.get(
            cache_key,
            ColumnQueryResult(data_value=None, data_max=None),
        )

        output_rows.append(
            {
                "DATABASE": row.DATABASE,
                "스키마": row.schema,
                "테이블명": row.tbNm,
                "테이블컬럼명": row.tbColNm,
                "데이터타입": row.data_type,
                "데이터최대": query_result.data_max,
                "데이터": query_result.data_value,
            }
        )

    return pd.DataFrame(output_rows, columns=list(EXCEL_COLUMNS))


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

    values = [
        [normalize_value(value) for value in row]
        for row in normalized_dataframe.itertuples(index=False, name=None)
    ]
    return pd.DataFrame(values, columns=normalized_dataframe.columns, dtype=object)


def write_output_excel(output_dataframe: pd.DataFrame) -> Path:
    """조회 결과를 새 Excel 파일로 저장한다."""
    safe_dataframe: pd.DataFrame = normalize_excel_values(output_dataframe)
    temporary_path: Path = OUTPUT_EXCEL_PATH.with_name(
        f"{OUTPUT_EXCEL_PATH.stem}.tmp.xlsx"
    )

    try:
        safe_dataframe.to_excel(
            temporary_path,
            sheet_name=SHEET_NAME,
            index=False,
            engine="openpyxl",
        )
        os.replace(temporary_path, OUTPUT_EXCEL_PATH)
    except Exception as error:
        if temporary_path.exists():
            temporary_path.unlink()
        raise RuntimeError(
            f"결과 Excel 저장 실패: {OUTPUT_EXCEL_PATH}\n{error}"
        ) from error

    return OUTPUT_EXCEL_PATH


def main() -> None:
    """프로그램 실행 진입점."""
    started_at: datetime = datetime.now()
    print("=" * 60)
    print("도메인 중복 데이터 통계 조회 프로그램")
    print(f"시작 시각: {started_at.strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    db_connection: Optional[PgConnection] = None

    try:
        source_dataframe: pd.DataFrame = read_statistics_sheet()
        print(f"입력 데이터 건수: {len(source_dataframe)}")

        db_connection = connect_postgresql()
        print("PostgreSQL 연결 성공")

        query_results: Dict[ColumnQueryKey, ColumnQueryResult] = (
            collect_query_results(
                db_connection=db_connection,
                source_dataframe=source_dataframe,
            )
        )

        output_dataframe: pd.DataFrame = build_output_dataframe(
            source_dataframe=source_dataframe,
            query_results=query_results,
        )
        output_path: Path = write_output_excel(output_dataframe)

        print("-" * 60)
        print(f"결과 파일 생성 완료: {output_path}")
        print(f"출력 데이터 건수: {len(output_dataframe)}")
        print(f"조회 완료 건수: {len(query_results)}")

    except (FileNotFoundError, ValueError, ConnectionError, RuntimeError) as error:
        print(f"오류 발생: {error}")
        sys.exit(1)
    except Exception as error:
        print(f"예상치 못한 오류 발생: {error}")
        sys.exit(1)
    finally:
        if db_connection is not None:
            try:
                db_connection.close()
            except psycopg2.Error as error:
                print(f"DB 연결 종료 중 오류: {error}")

    finished_at: datetime = datetime.now()
    elapsed_seconds: float = (finished_at - started_at).total_seconds()
    print(f"종료 시각: {finished_at.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"소요 시간: {elapsed_seconds:.2f}초")
    print("=" * 60)


if __name__ == "__main__":
    main()
