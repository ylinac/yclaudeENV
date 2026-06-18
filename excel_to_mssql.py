"""
Excel 匯入 MSSQL 工具
用法: python excel_to_mssql.py

需要安裝:
    pip install pandas pyodbc openpyxl
"""

import pandas as pd
import pyodbc
import sys
from pathlib import Path

# ── 連線設定 ──────────────────────────────────────────────
DB_SERVER   = "YOUR_SERVER"        # 例: 192.168.1.1 或 localhost\SQLEXPRESS
DB_PORT     = "1433"
DB_NAME     = "YOUR_DATABASE"
DB_USER     = "sa"
DB_PASSWORD = "YOUR_PASSWORD"

# ── 匯入設定 ──────────────────────────────────────────────
EXCEL_FILE  = r"C:\path\to\your_file.xlsx"   # Excel 路徑
SHEET_NAME  = 0                              # 工作表名稱或索引 (0 = 第一張)
TABLE_NAME  = "imported_data"                # 目標資料表名稱
BATCH_SIZE  = 1000                           # 每批 INSERT 筆數
SKIP_ROWS   = 0                              # 跳過前 N 行 (標題除外)

# ── pandas → SQL Server 型別對照 ──────────────────────────
TYPE_MAP = {
    "int64":          "BIGINT",
    "int32":          "INT",
    "float64":        "FLOAT",
    "bool":           "BIT",
    "datetime64[ns]": "DATETIME2",
    "object":         "NVARCHAR(MAX)",
}


def get_connection():
    conn_str = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={DB_SERVER},{DB_PORT};"
        f"DATABASE={DB_NAME};"
        f"UID={DB_USER};"
        f"PWD={DB_PASSWORD};"
        "Encrypt=no;"
    )
    return pyodbc.connect(conn_str)


def pandas_type_to_sql(dtype: str) -> str:
    return TYPE_MAP.get(dtype, "NVARCHAR(MAX)")


def table_exists(cursor, table_name: str) -> bool:
    cursor.execute(
        "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES "
        "WHERE TABLE_NAME = ?",
        table_name,
    )
    return cursor.fetchone()[0] > 0


def create_table(cursor, table_name: str, df: pd.DataFrame):
    cols = ", ".join(
        f"[{col}] {pandas_type_to_sql(str(df[col].dtype))}"
        for col in df.columns
    )
    ddl = f"CREATE TABLE [{table_name}] ({cols})"
    print(f"  建立資料表: {ddl}")
    cursor.execute(ddl)


def insert_batch(cursor, table_name: str, df: pd.DataFrame):
    cols = ", ".join(f"[{c}]" for c in df.columns)
    placeholders = ", ".join("?" * len(df.columns))
    sql = f"INSERT INTO [{table_name}] ({cols}) VALUES ({placeholders})"

    total = 0
    for i in range(0, len(df), BATCH_SIZE):
        batch = df.iloc[i : i + BATCH_SIZE]
        rows = [
            tuple(None if pd.isna(v) else v for v in row)
            for row in batch.itertuples(index=False)
        ]
        cursor.executemany(sql, rows)
        total += len(rows)
        print(f"  已插入 {total}/{len(df)} 筆...", end="\r")
    print()


def main():
    excel_path = Path(EXCEL_FILE)
    if not excel_path.exists():
        print(f"[錯誤] 找不到 Excel 檔案: {excel_path}")
        sys.exit(1)

    # 讀取 Excel
    print(f"讀取 Excel: {excel_path.name}  工作表: {SHEET_NAME}")
    df = pd.read_excel(excel_path, sheet_name=SHEET_NAME, skiprows=SKIP_ROWS)
    df.columns = [str(c).strip() for c in df.columns]   # 清理欄位名稱
    df = df.dropna(how="all")                            # 移除全空白列
    print(f"  共 {len(df)} 筆, {len(df.columns)} 欄")

    # 連線
    print(f"連線 MSSQL: {DB_SERVER} / {DB_NAME}")
    conn = get_connection()
    conn.autocommit = False
    cursor = conn.cursor()

    try:
        if table_exists(cursor, TABLE_NAME):
            print(f"  資料表 [{TABLE_NAME}] 已存在，直接匯入")
        else:
            print(f"  資料表 [{TABLE_NAME}] 不存在，自動建立")
            create_table(cursor, TABLE_NAME, df)

        print(f"插入資料到 [{TABLE_NAME}]...")
        insert_batch(cursor, TABLE_NAME, df)

        conn.commit()
        print(f"完成！共匯入 {len(df)} 筆資料。")

    except Exception as e:
        conn.rollback()
        print(f"[錯誤] {e}")
        raise
    finally:
        cursor.close()
        conn.close()


if __name__ == "__main__":
    main()
