"""
進階英語免修集體申請處理工具
用法: python excel_to_mssql.py

需要安裝:
    pip install pandas pyodbc openpyxl
"""

import pandas as pd
import pyodbc
import sys
from pathlib import Path

# ── 學籍系統連線（MSSQL 65）────────────────────────────────
DB_REGISTRY = {
    "server":   "192.168.1.65",
    "port":     "1433",
    "database": "YOUR_REGISTRY_DB",
    "user":     "sa",
    "password": "YOUR_PASSWORD",
}

# ── 免修系統連線（MSSQL 211）───────────────────────────────
DB_EXEMPTION = {
    "server":   "192.168.1.211",
    "port":     "1433",
    "database": "YOUR_EXEMPTION_DB",
    "user":     "sa",
    "password": "YOUR_PASSWORD",
}

# ── 匯入設定 ──────────────────────────────────────────────
EXCEL_FILE  = r"1142進階英語免修集體申請通過名單.xlsx"
SHEET_NAME  = 0
SKIP_ROWS   = 1
BATCH_SIZE  = 1000

# ── Excel 欄位名稱 ────────────────────────────────────────
COL_STUDENT_ID = "學號"
COL_NAME       = "姓名"

# ── 學籍資料表設定（MSSQL 65）────────────────────────────
REGISTRY_TABLE    = "學籍"          # 學籍資料表名稱
REGISTRY_ID_COL   = "學號"          # 學號欄位
REGISTRY_NAME_COL = "姓名"          # 姓名欄位

# ── 免修資料表設定（MSSQL 211）───────────────────────────
EXEMPTION_TABLE  = "couexe"         # 免修資料表名稱
EXEMPTION_ID_COL = "學號"           # 學號欄位

# ── pandas → SQL Server 型別對照 ──────────────────────────
TYPE_MAP = {
    "int64":          "BIGINT",
    "int32":          "INT",
    "float64":        "FLOAT",
    "bool":           "BIT",
    "datetime64[ns]": "DATETIME2",
    "object":         "NVARCHAR(MAX)",
}


# ────────────────────────────────────────────────────────
# 連線
# ────────────────────────────────────────────────────────

def _connect(cfg: dict) -> pyodbc.Connection:
    conn_str = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={cfg['server']},{cfg['port']};"
        f"DATABASE={cfg['database']};"
        f"UID={cfg['user']};"
        f"PWD={cfg['password']};"
        "Encrypt=no;"
    )
    return pyodbc.connect(conn_str)


# ────────────────────────────────────────────────────────
# 學籍核對（MSSQL 65）
# ────────────────────────────────────────────────────────

def verify_students(df: pd.DataFrame) -> tuple[pd.DataFrame, list[dict]]:
    """
    逐筆對照學籍系統。
    回傳 (通過的 DataFrame, 失敗紀錄列表)
    失敗原因: 'not_found'（學號不存在）或 'name_mismatch'（姓名不符）
    """
    print(f"\n[1/3] 學籍核對（{DB_REGISTRY['server']}）...")
    conn = _connect(DB_REGISTRY)
    cursor = conn.cursor()

    passed_rows = []
    failed = []

    for _, row in df.iterrows():
        sid  = str(row[COL_STUDENT_ID]).strip()
        name = str(row[COL_NAME]).strip()

        cursor.execute(
            f"SELECT [{REGISTRY_NAME_COL}] FROM [{REGISTRY_TABLE}] WHERE [{REGISTRY_ID_COL}] = ?",
            sid,
        )
        result = cursor.fetchone()

        if result is None:
            failed.append({"學號": sid, "姓名": name, "原因": "學號不存在"})
        elif result[0].strip() != name:
            failed.append({
                "學號": sid,
                "姓名(申請)": name,
                "姓名(學籍)": result[0].strip(),
                "原因": "姓名不符",
            })
        else:
            passed_rows.append(row)

    cursor.close()
    conn.close()

    passed_df = pd.DataFrame(passed_rows, columns=df.columns)
    print(f"  通過: {len(passed_df)} 筆　失敗: {len(failed)} 筆")
    return passed_df, failed


# ────────────────────────────────────────────────────────
# 免修紀錄查詢 & 寫入（MSSQL 211）
# ────────────────────────────────────────────────────────

def filter_existing_exemptions(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """
    排除 couexe 中已有免修紀錄的學生。
    回傳 (需寫入的 DataFrame, 已存在的學號列表)
    """
    print(f"\n[2/3] 查詢免修紀錄（{DB_EXEMPTION['server']} / {EXEMPTION_TABLE}）...")
    conn = _connect(DB_EXEMPTION)
    cursor = conn.cursor()

    to_insert = []
    already_exists = []

    for _, row in df.iterrows():
        sid = str(row[COL_STUDENT_ID]).strip()
        cursor.execute(
            f"SELECT COUNT(*) FROM [{EXEMPTION_TABLE}] WHERE [{EXEMPTION_ID_COL}] = ?",
            sid,
        )
        count = cursor.fetchone()[0]
        if count > 0:
            already_exists.append(sid)
        else:
            to_insert.append(row)

    cursor.close()
    conn.close()

    insert_df = pd.DataFrame(to_insert, columns=df.columns)
    print(f"  需寫入: {len(insert_df)} 筆　已存在(略過): {len(already_exists)} 筆")
    return insert_df, already_exists


def _pandas_type_to_sql(dtype: str) -> str:
    return TYPE_MAP.get(dtype, "NVARCHAR(MAX)")


def _table_exists(cursor, table_name: str) -> bool:
    cursor.execute(
        "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = ?",
        table_name,
    )
    return cursor.fetchone()[0] > 0


def _create_table(cursor, table_name: str, df: pd.DataFrame):
    cols = ", ".join(
        f"[{col}] {_pandas_type_to_sql(str(df[col].dtype))}"
        for col in df.columns
    )
    cursor.execute(f"CREATE TABLE [{table_name}] ({cols})")
    print(f"  已建立資料表 [{table_name}]")


def insert_exemptions(df: pd.DataFrame):
    """將通過核對且尚無紀錄的學生寫入 couexe。"""
    if df.empty:
        print("\n[3/3] 無需寫入，略過。")
        return

    print(f"\n[3/3] 寫入免修紀錄（{EXEMPTION_TABLE}）...")
    conn = _connect(DB_EXEMPTION)
    conn.autocommit = False
    cursor = conn.cursor()

    try:
        if not _table_exists(cursor, EXEMPTION_TABLE):
            _create_table(cursor, EXEMPTION_TABLE, df)

        cols         = ", ".join(f"[{c}]" for c in df.columns)
        placeholders = ", ".join("?" * len(df.columns))
        sql          = f"INSERT INTO [{EXEMPTION_TABLE}] ({cols}) VALUES ({placeholders})"

        total = 0
        for i in range(0, len(df), BATCH_SIZE):
            batch = df.iloc[i : i + BATCH_SIZE]
            rows  = [
                tuple(None if pd.isna(v) else v for v in row)
                for row in batch.itertuples(index=False)
            ]
            cursor.executemany(sql, rows)
            total += len(rows)
            print(f"  已插入 {total}/{len(df)} 筆...", end="\r")

        conn.commit()
        print(f"\n  完成，共寫入 {total} 筆。")

    except Exception as e:
        conn.rollback()
        print(f"\n[錯誤] 寫入失敗，已 rollback: {e}")
        raise
    finally:
        cursor.close()
        conn.close()


# ────────────────────────────────────────────────────────
# 報告
# ────────────────────────────────────────────────────────

def print_report(total: int, inserted: int, already_exists: list[str], failed: list[dict]):
    print("\n" + "=" * 50)
    print("處理結果摘要")
    print("=" * 50)
    print(f"  Excel 總筆數  : {total}")
    print(f"  成功寫入      : {inserted}")
    print(f"  已有紀錄(略過): {len(already_exists)}")
    print(f"  學籍核對失敗  : {len(failed)}")

    if already_exists:
        print("\n已有免修紀錄（略過）:")
        for sid in already_exists:
            print(f"    {sid}")

    if failed:
        print("\n學籍核對失敗清單:")
        for item in failed:
            reason = item.get("原因", "")
            sid    = item.get("學號", "")
            if reason == "姓名不符":
                print(f"    {sid}  申請姓名: {item.get('姓名(申請)','')}  學籍姓名: {item.get('姓名(學籍)','')}")
            else:
                print(f"    {sid}  {item.get('姓名','')}  ({reason})")
    print("=" * 50)


# ────────────────────────────────────────────────────────
# 主流程
# ────────────────────────────────────────────────────────

def main():
    excel_path = Path(EXCEL_FILE)
    if not excel_path.exists():
        print(f"[錯誤] 找不到 Excel 檔案: {excel_path}")
        sys.exit(1)

    print(f"讀取 Excel: {excel_path.name}  工作表: {SHEET_NAME}")
    df = pd.read_excel(excel_path, sheet_name=SHEET_NAME, skiprows=SKIP_ROWS)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")
    print(f"  共 {len(df)} 筆, 欄位: {list(df.columns)}")

    if COL_STUDENT_ID not in df.columns or COL_NAME not in df.columns:
        print(f"[錯誤] Excel 中找不到欄位「{COL_STUDENT_ID}」或「{COL_NAME}」，請確認欄位名稱設定。")
        sys.exit(1)

    total = len(df)

    # 步驟 1：學籍核對
    verified_df, failed = verify_students(df)

    # 步驟 2：排除已有免修紀錄
    to_insert_df, already_exists = filter_existing_exemptions(verified_df)

    # 步驟 3：寫入
    insert_exemptions(to_insert_df)

    # 報告
    print_report(total, len(to_insert_df), already_exists, failed)


if __name__ == "__main__":
    main()
