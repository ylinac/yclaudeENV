# 進階英語免修自動化處理系統

處理每學期外教中心提供的進階英語免修集體申請通過名單，自動核對學籍、寫入資料庫，並通知相關人員。

## 流程

```
外教中心提供 Excel 名單
        │
        ▼
核對學籍（MSSQL 65）
  - 學號是否存在
  - 姓名是否相符
        │
        ▼
查詢免修紀錄（MSSQL 211 couexe）
  - 已有紀錄 → 略過
  - 無紀錄   → 寫入資料庫
        │
        ▼
產生通過名單 → 公佈網頁 → Email 通知
```

## 檔案說明

| 檔案 | 說明 |
|------|------|
| `excel_to_mssql.py` | 將 Excel 資料批次匯入 MS SQL Server 的通用工具 |
| `1142進階英語免修集體申請通過名單.xlsx` | 114 學年第 2 學期外教中心提供之通過名單 |
| `AI自動化系統改善提案.pptx` | 系統改善提案簡報 |
| `進階英語免修需求.MD` | 需求規格說明 |

## 環境需求

- Python 3.8+
- ODBC Driver 17 for SQL Server

```bash
pip install pandas pyodbc openpyxl
```

## 使用方式

1. 開啟 `excel_to_mssql.py`，修改頂部的連線設定：

```python
DB_SERVER   = "YOUR_SERVER"
DB_NAME     = "YOUR_DATABASE"
DB_USER     = "sa"
DB_PASSWORD = "YOUR_PASSWORD"

EXCEL_FILE  = r"C:\path\to\your_file.xlsx"
TABLE_NAME  = "imported_data"
```

2. 執行：

```bash
python excel_to_mssql.py
```

腳本會自動建立資料表（若不存在）並批次寫入資料。
