# -*- coding: utf-8 -*-
"""
fx.py — 匯率表 (換算成 USD 用)

fx_rates 資料表：每個幣別、每個月一筆「1 USD = 多少該幣別」(per_usd)。
  - 自動：update_fx_rates() 依資料庫裡出現的幣別與發票日期範圍，向
          Frankfurter (歐洲央行 ECB 公布的每日參考匯率) 抓每日匯率，取每月
          平均。只送出幣別代碼與日期，不會送出任何發票資料。
  - 手動：source = '手動' 的列 (在儀表板「資料管理」頁修改的) 不會被自動
          更新覆蓋。
換算時找不到當月匯率 (例如本月 ECB 還沒公布) 就用該幣別最近一個月的匯率。
"""

import json
import urllib.request
from collections import defaultdict
from datetime import date
from typing import Dict, List, Tuple

import duckdb

API_URL = "https://api.frankfurter.dev/v1/{start}..{end}?base=USD&symbols={symbols}"
SOURCE_ECB = "ECB 月平均"
SOURCE_MANUAL = "手動"
TIMEOUT = 20


def ensure_table(con: duckdb.DuckDBPyConnection) -> None:
    con.execute("""CREATE TABLE IF NOT EXISTS fx_rates (
        currency VARCHAR, year INTEGER, month INTEGER, per_usd DOUBLE,
        source VARCHAR, updated_at TIMESTAMP DEFAULT current_timestamp,
        PRIMARY KEY (currency, year, month))""")


def fetch_monthly_rates(currencies: List[str], start: date, end: date
                        ) -> Tuple[Dict[Tuple[str, int, int], float], List[str]]:
    """回傳 ({(幣別, 年, 月): 1 USD 兌換數}, ECB 沒有提供的幣別)。
    一次抓一年，避免單次請求範圍太大。"""
    sums: Dict[Tuple[str, int, int], List[float]] = defaultdict(list)
    seen = set()
    for year in range(start.year, end.year + 1):
        s = max(start, date(year, 1, 1))
        e = min(end, date(year, 12, 31))
        url = API_URL.format(start=s.isoformat(), end=e.isoformat(), symbols=",".join(currencies))
        req = urllib.request.Request(url, headers={"User-Agent": "expense-dashboard"})
        with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
            data = json.load(resp)
        for day, rates in data.get("rates", {}).items():
            d = date.fromisoformat(day)
            if not (s <= d <= e):  # API 會往前補到最近一個有匯率的營業日
                continue
            for cur, value in rates.items():
                sums[(cur, d.year, d.month)].append(float(value))
                seen.add(cur)
    monthly = {k: sum(v) / len(v) for k, v in sums.items()}
    return monthly, [c for c in currencies if c not in seen]


def update_fx_rates(con: duckdb.DuckDBPyConnection) -> str:
    """依資料庫裡的幣別與日期範圍更新 ECB 月平均匯率，回傳結果說明。
    網路失敗會丟出例外，由呼叫端決定要不要中斷。"""
    ensure_table(con)
    row = con.execute("""SELECT list(DISTINCT currency), min(invoice_date), max(invoice_date)
        FROM expense_items WHERE currency IS NOT NULL AND currency <> 'USD'
          AND invoice_date IS NOT NULL""").fetchone()
    currencies, start, end = row
    if not currencies:
        return "資料裡沒有需要換匯的幣別"
    end = min(end, date.today())
    monthly, missing = fetch_monthly_rates(sorted(currencies), start, end)
    rows = [(cur, y, m, rate, SOURCE_ECB) for (cur, y, m), rate in monthly.items()]
    if rows:
        # 手動輸入的匯率不覆蓋
        con.executemany(f"""
            INSERT INTO fx_rates (currency, year, month, per_usd, source) VALUES (?, ?, ?, ?, ?)
            ON CONFLICT (currency, year, month) DO UPDATE
              SET per_usd = excluded.per_usd, source = excluded.source,
                  updated_at = now()
              WHERE fx_rates.source <> '{SOURCE_MANUAL}'""", rows)
    msg = f"已更新 {len(rows)} 筆月平均匯率 ({', '.join(sorted(currencies))}，{start:%Y-%m} ~ {end:%Y-%m})"
    if missing:
        msg += f"；ECB 沒有提供 {', '.join(missing)}，請在匯率表手動輸入"
    return msg


def rate_sql(fixed: Dict[str, float] = None) -> Tuple[str, list]:
    """換算成 USD 的 items CTE：amount / vat 換成 USD，另加 fx_rate (1 USD =
    多少原幣) 與 amount_orig (原幣金額)。fixed：{幣別: 1 USD 兌換數}，有
    給的幣別一律用固定匯率。回傳 (SQL, 參數)。"""
    fixed = fixed or {}
    fixed_cases = " ".join("WHEN e.currency = ? THEN ?" for _ in fixed)
    params: list = []
    for cur, rate in fixed.items():
        params.extend([cur, rate])
    sql = f"""
    latest_fx AS (
        SELECT currency, arg_max(per_usd, year * 100 + month) AS per_usd
        FROM fx_rates GROUP BY currency),
    items_fx AS (
        SELECT e.*, CASE WHEN e.currency = 'USD' THEN 1.0 {fixed_cases}
                         ELSE coalesce(f.per_usd, l.per_usd) END AS fx_rate
        FROM expense_items e
        LEFT JOIN fx_rates f ON f.currency = e.currency AND f.year = e.year AND f.month = e.month
        LEFT JOIN latest_fx l ON l.currency = e.currency),
    items AS (
        SELECT * EXCLUDE (amount, vat), amount AS amount_orig,
               amount / fx_rate AS amount, vat / fx_rate AS vat
        FROM items_fx)"""
    return sql, params
