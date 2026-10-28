# -*- coding: utf-8 -*-
"""
dashboard.py — 主管級費用分析儀表板頁 (由 app.py 載入)

讀 process_data.py 產生的 DuckDB (位置見 settings.py)。所有彙總都在 DuckDB
裡算完才傳回來，畫面只拿小量結果，明細表格也只顯示前 MAX_TABLE_ROWS 列，
資料再多也秒開。每次查詢用唯讀連線、查完就關，資料更新時 (資料庫檔案修改
時間變了) 快取自動失效。

金額可選「換算 USD」(作帳匯率檔最新一期，或自訂固定匯率) 或
「原幣」。所有查詢都從 items 這個 CTE 讀，它依幣別設定把 amount / vat
換好，下面的圖表/KPI 不用管換匯。
"""

import io
from typing import Dict, List, Optional, Tuple

import duckdb
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import fx
import pivot_export
from expense_normalize import EXPENSE_CATEGORIES, OTHER_CATEGORY
from settings import get_db_path

MAX_TABLE_ROWS = 2_000
EXCEL_MAX_ROWS = 1_000_000
MODE_USD, MODE_ORIG = "換算 USD", "原幣"
FX_MONTHLY, FX_FIXED = "作帳匯率", "固定匯率"

# 類別色 (固定順序，顏色跟著項目走，不跟排名走)
SERIES_COLORS = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100",
                 "#e87ba4", "#008300", "#6250d6", "#e34948"]
OTHER_COLOR = "#a3a29c"
CATEGORY_ORDER = [name for name, _ in EXPENSE_CATEGORIES]
CATEGORY_COLORS = {name: SERIES_COLORS[i] for i, name in enumerate(CATEGORY_ORDER)}
CATEGORY_COLORS[OTHER_CATEGORY] = OTHER_COLOR

# 側邊欄篩選器：(資料庫欄位, 標籤)
FILTER_FIELDS = (
    ("supplier", "供應商 (SUPPLIER)"),
    ("expense_category", "費用大類"),
    ("expense_item", "費用項目 (Description)"),
    ("port_of_loading", "啟運港 (Port of Loading)"),
    ("port_of_discharge", "目的港 (Port of Discharge)"),
    ("consignee", "集團公司 (CONSIGNEE)"),
    ("currency", "原幣別"),
)


# ============================================================
# 資料存取
# ============================================================

@st.cache_data(show_spinner=False, max_entries=256)
def query(db: str, db_mtime: float, sql: str, params: Tuple = ()) -> pd.DataFrame:
    """唯讀連線執行查詢後立刻關閉。db_mtime 只用來讓資料更新後快取失效。"""
    con = duckdb.connect(db, read_only=True)
    try:
        return con.execute(sql, list(params)).df()
    finally:
        con.close()


def build_where(filters: Dict[str, List], years: Optional[List[int]],
                include_dup: bool, cutoff: Optional[Tuple[int, int]] = None) -> Tuple[str, Tuple]:
    """篩選條件 -> (WHERE 子句, 參數)。cutoff：(月, 日)，只取每年 1/1 到
    該月日 (同期比較用，避免拿不完整的最後一個月去比去年整個月)。"""
    clauses, params = [], []
    if not include_dup:
        clauses.append("NOT is_duplicate")
    if years is not None:
        clauses.append("list_contains(?, year)")
        params.append(years)
    if cutoff is not None:
        clauses.append("(month < ? OR (month = ? AND day(invoice_date) <= ?))")
        params.extend([cutoff[0], cutoff[0], cutoff[1]])
    for col, values in filters.items():
        if values:
            clauses.append(f"list_contains(?, {col})")
            params.append(values)
    return ("WHERE " + " AND ".join(clauses)) if clauses else "", tuple(params)


def fmt_big(value: Optional[float]) -> str:
    """大金額縮寫成 億/萬，KPI 卡片用。"""
    if value is None or pd.isna(value):
        return "—"
    sign = "-" if value < 0 else ""
    v = abs(value)
    if v >= 1e8:
        return f"{sign}{v / 1e8:,.2f} 億"
    if v >= 1e4:
        return f"{sign}{v / 1e4:,.1f} 萬"
    return f"{sign}{v:,.0f}"


def pct_change(cur: Optional[float], prev: Optional[float]) -> Optional[str]:
    if cur is None or prev is None or pd.isna(cur) or pd.isna(prev) or prev == 0:
        return None
    return f"{(cur - prev) / abs(prev):+.1%}"


def style_fig(fig: go.Figure, height: int = 420) -> go.Figure:
    """統一圖表樣式：細格線、無多餘外框、懸停提示。"""
    fig.update_layout(
        height=height, margin=dict(l=8, r=8, t=8, b=8),
        hoverlabel=dict(namelength=-1), legend=dict(orientation="h", y=-0.15),
        bargap=0.25,
    )
    fig.update_xaxes(showgrid=False, zeroline=False)
    fig.update_yaxes(gridcolor="rgba(128,128,128,0.18)", zeroline=False)
    return fig


# ============================================================
# 資料庫
# ============================================================

db_path = get_db_path()
if not db_path.is_file():
    st.title("費用分析儀表板")
    st.error(f"找不到資料庫：{db_path}")
    st.info("請到左側「資料管理」頁：指定資料庫位置 (例如公司共用資料夾裡的 "
            "expenses.duckdb)，或從轉檔結果資料夾建立新的資料庫。")
    st.stop()
DB, MTIME = str(db_path), db_path.stat().st_mtime


def sql_df(sql: str, params: Tuple = ()) -> pd.DataFrame:
    """執行查詢；資料庫正在被更新 (被鎖住) 時提示稍後再試，不讓整頁報錯。"""
    try:
        return query(DB, MTIME, sql, params)
    except duckdb.Error as e:
        if "lock" in str(e).lower() or "being used" in str(e).lower():
            st.warning("資料庫正在更新中，請稍候幾秒再重新整理頁面。")
            st.stop()
        raise


def options(col: str) -> List:
    df = sql_df(f"SELECT DISTINCT {col} AS v FROM expense_items "
                f"WHERE {col} IS NOT NULL ORDER BY 1")
    return df["v"].tolist()


# ============================================================
# 側邊欄
# ============================================================

st.sidebar.header("💱 幣別")
has_fx = bool(sql_df("SELECT count(*) AS n FROM information_schema.tables "
                     "WHERE table_name = 'fx_rates'")["n"].iloc[0])
currencies = options("currency") if "currency" in sql_df(
    "SELECT * FROM expense_items LIMIT 0").columns else []
if not has_fx or not currencies:
    st.sidebar.caption("這份資料庫還沒有幣別/匯率資料，請到「資料管理」頁更新資料後再換算 USD。")
    mode = MODE_ORIG
else:
    mode = st.sidebar.radio("金額顯示", [MODE_USD, MODE_ORIG], horizontal=True)

fixed_rates: Dict[str, float] = {}
if mode == MODE_USD:
    fx_mode = st.sidebar.radio("匯率", [FX_MONTHLY, FX_FIXED], horizontal=True,
                               help="每月平均匯率：依發票月份換算 (當月沒有匯率時用最近一個月)")
    if fx_mode == FX_FIXED:
        latest_rates = sql_df("SELECT currency, arg_max(per_usd, year * 100 + month) AS r "
                              "FROM fx_rates GROUP BY 1")
        latest_map = dict(zip(latest_rates["currency"], latest_rates["r"]))
        for cur in currencies:
            if cur != "USD":
                fixed_rates[cur] = st.sidebar.number_input(
                    f"1 USD = ? {cur}", min_value=0.0001,
                    value=float(round(latest_map.get(cur, 1.0), 2)), step=100.0, format="%.2f")
    cte_sql, cte_params = fx.rate_sql(fixed_rates)
    CTE, CTE_PARAMS = "WITH " + cte_sql, tuple(cte_params)
    unit = "USD"
else:
    CTE = ("WITH items AS (SELECT *, amount AS amount_orig, "
           "CAST(NULL AS DOUBLE) AS fx_rate FROM expense_items)")
    CTE_PARAMS = ()
    unit = None  # 原幣：等篩選完才知道是哪個幣別


def q(body: str, params: Tuple = ()) -> pd.DataFrame:
    """從 items (已依幣別設定換算的明細) 查詢。"""
    return sql_df(f"{CTE}\n{body}", CTE_PARAMS + tuple(params))


st.sidebar.header("📊 篩選條件")
all_years = [int(y) for y in options("year")]
if not all_years:
    st.warning("資料庫裡沒有可辨識日期的資料。")
    st.stop()
default_years = [y for y in all_years if y in (2025, 2026)] or all_years[-2:]
years = st.sidebar.multiselect("年份", all_years, default=default_years)
if not years:
    st.sidebar.warning("請至少選一個年份")
    st.stop()

filters: Dict[str, List] = {}
for col, label in FILTER_FIELDS:
    if col == "currency" and len(currencies) <= 1:
        continue
    opts = options(col)
    # 費用項目跟著已選的費用大類縮小選項
    if col == "expense_item" and filters.get("expense_category"):
        opts = sql_df("SELECT DISTINCT expense_item AS v FROM expense_items "
                      "WHERE list_contains(?, expense_category) ORDER BY 1",
                      (filters["expense_category"],))["v"].tolist()
    filters[col] = st.sidebar.multiselect(label, opts, placeholder="全部")

include_dup = st.sidebar.checkbox("包含重複發票", value=False,
                                  help="同一張發票 (供應商+發票號碼) 出現在好幾份 Excel 時，預設只算一次")
st.sidebar.caption(f"資料庫：{db_path}\n\n資料更新時間："
                   f"{pd.Timestamp(MTIME, unit='s', tz='UTC').tz_convert('Asia/Taipei'):%Y-%m-%d %H:%M}")

where, params = build_where(filters, years, include_dup)

# 原幣模式：篩選後的資料只有一種幣別才能加總
if mode == MODE_ORIG:
    used = q(f"SELECT DISTINCT currency AS c FROM items {where}", params)["c"].dropna().tolist()
    if len(used) > 1:
        st.title("費用分析儀表板")
        st.warning(f"目前篩選的資料混合了 {len(used)} 種幣別 ({'、'.join(used)})，原幣金額不能直接加總。"
                   "請在側邊欄「原幣別」只選一種，或切換成「換算 USD」。")
        st.stop()
    unit = used[0] if used else ""

# ============================================================
# KPI
# ============================================================

st.title("費用分析儀表板")
if mode == MODE_USD:
    fx_note = ("依作帳匯率換算" if not fixed_rates else
               "固定匯率：" + "、".join(f"1 USD = {r:,.2f} {c}" for c, r in fixed_rates.items()))
else:
    fx_note = "未換匯"
st.caption(f"年份：{'、'.join(map(str, sorted(years)))}　｜　金額單位：{unit}（{fx_note}）")

KPI_SQL = """SELECT sum(amount) AS amount, sum(vat) AS vat,
    count(DISTINCT coalesce(supplier, '') || '|' || coalesce(invoice_no, source_file)) AS invoices,
    count(*) AS lines, count(*) FILTER (WHERE amount IS NULL AND amount_orig IS NOT NULL) AS no_rate
    FROM items {where}"""
kpi = q(KPI_SQL.format(where=where), params).iloc[0]
if not kpi["lines"]:
    st.warning("沒有符合篩選條件的資料。")
    st.stop()
if kpi["no_rate"]:
    st.warning(f"有 {int(kpi['no_rate']):,} 筆明細的幣別找不到匯率，沒有計入 USD 金額。"
               "請到「資料管理」頁更新或手動輸入匯率。")
avg_invoice = kpi["amount"] / kpi["invoices"] if kpi["invoices"] else None

# 同期比較：最新選取年份 vs 前一年，都只取 1/1 ~ 最新年份資料的最後一天
latest = max(years)
cur_kpi = prev_kpi = None
w, p = build_where(filters, [latest], include_dup)
last_day = q(f"SELECT max(invoice_date) AS d FROM items {w}", p)["d"].iloc[0]
if last_day is not None and not pd.isna(last_day):
    last_day = pd.Timestamp(last_day)
    cutoff = (last_day.month, last_day.day)
    w, p = build_where(filters, [latest], include_dup, cutoff)
    cur_kpi = q(KPI_SQL.format(where=w), p).iloc[0]
    w, p = build_where(filters, [latest - 1], include_dup, cutoff)
    prev_kpi = q(KPI_SQL.format(where=w), p).iloc[0]
    if not prev_kpi["lines"]:
        prev_kpi = None


def yoy(field: str, per_invoice: bool = False) -> Optional[str]:
    if cur_kpi is None or prev_kpi is None:
        return None
    if per_invoice:
        cur = cur_kpi["amount"] / cur_kpi["invoices"] if cur_kpi["invoices"] else None
        prev = prev_kpi["amount"] / prev_kpi["invoices"] if prev_kpi["invoices"] else None
        return pct_change(cur, prev)
    return pct_change(cur_kpi[field], prev_kpi[field])


money_fmt = "{:,.2f}" if unit == "USD" else "{:,.0f}"
c1, c2, c3, c4 = st.columns(4)
c1.metric(f"總費用金額 ({unit})", fmt_big(kpi["amount"]), yoy("amount"), border=True,
          help=money_fmt.format(kpi["amount"] or 0))
c2.metric(f"總 VAT ({unit})", fmt_big(kpi["vat"]), yoy("vat"), border=True,
          help=money_fmt.format(kpi["vat"] or 0))
c3.metric("總發票數", f"{int(kpi['invoices']):,} 張", yoy("invoices"), border=True,
          help=f"費用明細 {int(kpi['lines']):,} 筆")
c4.metric(f"平均單筆費用 ({unit}/每張發票)", fmt_big(avg_invoice), yoy("amount", per_invoice=True),
          border=True, help=money_fmt.format(avg_invoice) if avg_invoice else None)
if prev_kpi is not None:
    st.caption(f"增減率：{latest} 年 1/1–{cutoff[0]}/{cutoff[1]} vs {latest - 1} 年同期 "
               f"(以發票日期計，套用相同篩選條件"
               + ("；USD 金額會同時反映匯率變動" if mode == MODE_USD and not fixed_rates else "") + ")")
else:
    st.caption(f"增減率：{latest - 1} 年同期沒有資料，無法比較")

# ============================================================
# 費用項目排行榜 + 月度趨勢
# ============================================================

left, right = st.columns([5, 6])

with left:
    st.subheader("費用項目排行榜")
    a, b = st.columns([3, 2])
    dim = a.radio("依", ["費用項目", "費用大類"], horizontal=True, label_visibility="collapsed")
    top_n = b.slider("前幾名", 5, 30, 10, label_visibility="collapsed")
    dim_col = "expense_item" if dim == "費用項目" else "expense_category"
    top = q(f"""
        SELECT {dim_col} AS name, sum(amount) AS amount, count(*) AS lines
        FROM items {where} GROUP BY 1 ORDER BY 2 DESC NULLS LAST LIMIT {int(top_n)}""", params)
    top["share"] = top["amount"] / kpi["amount"] if kpi["amount"] else 0
    top = top.iloc[::-1]  # 橫條圖由上往下 = 金額由大到小
    fig = go.Figure(go.Bar(
        x=top["amount"], y=top["name"], orientation="h", marker_color=SERIES_COLORS[0],
        marker_cornerradius=4, text=[fmt_big(v) for v in top["amount"]],
        textposition="outside", cliponaxis=False,
        customdata=top[["share", "lines"]].to_numpy(),
        hovertemplate=f"<b>%{{y}}</b><br>金額 %{{x:,.0f}} {unit}<br>佔比 %{{customdata[0]:.1%}}"
                      "<br>明細 %{customdata[1]:,} 筆<extra></extra>",
    ))
    fig.update_xaxes(showticklabels=False)
    st.plotly_chart(style_fig(fig, height=max(320, 34 * len(top) + 40)), width="stretch")

with right:
    st.subheader(f"月度費用趨勢 (年度比較，{unit})")
    trend = q(f"""
        SELECT year, month, sum(amount) AS amount FROM items {where}
        GROUP BY 1, 2 ORDER BY 1, 2""", params)
    fig = go.Figure()
    # 年份顏色依「全部年份」的固定順序指定，篩選年份時顏色不會跑掉
    for y in sorted(trend["year"].unique()):
        d = trend[trend["year"] == y]
        color = SERIES_COLORS[all_years.index(int(y)) % len(SERIES_COLORS)]
        fig.add_trace(go.Scatter(
            x=d["month"], y=d["amount"], name=f"{int(y)} 年", mode="lines+markers",
            line=dict(color=color, width=2), marker=dict(size=8, color=color),
            hovertemplate=f"{int(y)} 年 %{{x}} 月：%{{y:,.0f}} {unit}<extra></extra>",
        ))
    fig.update_layout(hovermode="x unified")
    fig.update_xaxes(tickmode="array", tickvals=list(range(1, 13)),
                     ticktext=[f"{m}月" for m in range(1, 13)], range=[0.6, 12.4])
    st.plotly_chart(style_fig(fig), width="stretch")
    if cur_kpi is not None and not last_day.is_month_end:
        st.caption(f"{latest} 年資料只到 {last_day:%m/%d}，{last_day.month} 月是不完整月份。")

# 年度比較表 (主管最常問：哪一項漲最多)
with st.expander("📋 費用項目年度比較表", expanded=False):
    by_year = q(f"""
        SELECT expense_item, year, sum(amount) AS amount FROM items {where}
        GROUP BY 1, 2""", params)
    comp = by_year.pivot_table(index="expense_item", columns="year", values="amount",
                               aggfunc="sum")
    comp.columns = [f"{int(c)} 年" for c in comp.columns]
    year_cols = list(comp.columns)
    comp = comp.reset_index().rename(columns={"expense_item": "費用項目"})
    if len(year_cols) >= 2:
        a, b = year_cols[-2], year_cols[-1]
        comp["增減金額"] = comp[b].fillna(0) - comp[a].fillna(0)
        comp["增減率"] = comp["增減金額"] / comp[a].where(comp[a] != 0)
    comp = comp.sort_values(year_cols[-1], ascending=False, na_position="last")
    fmt = {c: money_fmt for c in year_cols + ["增減金額"] if c in comp}
    if "增減率" in comp:
        fmt["增減率"] = "{:+.1%}"
    st.dataframe(comp.style.format(fmt, na_rep="—"), hide_index=True, width="stretch")
    st.caption(f"金額單位：{unit}。此表為所選年份的全年合計；KPI 增減率為同期 (1/1 至最新資料日期) 比較。")

# ============================================================
# 港口 / 航線成本分析
# ============================================================

st.subheader("港口 / 航線成本分析")
tab_pod, tab_route = st.tabs(["目的港 × 費用大類", "航線 (啟運港 → 目的港)"])
TOP_PORTS = 10

with tab_pod:
    pod = q(f"""
        , t AS (
            SELECT coalesce(port_of_discharge, '(未填)') AS port, expense_category AS cat,
                   sum(amount) AS amount
            FROM items {where} GROUP BY 1, 2),
        ranked AS (
            SELECT port, rank() OVER (ORDER BY sum(amount) DESC) AS rk FROM t GROUP BY port)
        SELECT CASE WHEN r.rk <= {TOP_PORTS} THEN t.port ELSE '其他港口' END AS port,
               t.cat, sum(t.amount) AS amount
        FROM t JOIN ranked r USING (port) GROUP BY 1, 2""", params)
    totals = pod.groupby("port")["amount"].sum().sort_values()
    # 「其他港口」固定放最下面
    order = [p for p in totals.index if p == "其他港口"] + [p for p in totals.index if p != "其他港口"]
    fig = go.Figure()
    for cat in CATEGORY_ORDER + [OTHER_CATEGORY]:
        d = pod[pod["cat"] == cat].set_index("port").reindex(order)
        if d["amount"].notna().any():
            share = d["amount"] / totals.reindex(order)
            fig.add_trace(go.Bar(
                y=order, x=d["amount"], name=cat, orientation="h",
                marker=dict(color=CATEGORY_COLORS[cat], line=dict(width=1, color="rgba(255,255,255,0.9)")),
                customdata=share, hovertemplate=f"<b>%{{y}}</b><br>{cat}：%{{x:,.0f}} {unit}"
                                                f"<br>佔該港 %{{customdata:.1%}}<extra></extra>"))
    total_all = totals.sum()
    fig.add_trace(go.Scatter(  # 每列右邊標總額與佔比
        y=order, x=totals.reindex(order), mode="text", showlegend=False, hoverinfo="skip",
        text=[f"  {fmt_big(v)} ({v / total_all:.0%})" for v in totals.reindex(order)],
        textposition="middle right"))
    fig.update_layout(barmode="stack")
    fig.update_xaxes(showticklabels=False)
    st.plotly_chart(style_fig(fig, height=max(360, 36 * len(order) + 80)), width="stretch")
    st.caption(f"依目的港費用總額排序，前 {TOP_PORTS} 大以外合併為「其他港口」；括號為佔全部費用比例。")

with tab_route:
    route = q(f"""
        SELECT coalesce(port_of_loading, '(未填)') || ' → ' || coalesce(port_of_discharge, '(未填)')
                   AS route,
               sum(amount) AS amount,
               count(DISTINCT coalesce(supplier, '') || '|' || coalesce(invoice_no, source_file)) AS invoices
        FROM items {where} GROUP BY 1 ORDER BY 2 DESC NULLS LAST LIMIT 15""", params)
    route["share"] = route["amount"] / kpi["amount"]
    route = route.iloc[::-1]
    fig = go.Figure(go.Bar(
        x=route["amount"], y=route["route"], orientation="h", marker_color=SERIES_COLORS[0],
        marker_cornerradius=4,
        text=[f"{fmt_big(v)} ({s:.0%})" for v, s in zip(route["amount"], route["share"])],
        textposition="outside", cliponaxis=False,
        customdata=route[["invoices", "share"]].to_numpy(),
        hovertemplate=f"<b>%{{y}}</b><br>金額 %{{x:,.0f}} {unit}<br>佔比 %{{customdata[1]:.1%}}"
                      "<br>發票 %{customdata[0]:,} 張<extra></extra>"))
    fig.update_xaxes(showticklabels=False)
    st.plotly_chart(style_fig(fig, height=max(360, 34 * len(route) + 40)), width="stretch")
    st.caption("費用最高的前 15 條航線；括號為佔全部費用比例。")

if mode == MODE_USD and not fixed_rates:
    with st.expander("💱 本頁使用的月平均匯率 (1 USD = ?)", expanded=False):
        rates = sql_df("SELECT currency AS 幣別, year AS 年, month AS 月, per_usd AS 匯率, "
                       "source AS 來源 FROM fx_rates WHERE list_contains(?, year) "
                       "ORDER BY 1, 2, 3", (years,))
        st.dataframe(rates.style.format({"匯率": "{:,.2f}"}), hide_index=True, width="stretch")

# ============================================================
# 明細檢視與匯出
# ============================================================

st.subheader("明細資料")
detail_columns = {
    "invoice_date": "INVOICE DATE 發票日期", "invoice_no": "INVOICE NO 發票號碼",
    "supplier": "SUPPLIER 供應商", "consignee": "CONSIGNEE 集團公司",
    "expense_category": "費用大類", "expense_item": "費用項目 (標準化)",
    "description_raw": "Description 原始費用名稱", "currency": "原幣別",
    "amount_orig": "AMOUNT 原幣金額",
}
if mode == MODE_USD:
    detail_columns.update({"fx_rate": "匯率 (1 USD = ?)", "amount": "AMOUNT (USD)",
                           "vat": "VAT (USD)"})
else:
    detail_columns["vat"] = "VAT"
detail_columns.update({
    "order_no": "Order No.", "bl_no": "B/L NO.", "port_of_loading": "Port of Loading 啟運港",
    "port_of_discharge": "Port of discharge 目的港", "onboard_date": "開船日",
    "arrive_date": "到達日", "volume": "Volume 材積", "container_no": "Container no.",
    "shipper": "Shipper", "source_file": "來源檔案",
})
detail_body = (f"SELECT {', '.join(detail_columns)} FROM items {where} "
               "ORDER BY invoice_date DESC, invoice_no")
n_lines = int(kpi["lines"])
detail = q(f"{detail_body} LIMIT {MAX_TABLE_ROWS}", params).rename(columns=detail_columns)
num_cols = {c: st.column_config.NumberColumn(format="%,.2f" if "USD" in c or "匯率" in c else "%,.0f")
            for k, c in detail_columns.items() if k in ("amount_orig", "amount", "vat", "fx_rate")}
st.dataframe(detail, hide_index=True, width="stretch", height=420, column_config=num_cols)
st.caption(f"符合條件共 {n_lines:,} 筆明細"
           + (f"，畫面只顯示最新 {MAX_TABLE_ROWS:,} 筆；匯出檔包含全部。" if n_lines > MAX_TABLE_ROWS else "。"))


@st.cache_data(show_spinner=False, max_entries=8)
def export_csv(db: str, db_mtime: float, sql: str, params: Tuple, columns: Tuple) -> bytes:
    df = query(db, db_mtime, sql, params).rename(columns=dict(columns))
    return df.to_csv(index=False).encode("utf-8-sig")  # 加 BOM，Excel 直接開不會亂碼


@st.cache_data(show_spinner=False, max_entries=4)
def export_excel(db: str, db_mtime: float, sql: str, params: Tuple, columns: Tuple) -> bytes:
    df = query(db, db_mtime, sql, params).rename(columns=dict(columns))
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="明細")
        sheet = writer.sheets["明細"]
        sheet.freeze_panes = "A2"
        for i, col in enumerate(df.columns, start=1):
            sheet.column_dimensions[sheet.cell(row=1, column=i).column_letter].width = \
                min(max(len(str(col)) + 2, 12), 40)
    return buf.getvalue()


@st.cache_data(show_spinner=False, max_entries=2)
def export_pivot(db: str, db_mtime: float, where: str, params: Tuple, fixed: Tuple) -> bytes:
    """可樞紐分析的 Excel (pivot_export)：一律換算 USD，沿用目前的篩選與固定匯率設定。"""
    cte, cte_params = pivot_export.items_cte(dict(fixed))
    df = query(db, db_mtime, f"{cte}\n{pivot_export.detail_select(where)}",
               tuple(cte_params) + params)
    failed = query(db, db_mtime, "SELECT source_file, error FROM etl_files "
                   "WHERE error IS NOT NULL ORDER BY 1")
    buf = io.BytesIO()
    pivot_export.write_workbook(df, failed, buf)
    return buf.getvalue()


export_args = (DB, MTIME, f"{CTE}\n{detail_body}", CTE_PARAMS + params,
               tuple(detail_columns.items()))
export_key = export_args[2:4]
stamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M")
e1, e2, e3, _ = st.columns([1, 1, 1.4, 1.6])
with e1:
    if st.button("⬇ 準備 CSV", width="stretch"):
        with st.spinner("產生 CSV 中..."):
            st.session_state["csv"] = (export_key, export_csv(*export_args))
    csv = st.session_state.get("csv")
    if csv and csv[0] == export_key:
        st.download_button("下載 CSV", csv[1], f"費用明細_{stamp}.csv", "text/csv",
                           width="stretch", type="primary")
with e2:
    if n_lines > EXCEL_MAX_ROWS:
        st.button("⬇ 準備 Excel", disabled=True, width="stretch",
                  help=f"超過 Excel 上限 {EXCEL_MAX_ROWS:,} 列，請縮小範圍或改用 CSV")
    else:
        if st.button("⬇ 準備 Excel", width="stretch"):
            with st.spinner(f"產生 Excel 中 ({n_lines:,} 筆)..."):
                st.session_state["xlsx"] = (export_key, export_excel(*export_args))
        xlsx = st.session_state.get("xlsx")
        if xlsx and xlsx[0] == export_key:
            st.download_button("下載 Excel", xlsx[1], f"費用明細_{stamp}.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               width="stretch", type="primary")
with e3:
    if n_lines > EXCEL_MAX_ROWS:
        st.button("⬇ 準備樞紐分析 Excel", disabled=True, width="stretch",
                  help=f"超過 Excel 上限 {EXCEL_MAX_ROWS:,} 列，請縮小範圍或改用 CSV")
    else:
        if st.button("⬇ 準備樞紐分析 Excel", width="stretch",
                     help="含儀表板頁、4 張樞紐分析表與圖、交叉分析篩選器；金額一律換算 USD。"
                          "開檔後按「啟用編輯」，樞紐分析表會自動重新整理。"):
            with st.spinner(f"產生樞紐分析 Excel 中 ({n_lines:,} 筆)..."):
                st.session_state["pivot"] = (export_key, export_pivot(
                    DB, MTIME, where, params, tuple(sorted(fixed_rates.items()))))
        pivot = st.session_state.get("pivot")
        if pivot and pivot[0] == export_key:
            st.download_button("下載樞紐分析 Excel", pivot[1], f"費用分析樞紐_{stamp}.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               width="stretch", type="primary")
