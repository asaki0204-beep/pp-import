import io
import math
import re
from datetime import date

import pandas as pd
import requests
import streamlit as st

# ── 定数 ──────────────────────────────────────────────────────────────────────
MURC_XLS_BASE = "https://www.murc-kawasesouba.jp/fx/xls"
MURC_PAGE_URL = "https://www.murc-kawasesouba.jp/fx/past_3month.php"
OUTPUT_ENC    = "cp932"
KNOWN_CCY     = {
    "USD", "EUR", "GBP", "AUD", "CAD", "CHF",
    "NZD", "SGD", "HKD", "CNY", "THB", "PHP",
}

# ── ユーティリティ ─────────────────────────────────────────────────────────────
def parse_number(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    s = s.str.replace(r"^\((.+)\)$", r"-\1", regex=True)
    s = s.str.replace(",", "", regex=False)
    return pd.to_numeric(s, errors="coerce")


def fmt_jpy(value) -> str:
    if pd.isna(value):
        return ""
    return f"{round(value):,}"


def _jpy_round(x) -> float:
    if pd.isna(x):
        return x
    return float(math.floor(abs(x) + 0.5) * (1 if x >= 0 else -1))


def _prev_ym(year: int, month: int) -> tuple:
    return (year - 1, 12) if month == 1 else (year, month - 1)


def _next_ym(year: int, month: int) -> tuple:
    return (year + 1, 1) if month == 12 else (year, month + 1)


def _parse_ym(v):
    if hasattr(v, "year"):
        return v.year, v.month
    s = str(v).strip()
    m = re.search(r"(\d{4})年\s*(\d{1,2})月", s)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = re.fullmatch(r"(\d{4})(\d{2})", s)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = re.fullmatch(r"(\d{4})[/\-](\d{1,2})", s)
    if m:
        return int(m.group(1)), int(m.group(2))
    return None, None


def _parse_dates_multi(series: pd.Series) -> pd.Series:
    result = pd.to_datetime(series, format="%Y/%m/%d", errors="coerce")
    mask = result.isna()
    if mask.any():
        result[mask] = pd.to_datetime(series[mask], format="%m/%d/%Y", errors="coerce")
    mask = result.isna()
    if mask.any():
        result[mask] = pd.to_datetime(series[mask], format="%Y年%m月%d日", errors="coerce")
    return result


# ── CSV 読み込み（バイト対応） ─────────────────────────────────────────────────
def read_csv_from_bytes(data: bytes) -> pd.DataFrame:
    for enc in ("utf-8-sig", "cp932", "utf-8"):
        try:
            return pd.read_csv(io.BytesIO(data), encoding=enc)
        except (UnicodeDecodeError, LookupError):
            continue
    raise ValueError("CSVの読み込みに失敗しました")


# ── Excel エンジン検出（バイト対応） ──────────────────────────────────────────
def _detect_engine_from_bytes(data: bytes) -> str:
    magic = data[:4]
    if magic == b"\xd0\xcf\x11\xe0":
        return "xlrd"
    if magic[:2] == b"PK":
        return "openpyxl"
    return "xlrd"


# ── MURC Excel パーサー ────────────────────────────────────────────────────────
def _parse_murc_monthly_usd(df_raw: pd.DataFrame) -> dict:
    """USD 専用パーサー → {(year, month): rate}"""
    rates: dict = {}
    year = None
    for i in range(min(5, len(df_raw))):
        for v in df_raw.iloc[i]:
            m = re.search(r"(\d{4})年", str(v))
            if m:
                year = int(m.group(1))
                break
        if year:
            break
    if year is None:
        return {}
    cur_month = None
    for _, row in df_raw.iterrows():
        c0 = str(row.iloc[0]).strip()
        if re.match(r"^年間", c0):
            break
        mm = re.match(r"^(\d{1,2})月$", c0)
        if mm:
            cur_month = int(mm.group(1))
        c1 = str(row.iloc[1]).strip() if len(row) > 1 else ""
        if c1 == "平均" and cur_month is not None:
            try:
                val = row.iloc[4]
                ttm = float(str(val).replace(",", "").strip())
                if not pd.isna(ttm) and ttm > 0:
                    rates.setdefault((year, cur_month), ttm)
            except (ValueError, IndexError, TypeError):
                pass
    return rates


def _parse_murc_all_currencies(df_raw: pd.DataFrame) -> dict:
    """多通貨複合ヘッダー形式 → {(year, month, currency): rate}"""
    rates: dict = {}
    year = None
    for i in range(min(5, len(df_raw))):
        for v in df_raw.iloc[i]:
            m = re.search(r"(\d{4})年", str(v))
            if m:
                year = int(m.group(1))
                break
        if year:
            break
    if year is None:
        return {}

    ttm_col_indices: list = []
    currency_col_map: dict = {}
    for i in range(min(12, len(df_raw))):
        row_vals = [str(v).strip() for v in df_raw.iloc[i]]
        if row_vals.count("TTM") >= 2:
            ttm_col_indices = [j for j, v in enumerate(row_vals) if v == "TTM"]
        found = [(j, v) for j, v in enumerate(row_vals) if v in KNOWN_CCY]
        if found:
            for j, code in found:
                currency_col_map[j] = code

    currency_ttm: dict = {}
    for ttm_col in ttm_col_indices:
        best_j, best_code = -1, None
        for cur_j, code in currency_col_map.items():
            if cur_j <= ttm_col and cur_j > best_j:
                best_j, best_code = cur_j, code
        if best_code and best_code not in currency_ttm:
            currency_ttm[best_code] = ttm_col

    if not currency_ttm:
        currency_ttm = {"USD": 4}

    cur_month = None
    for _, row in df_raw.iterrows():
        c0 = str(row.iloc[0]).strip()
        if re.match(r"^年間", c0):
            break
        mm = re.match(r"^(\d{1,2})月$", c0)
        if mm:
            cur_month = int(mm.group(1))
        c1 = str(row.iloc[1]).strip() if len(row) > 1 else ""
        if c1 == "平均" and cur_month is not None:
            for currency, col_idx in currency_ttm.items():
                try:
                    val = row.iloc[col_idx]
                    ttm = float(str(val).replace(",", "").strip())
                    if not pd.isna(ttm) and ttm > 0:
                        rates.setdefault((year, cur_month, currency), ttm)
                except (ValueError, IndexError, TypeError):
                    pass
    return rates


def _parse_murc_archive(df_raw: pd.DataFrame) -> dict:
    """アーカイブ形式 → {(year, month, currency): rate}"""
    _JP_NAME = {
        "米ドル": "USD", "ユーロ": "EUR", "英ポンド": "GBP",
        "豪ドル": "AUD", "カナダドル": "CAD", "スイスフラン": "CHF",
        "NZドル": "NZD", "シンガポールドル": "SGD",
        "香港ドル": "HKD", "中国元": "CNY", "タイバーツ": "THB",
        "フィリピン": "PHP",
    }
    rates: dict = {}
    header_row_idx = None
    for i in range(min(15, len(df_raw))):
        vals = [str(v).strip() for v in df_raw.iloc[i]]
        if sum(1 for v in vals if v.upper() in KNOWN_CCY) >= 2 or "年月" in vals:
            header_row_idx = i
            break
    if header_row_idx is None:
        return {}

    header = [str(v).strip() for v in df_raw.iloc[header_row_idx]]
    col_map: dict = {}
    for j, v in enumerate(header):
        v_up = v.upper()
        if v_up in KNOWN_CCY and v_up not in col_map:
            col_map[v_up] = j
        for jp, code in _JP_NAME.items():
            if jp in v and code not in col_map:
                col_map[code] = j

    if not col_map:
        return {}

    for _, row in df_raw.iloc[header_row_idx + 1:].iterrows():
        ym_val = row.iloc[1] if len(row) > 1 else None
        if ym_val is None:
            continue
        try:
            if pd.isna(ym_val):
                continue
        except (TypeError, ValueError):
            pass
        year, month = _parse_ym(ym_val)
        if year is None or month is None:
            continue
        for ccy, col_idx in col_map.items():
            try:
                val = row.iloc[col_idx]
                if pd.isna(val):
                    continue
                ttm = float(str(val).replace(",", "").strip())
                if ttm > 0:
                    rates.setdefault((year, month, ccy), ttm)
            except (ValueError, IndexError, TypeError):
                pass
    return rates


def load_rates_from_bytes(data: bytes) -> dict:
    """MURC Excel バイト列からレート辞書を返す → {(year, month, ccy): rate}"""
    primary  = _detect_engine_from_bytes(data)
    engines  = ["xlrd", "openpyxl"] if primary == "xlrd" else ["openpyxl", "xlrd"]

    for engine in engines:
        try:
            xf     = pd.ExcelFile(io.BytesIO(data), engine=engine)
            sheet  = next((s for s in xf.sheet_names if "月毎" in str(s)), xf.sheet_names[0])
            df_raw = pd.read_excel(io.BytesIO(data), sheet_name=sheet, header=None, engine=engine)

            rates = _parse_murc_archive(df_raw)
            if rates:
                return rates

            rates = _parse_murc_all_currencies(df_raw)
            if rates:
                return rates

            usd_raw = _parse_murc_monthly_usd(df_raw)
            if usd_raw:
                return {(y, m, "USD"): r for (y, m), r in usd_raw.items()}

            break
        except Exception:
            continue

    return {}


# ── 為替レートダウンロード・キャッシュ ────────────────────────────────────────
@st.cache_resource
def _rate_store():
    return {"rates": {}, "downloaded": set()}


def _download_kawase_bytes(year: int) -> bytes | None:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36"
        ),
        "Referer": MURC_PAGE_URL,
    }
    url = f"{MURC_XLS_BASE}/murc_{year}.xls"
    try:
        resp = requests.get(url, headers=headers, timeout=30)
        if resp.status_code == 200 and len(resp.content) > 2000:
            return resp.content
    except requests.RequestException:
        pass
    return None


def ensure_rates(years: set, log: list) -> dict:
    store = _rate_store()
    for year in sorted(years):
        if year in store["downloaded"]:
            log.append(f"  {year}年: キャッシュ済み")
            continue
        log.append(f"  {year}年: ダウンロード中...")
        data = _download_kawase_bytes(year)
        if data:
            rates = load_rates_from_bytes(data)
            store["rates"].update(rates)
            store["downloaded"].add(year)
            ccy_count = len({k[2] for k in rates if k[0] == year})
            cnt       = sum(1 for k in rates if k[0] == year)
            log.append(f"  {year}年: {cnt} 件 / {ccy_count} 通貨")
        else:
            log.append(f"  {year}年: ⚠️ 取得失敗（このままレートなしで処理します）")
    return store["rates"]


def get_rate_mc(rates_mc: dict, date_val, currency: str):
    if pd.isna(date_val):
        return None
    return rates_mc.get((date_val.year, date_val.month, currency))


# ── ファイル種類自動検出 ───────────────────────────────────────────────────────
def detect_file_type(data: bytes) -> str:
    try:
        df = read_csv_from_bytes(data)
        cols = set(str(c).strip() for c in df.columns)
    except Exception:
        return "不明"

    if "日付" in cols and "合計" in cols and "手数料" in cols:
        return "PayPal"
    if "Transaction Date" in cols and "Credit Amount" in cols:
        return "Payoneer（旧形式）"
    if "Date" in cols and "Amount" in cols:
        if "Running Balance" in cols:
            return "Wise"
        return "Payoneer（新形式）"
    return "不明"


# ── 為替差損益行挿入 ──────────────────────────────────────────────────────────
def _insert_fx_rows(
    df: pd.DataFrame,
    parse_ym,
    get_balance,
    build_fx_row,
    currency: str,
    rates_mc: dict,
    get_net_amount=None,
) -> pd.DataFrame:
    if currency == "JPY" or df.empty:
        return df

    cols    = list(df.columns)
    records = df.to_dict("records")

    month_last_idx: dict = {}
    month_end_bal:  dict = {}
    for i, row in enumerate(records):
        ym  = parse_ym(row)
        if ym[0] is None:
            continue
        bal = get_balance(row)
        if bal is not None and not pd.isna(bal):
            month_last_idx[ym] = i
            month_end_bal[ym]  = bal

    if not month_last_idx:
        return df

    txn_months = sorted(month_last_idx.keys())

    fx_inserts: dict = {}

    if get_net_amount is not None and records:
        first_row = records[0]
        first_ym  = parse_ym(first_row)
        if first_ym[0] is not None:
            fy, fm    = first_ym
            py, pm    = _prev_ym(fy, fm)
            bal_after = get_balance(first_row)
            net_amt   = get_net_amount(first_row)
            if bal_after is not None and net_amt is not None:
                prev_balance = bal_after - net_amt
                curr_rate = rates_mc.get((py, pm, currency))
                next_rate = rates_mc.get((fy, fm, currency))
                if (curr_rate is not None and next_rate is not None
                        and prev_balance != 0 and not pd.isna(prev_balance)):
                    fx_row = build_fx_row(
                        first_row, py, pm, curr_rate, next_rate, prev_balance, cols
                    )
                    if fx_row is not None:
                        fx_inserts[-1] = [fx_row]

    for i, ym in enumerate(txn_months[:-1]):
        year, month = ym
        ny, nm      = txn_months[i + 1]
        balance     = month_end_bal[ym]
        if balance == 0:
            continue
        curr_rate = rates_mc.get((year, month, currency))
        next_rate = rates_mc.get((ny, nm, currency))
        if curr_rate is None or next_rate is None:
            continue
        last_row       = records[month_last_idx[ym]]
        py_arg, pm_arg = _prev_ym(ny, nm)
        fx_row         = build_fx_row(
            last_row, py_arg, pm_arg, curr_rate, next_rate, balance, cols
        )
        if fx_row is None:
            continue
        fx_inserts.setdefault(month_last_idx[ym], []).append(fx_row)

    result = []
    if -1 in fx_inserts:
        result.extend(fx_inserts[-1])
    for i, row in enumerate(records):
        result.append(row)
        if i in fx_inserts:
            result.extend(fx_inserts[i])

    return pd.DataFrame(result, columns=cols)


# ── PayPal 手数料行分割 ───────────────────────────────────────────────────────
def _split_fee_rows(df: pd.DataFrame) -> pd.DataFrame:
    if "手数料" not in df.columns:
        return df
    result = []
    for _, row in df.iterrows():
        result.append(row)
        fee_raw = row.get("手数料")
        try:
            fee = float(fee_raw) if pd.notna(fee_raw) else 0.0
        except (ValueError, TypeError):
            fee = 0.0
        if fee == 0.0:
            continue
        fee_row            = row.copy()
        fee_row["摘要"]   = "PayPal手数料"
        fee_row["合計"]   = fee
        fee_row["手数料"] = 0
        rate_raw = row.get("レート")
        try:
            rate = float(rate_raw) if pd.notna(rate_raw) else None
        except (ValueError, TypeError):
            rate = None
        fee_row["入出金"] = _jpy_round(fee * rate) if rate is not None else None
        result.append(fee_row)
    return pd.DataFrame(result, columns=df.columns).reset_index(drop=True)


# ── 年検出 ────────────────────────────────────────────────────────────────────
def detect_years_from_files(uploaded_files) -> set:
    years = set()
    for uf in uploaded_files:
        try:
            data = uf.read()
            uf.seek(0)
            df = read_csv_from_bytes(data).head(200)
            for col in df.columns:
                if any(kw in str(col) for kw in ("Date", "日付", "date")):
                    for val in df[col].dropna().astype(str):
                        m = re.search(r"\b(20\d{2})\b", val)
                        if m:
                            years.add(int(m.group(1)))
                    break
        except Exception:
            pass
    if not years:
        today = date.today()
        years = {today.year - 1, today.year}
    return years


# ── 処理関数 ──────────────────────────────────────────────────────────────────
def process_paypal(rates_mc: dict, dfs: list, log: list) -> dict:
    """PayPal CSV → {currency_code: DataFrame}"""
    df = pd.concat(dfs, ignore_index=True)
    df.columns = [str(c).strip() for c in df.columns]

    if "残高への影響" in df.columns:
        df = df[df["残高への影響"].astype(str).str.strip() != "備考"].copy()

    if df.empty:
        log.append("⚠️ PayPal: 処理対象データなし")
        return {}

    df["日付"] = _parse_dates_multi(df["日付"])
    for col in ("合計", "手数料", "正味", "残高"):
        if col in df.columns:
            df[col] = parse_number(df[col])

    name = df.get("名前", pd.Series("", index=df.index)).fillna("").astype(str).str.strip()
    typ  = df.get("タイプ", pd.Series("", index=df.index)).fillna("").astype(str).str.strip()
    df["摘要"] = name.where(name == "", name + " ") + typ

    if "通貨" not in df.columns:
        df["通貨"] = "UNKNOWN"

    results = {}
    for currency in df["通貨"].astype(str).str.strip().unique():
        df_cur = df[df["通貨"].astype(str).str.strip() == currency].copy()
        if df_cur.empty:
            continue

        if currency == "JPY":
            df_cur["レート"] = 1.0
        else:
            df_cur["レート"] = df_cur["日付"].apply(
                lambda d: get_rate_mc(rates_mc, d, currency)
            )

        df_cur["入出金"]     = (df_cur["合計"] * df_cur["レート"]).apply(_jpy_round)
        df_cur["円換算残高"] = (df_cur["残高"] * df_cur["レート"]).apply(_jpy_round)

        out_cols = [
            "日付", "名前", "タイプ", "摘要", "ステータス", "通貨",
            "合計", "手数料", "正味", "残高", "送信者メールアドレス", "残高への影響",
            "レート", "入出金", "円換算残高",
        ]
        out_cols = [c for c in out_cols if c in df_cur.columns]

        df_cur = df_cur.sort_values("日付", ascending=True, na_position="last").reset_index(drop=True)
        df_out = df_cur[out_cols].copy()
        df_out["日付"] = df_out["日付"].dt.strftime("%Y/%m/%d")
        df_out = _split_fee_rows(df_out)

        if currency != "JPY":
            def _parse_ym_pp(row, _fmt="%Y/%m/%d"):
                try:
                    d = pd.to_datetime(str(row.get("日付", "")), format=_fmt)
                    return d.year, d.month
                except Exception:
                    return None, None

            def _get_bal_pp(row):
                try:
                    v = row.get("残高")
                    return float(v) if pd.notna(v) else None
                except (TypeError, ValueError):
                    return None

            def _get_amt_pp(row):
                try:
                    v = row.get("正味")
                    return float(v) if pd.notna(v) else None
                except (TypeError, ValueError):
                    return None

            def _build_pp_fx(row, year, month, curr_rate, next_rate, balance, cols,
                              _ccy=currency):
                ny  = year + (1 if month == 12 else 0)
                nm  = 1 if month == 12 else month + 1
                fx  = _jpy_round((next_rate - curr_rate) * balance)
                if fx == 0:
                    return None
                r = {c: "" for c in cols}
                r["日付"]       = f"{ny}/{nm:02d}/01"
                r["摘要"]       = "為替差損益"
                r["通貨"]       = _ccy
                for c in ("合計", "手数料", "正味"):
                    if c in r:
                        r[c] = 0
                r["残高"]       = balance
                r["レート"]     = next_rate
                r["入出金"]     = fx
                r["円換算残高"] = _jpy_round(balance * next_rate)
                return r

            df_out = _insert_fx_rows(
                df_out, _parse_ym_pp, _get_bal_pp, _build_pp_fx, currency, rates_mc,
                get_net_amount=_get_amt_pp,
            )

        for col in ("入出金", "円換算残高"):
            if col in df_out.columns:
                df_out[col] = (
                    pd.to_numeric(df_out[col], errors="coerce").round(0).astype("Int64")
                )

        results[currency] = df_out
        log.append(f"  PayPal {currency}: {len(df_out):,} 行")

    return results


def process_payoneer_new(rates_mc: dict, dfs: list, log: list) -> dict:
    """Payoneer 新形式 CSV → {key: DataFrame}"""
    df = pd.concat(dfs, ignore_index=True)
    df.columns = [str(c).strip() for c in df.columns]

    df["_date"] = pd.to_datetime(df["Date"], format="%d %b, %Y", errors="coerce")
    df["Date"]  = df["_date"].dt.strftime("%Y/%m/%d")
    df["Amount"] = parse_number(df["Amount"])

    df["_orig_idx"] = range(len(df))
    df = (
        df.sort_values(["_date", "_orig_idx"], ascending=[True, False], na_position="last")
        .reset_index(drop=True)
    )
    df.drop(columns=["_orig_idx"], inplace=True)

    currencies = (
        df["Currency"].astype(str).str.strip().unique().tolist()
        if "Currency" in df.columns else ["USD"]
    )

    results = {}
    for currency in currencies:
        df_cur = (
            df[df["Currency"].astype(str).str.strip() == currency].copy()
            if "Currency" in df.columns else df.copy()
        )
        if df_cur.empty:
            continue

        if currency == "JPY":
            df_cur["レート"] = 1.0
        else:
            df_cur["レート"] = df_cur["_date"].apply(
                lambda d: get_rate_mc(rates_mc, d, currency)
            )

        df_cur["入出金"] = (
            pd.to_numeric(df_cur["Amount"] * df_cur["レート"], errors="coerce")
            .round(0).astype("Int64")
        )

        out_cols = ["Date", "Description", "Amount", "Currency", "レート", "入出金"]
        out_cols = [c for c in out_cols if c in df_cur.columns]
        df_out   = df_cur[out_cols].copy()

        key = f"Payoneer_{currency}" if len(currencies) > 1 else "Payoneer"
        results[key] = df_out
        log.append(f"  Payoneer（新形式）{currency}: {len(df_out):,} 行")

    return results


def process_payoneer_old(rates_mc: dict, dfs: list, log: list) -> dict:
    """Payoneer 旧形式 CSV → {key: DataFrame}"""
    df = pd.concat(dfs, ignore_index=True)
    df.columns = [str(c).strip() for c in df.columns]

    df["_date"] = pd.to_datetime(df["Transaction Date"], format="%m/%d/%Y", errors="coerce")
    df["Transaction Date"] = df["_date"].dt.strftime("%Y年%m月%d日")

    for col in ("Credit Amount", "Debit Amount", "Running Balance"):
        if col in df.columns:
            df[col] = parse_number(df[col]).fillna(0.0)

    df["_orig_idx"] = range(len(df))
    df = (
        df.sort_values(["_date", "_orig_idx"], ascending=[True, False], na_position="last")
        .reset_index(drop=True)
    )
    df.drop(columns=["_orig_idx"], inplace=True)

    df["残高"]       = df["Running Balance"]
    df["レート"]     = df["_date"].apply(lambda d: get_rate_mc(rates_mc, d, "USD"))
    df["入出金_num"] = (df["Credit Amount"] - df["Debit Amount"]) * df["レート"]
    df["入出金"]     = df["入出金_num"].apply(fmt_jpy)
    df["円残高_num"] = df["残高"] * df["レート"]
    df["円残高"]     = df["円残高_num"].apply(fmt_jpy)

    out_cols = [
        "Transaction Date", "Description",
        "Credit Amount", "Debit Amount",
        "残高", "レート", "入出金", "円残高",
    ]
    out_cols = [c for c in out_cols if c in df.columns]
    df_out   = df[out_cols].copy()

    def _parse_ym_po(row):
        try:
            d = pd.to_datetime(str(row.get("Transaction Date", "")), format="%Y年%m月%d日")
            return d.year, d.month
        except Exception:
            return None, None

    def _get_bal_po(row):
        try:
            v = row.get("残高")
            return float(v) if pd.notna(v) else None
        except (TypeError, ValueError):
            return None

    def _get_amt_po(row):
        try:
            c = float(row.get("Credit Amount", 0) or 0)
            d = float(row.get("Debit Amount",  0) or 0)
            return c - d
        except (TypeError, ValueError):
            return None

    def _build_po_fx(row, year, month, curr_rate, next_rate, balance, cols):
        ny  = year + (1 if month == 12 else 0)
        nm  = 1 if month == 12 else month + 1
        fx  = _jpy_round((next_rate - curr_rate) * balance)
        if fx == 0:
            return None
        r = {c: 0 for c in cols}
        r["Transaction Date"] = f"{ny}年{nm:02d}月01日"
        r["Description"]      = "為替差損益"
        r["Credit Amount"]    = 0
        r["Debit Amount"]     = 0
        r["残高"]             = balance
        r["レート"]           = next_rate
        r["入出金"]           = fmt_jpy(fx)
        r["円残高"]           = fmt_jpy(_jpy_round(balance * next_rate))
        return r

    df_out = _insert_fx_rows(
        df_out, _parse_ym_po, _get_bal_po, _build_po_fx, "USD", rates_mc,
        get_net_amount=_get_amt_po,
    )

    log.append(f"  Payoneer（旧形式）: {len(df_out):,} 行")
    return {"Payoneer": df_out}


def process_wise(rates_mc: dict, dfs: list, log: list) -> dict:
    """Wise CSV → {key: DataFrame}"""
    df = pd.concat(dfs, ignore_index=True)
    df.columns = [str(c).strip() for c in df.columns]

    df["_date"] = pd.to_datetime(df["Date"], format="%d-%m-%Y", errors="coerce")
    mask = df["_date"].isna()
    if mask.any():
        df.loc[mask, "_date"] = pd.to_datetime(
            df.loc[mask, "Date"], format="%Y-%m-%d", errors="coerce"
        )
    df["Date"] = df["_date"].dt.strftime("%Y/%m/%d")

    for col in ("Amount", "Running Balance"):
        if col in df.columns:
            df[col] = parse_number(df[col]).fillna(0.0)

    if "Currency" not in df.columns:
        log.append("⚠️ Wise: 'Currency'列が見つかりません。スキップします。")
        return {}

    currencies = [c for c in df["Currency"].dropna().unique() if str(c).upper() != "JPY"]
    if not currencies:
        log.append("⚠️ Wise: JPY以外の通貨が見つかりません。スキップします。")
        return {}

    results = {}
    for currency in sorted(currencies):
        df_cur = df[df["Currency"] == currency].copy()
        df_cur = df_cur.sort_values("_date", ascending=True, na_position="last").reset_index(drop=True)

        df_cur["レート"] = df_cur["_date"].apply(
            lambda d: get_rate_mc(rates_mc, d, str(currency))
        )
        df_cur["円換算金額"] = (df_cur["Amount"] * df_cur["レート"]).apply(
            lambda x: fmt_jpy(x) if pd.notna(x) else ""
        )

        if "Running Balance" in df_cur.columns:
            df_cur["円換算残高"] = df_cur.apply(
                lambda row: fmt_jpy(_jpy_round(row["Running Balance"] * row["レート"]))
                if pd.notna(row["レート"]) else "",
                axis=1,
            )
        else:
            df_cur["円換算残高"] = ""

        if "Description" not in df_cur.columns:
            df_cur["Description"] = ""

        out_cols = ["Date", "Description", "Amount"]
        if "Running Balance" in df_cur.columns:
            out_cols.append("Running Balance")
        out_cols += ["レート", "円換算金額", "円換算残高"]
        out_cols = [c for c in out_cols if c in df_cur.columns]
        df_out   = df_cur[out_cols].copy()

        if "Running Balance" in df_cur.columns:
            def _parse_ym_wise(row):
                try:
                    d = pd.to_datetime(str(row.get("Date", "")), format="%Y/%m/%d")
                    return d.year, d.month
                except Exception:
                    return None, None

            def _get_bal_wise(row):
                try:
                    v = row.get("Running Balance")
                    return float(v) if pd.notna(v) else None
                except (TypeError, ValueError):
                    return None

            def _get_amt_wise(row):
                try:
                    v = row.get("Amount")
                    return float(v) if pd.notna(v) else None
                except (TypeError, ValueError):
                    return None

            def _build_wise_fx(row, year, month, curr_rate, next_rate, balance, cols):
                ny  = year + (1 if month == 12 else 0)
                nm  = 1 if month == 12 else month + 1
                fx  = _jpy_round((next_rate - curr_rate) * balance)
                if fx == 0:
                    return None
                r = {c: 0 for c in cols}
                r["Date"]        = f"{ny}/{nm:02d}/01"
                r["Description"] = "為替差損益"
                r["Amount"]      = 0
                if "Running Balance" in cols:
                    r["Running Balance"] = balance
                r["レート"]      = next_rate
                r["円換算金額"]  = fmt_jpy(fx)
                r["円換算残高"]  = fmt_jpy(_jpy_round(balance * next_rate))
                return r

            df_out = _insert_fx_rows(
                df_out, _parse_ym_wise, _get_bal_wise, _build_wise_fx,
                str(currency), rates_mc, get_net_amount=_get_amt_wise,
            )

        results[f"Wise_{currency}"] = df_out
        log.append(f"  Wise {currency}: {len(df_out):,} 行")

    return results


# ── CSV 出力（CP932） ──────────────────────────────────────────────────────────
def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    csv_str = df.to_csv(index=False)
    try:
        return csv_str.encode(OUTPUT_ENC)
    except UnicodeEncodeError:
        return csv_str.encode(OUTPUT_ENC, errors="replace")


# ── Streamlit UI ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="PayPal/Payoneer/Wise 加工ツール", layout="centered")
st.markdown("# PayPal / Payoneer / Wise<br>取引明細加工ツール", unsafe_allow_html=True)

uploaded = st.file_uploader(
    "CSVファイルを選択（複数可・種類は自動判定）",
    type=["csv"],
    accept_multiple_files=True,
)

# ファイルが変わったらセッション結果をリセット
current_names = tuple(sorted(uf.name for uf in uploaded)) if uploaded else ()
if st.session_state.get("_last_files") != current_names:
    st.session_state["_last_files"]  = current_names
    st.session_state["_results"]     = {}
    st.session_state["_log"]         = []

if uploaded:
    # ── 検出結果テーブル ──
    file_info = []
    for uf in uploaded:
        data  = uf.read()
        uf.seek(0)
        ftype = detect_file_type(data)
        file_info.append({"ファイル名": uf.name, "検出種類": ftype})

    st.dataframe(
        pd.DataFrame(file_info),
        hide_index=True,
        use_container_width=True,
    )

    unknown = [r for r in file_info if r["検出種類"] == "不明"]
    if unknown:
        st.warning(
            f"{len(unknown)} 件のファイルの種類を判定できませんでした。"
            "PayPal / Payoneer / Wise の CSV ファイルを選択してください。"
        )

    # ── 処理実行ボタン ──
    if st.button("▶ 処理実行", type="primary", use_container_width=True):
        log: list = []

        for uf in uploaded:
            uf.seek(0)
        years = detect_years_from_files(uploaded)
        for uf in uploaded:
            uf.seek(0)
        log.append(f"対象年: {sorted(years)}")
        log.append("--- 為替レート ---")

        with st.spinner("為替レートを取得中..."):
            rates_mc = ensure_rates(years, log)

        # ファイルを種類別に分類
        paypal_dfs       = []
        payoneer_new_dfs = []
        payoneer_old_dfs = []
        wise_dfs         = []

        for uf, info in zip(uploaded, file_info):
            uf.seek(0)
            try:
                data  = uf.read()
                df    = read_csv_from_bytes(data)
                ftype = info["検出種類"]
                if ftype == "PayPal":
                    paypal_dfs.append(df)
                elif ftype == "Payoneer（新形式）":
                    payoneer_new_dfs.append(df)
                elif ftype == "Payoneer（旧形式）":
                    payoneer_old_dfs.append(df)
                elif ftype == "Wise":
                    wise_dfs.append(df)
            except Exception as e:
                log.append(f"⚠️ {uf.name}: 読み込みエラー - {e}")

        log.append("--- 処理結果 ---")
        all_results: dict = {}

        with st.spinner("処理中..."):
            if paypal_dfs:
                for ccy, df_r in process_paypal(rates_mc, paypal_dfs, log).items():
                    all_results[f"PayPal_{ccy}_Processed"] = df_r
            if payoneer_new_dfs:
                for key, df_r in process_payoneer_new(rates_mc, payoneer_new_dfs, log).items():
                    all_results[f"{key}_Processed"] = df_r
            if payoneer_old_dfs:
                for key, df_r in process_payoneer_old(rates_mc, payoneer_old_dfs, log).items():
                    all_results[f"{key}_Processed"] = df_r
            if wise_dfs:
                for key, df_r in process_wise(rates_mc, wise_dfs, log).items():
                    all_results[f"{key}_Processed"] = df_r

        # 結果を session_state に保存（bytes に変換しておく）
        st.session_state["_results"] = {
            name: df_to_csv_bytes(df) for name, df in all_results.items()
        }
        st.session_state["_log"] = log

    # ── 処理ログ表示 ──
    if st.session_state.get("_log"):
        with st.expander("処理ログ", expanded=True):
            st.text("\n".join(st.session_state["_log"]))

    # ── ダウンロードボタン ──
    if st.session_state.get("_results"):
        st.subheader("ダウンロード")
        for fname, csv_bytes in st.session_state["_results"].items():
            st.download_button(
                label=f"⬇ {fname}.csv",
                data=csv_bytes,
                file_name=f"{fname}.csv",
                mime="application/octet-stream",
                use_container_width=True,
            )
