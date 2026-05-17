
# app.py
# Material Action Dashboard - Revised Logic
# 실행: streamlit run app.py

from __future__ import annotations

import io
import re
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


# =========================================================
# 0. Page Config / CSS
# =========================================================

st.set_page_config(
    page_title="Material Action Dashboard",
    page_icon="📦",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container {
        padding-top: 3.5rem !important;
        padding-bottom: 2rem !important;
    }
    header[data-testid="stHeader"] {
        height: 0px;
    }
    .big-title {
        font-size: 2rem;
        font-weight: 800;
        margin-bottom: 0.5rem;
    }
    .section-sub {
        color: #6b7280;
        font-size: 0.95rem;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.55rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# 1. Utility
# =========================================================

def clean(x: Any) -> str:
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    return str(x).replace("\n", " ").replace("\r", " ").strip()


def norm(x: Any) -> str:
    s = clean(x).lower()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[^0-9a-z가-힣#./+\- ]", "", s)
    return s.strip()


def to_num(x: Any, default: float = 0.0) -> float:
    if x is None:
        return default
    try:
        if pd.isna(x):
            return default
    except Exception:
        pass
    if isinstance(x, (int, float, np.number)):
        return float(x)
    s = str(x).strip().replace(",", "").replace("$", "").replace("%", "")
    if s == "" or s.lower() in ["nan", "none"]:
        return default
    try:
        return float(s)
    except Exception:
        return default


def parse_date(x: Any) -> Optional[pd.Timestamp]:
    if x is None:
        return None
    try:
        if pd.isna(x):
            return None
    except Exception:
        pass
    try:
        dt = pd.to_datetime(x, errors="coerce")
        if pd.isna(dt):
            return None
        return pd.Timestamp(dt).normalize()
    except Exception:
        return None


def fmt_qty(x: Any) -> str:
    n = to_num(x, 0)
    if abs(n - round(n)) < 1e-9:
        return f"{n:,.0f}"
    return f"{n:,.2f}"


def find_sheet(sheet_names: List[str], candidates: List[str]) -> Optional[str]:
    for c in candidates:
        for s in sheet_names:
            if c in s:
                return s
    return None


def risk_label(r: str) -> str:
    if r == "긴급":
        return "🔴 긴급"
    if r == "주의":
        return "🟠 주의"
    if r == "정상":
        return "🟢 정상"
    return "⚪ 확인 필요"


def is_material_ok(status: Any) -> bool:
    s = norm(status)
    # 생산계획 시트의 MATERIAL 열이 OK면, 회사 기준으로 해당 생산건 자재 준비 완료로 우선 판단
    return s == "ok" or s.startswith("ok ")


def is_material_risky(status: Any, remark: Any = "") -> bool:
    s = norm(status + " " + remark)
    risky_words = [
        "tba", "eta", "delay", "delayed", "waiting", "lack", "short",
        "부족", "지연", "미입고", "대기", "확인", "lining eta", "bc tba"
    ]
    return any(w in s for w in risky_words)


def item_key(x: Any) -> str:
    # 자재명 매칭을 위한 최소 정규화
    s = norm(x)
    s = re.sub(r"\([^)]*\)", "", s)  # 괄호 내용 제거한 보조키
    s = re.sub(r"\s+", " ", s).strip()
    return s


def item_key_strict(x: Any) -> str:
    return norm(x)


def parse_leadtime_days(x: Any, default: int = 30) -> int:
    s = norm(x)
    nums = re.findall(r"\d+", s)
    if not nums:
        return default
    n = int(nums[0])
    if "week" in s or "weeks" in s or "주" in s:
        return n * 7
    return n


# =========================================================
# 2. Load Excel
# =========================================================

@st.cache_data(show_spinner=False)
def load_excel(file_bytes: bytes) -> Tuple[List[str], Dict[str, pd.DataFrame]]:
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets = {}
    for s in xls.sheet_names:
        try:
            sheets[s] = pd.read_excel(io.BytesIO(file_bytes), sheet_name=s, header=None)
        except Exception:
            sheets[s] = pd.DataFrame()
    return xls.sheet_names, sheets


# =========================================================
# 3. Parsers
# =========================================================

def parse_production(raw: pd.DataFrame) -> pd.DataFrame:
    cols = [
        "prod_key", "source_row", "buyer", "prod_year", "prod_month", "ship_month",
        "pr_no", "buyer_po", "item_code", "style", "color", "prod_qty",
        "ref_qty", "sewing_from", "sewing_to", "ship_date", "completed_qty",
        "balance", "completed_pct", "material_status", "remark"
    ]

    if raw is None or raw.empty or raw.shape[1] < 12:
        return pd.DataFrame(columns=cols)

    rows = []
    for i in range(0, len(raw)):
        r = raw.iloc[i]

        buyer = clean(r.iloc[1] if len(r) > 1 else "")
        buyer_po = clean(r.iloc[6] if len(r) > 6 else "")
        style = clean(r.iloc[8] if len(r) > 8 else "")
        color = clean(r.iloc[9] if len(r) > 9 else "")
        prod_qty = to_num(r.iloc[10] if len(r) > 10 else 0)

        if not buyer or not style or prod_qty <= 0:
            continue
        if norm(buyer) in ["buyer", "ppm"] or norm(style) in ["style"]:
            continue
        if "total" in norm(style) or "city" in norm(style):
            continue

        sewing_from = parse_date(r.iloc[18] if len(r) > 18 else None)
        sewing_to = parse_date(r.iloc[19] if len(r) > 19 else None)
        ship_date = parse_date(r.iloc[20] if len(r) > 20 else None)

        if sewing_from is None and ship_date is not None:
            sewing_from = ship_date

        completed_qty = to_num(r.iloc[21] if len(r) > 21 else 0)
        balance = to_num(r.iloc[22] if len(r) > 22 else 0)
        completed_pct = to_num(r.iloc[23] if len(r) > 23 else 0)

        prod_key = f"{buyer_po}|{style}|{color}|{prod_qty}|row{i+1}"

        rows.append({
            "prod_key": prod_key,
            "source_row": i + 1,
            "buyer": buyer,
            "prod_year": clean(r.iloc[2] if len(r) > 2 else ""),
            "prod_month": clean(r.iloc[3] if len(r) > 3 else ""),
            "ship_month": clean(r.iloc[4] if len(r) > 4 else ""),
            "pr_no": clean(r.iloc[5] if len(r) > 5 else ""),
            "buyer_po": buyer_po,
            "item_code": clean(r.iloc[7] if len(r) > 7 else ""),
            "style": style,
            "color": color,
            "prod_qty": prod_qty,
            "ref_qty": to_num(r.iloc[11] if len(r) > 11 else 0),
            "sewing_from": sewing_from,
            "sewing_to": sewing_to,
            "ship_date": ship_date,
            "completed_qty": completed_qty,
            "balance": balance,
            "completed_pct": completed_pct,
            "material_status": clean(r.iloc[24] if len(r) > 24 else ""),
            "remark": clean(r.iloc[25] if len(r) > 25 else ""),
        })

    return pd.DataFrame(rows, columns=cols)


def parse_bom(raw: pd.DataFrame) -> pd.DataFrame:
    cols = [
        "source_row", "style", "style_color", "style_item_code", "supplier",
        "lead_time", "leadtime_days", "item", "item_key", "item_key_strict",
        "spec", "material_color", "unit", "buyer_price", "aj_price",
        "moq", "usage_per_unit"
    ]

    if raw is None or raw.empty or raw.shape[0] < 13 or raw.shape[1] < 19:
        return pd.DataFrame(columns=cols)

    style_cols = []
    max_col = min(raw.shape[1], 28)
    for c in range(18, max_col):
        style = clean(raw.iat[1, c] if raw.shape[0] > 1 else "")
        item_code = clean(raw.iat[2, c] if raw.shape[0] > 2 else "")
        style_color = clean(raw.iat[3, c] if raw.shape[0] > 3 else "")
        if style:
            style_cols.append((c, style, item_code, style_color))

    rows = []
    for i in range(12, len(raw)):
        r = raw.iloc[i]
        item = clean(r.iloc[7] if len(r) > 7 else "")
        if not item:
            continue

        supplier = clean(r.iloc[4] if len(r) > 4 else "")
        lead_time = clean(r.iloc[6] if len(r) > 6 else "")
        lead_days = parse_leadtime_days(lead_time)

        for c, style, item_code, style_color in style_cols:
            usage = to_num(r.iloc[c] if len(r) > c else 0)
            if usage <= 0:
                continue
            rows.append({
                "source_row": i + 1,
                "style": style,
                "style_color": style_color,
                "style_item_code": item_code,
                "supplier": supplier,
                "lead_time": lead_time,
                "leadtime_days": lead_days,
                "item": item,
                "item_key": item_key(item),
                "item_key_strict": item_key_strict(item),
                "spec": clean(r.iloc[8] if len(r) > 8 else ""),
                "material_color": clean(r.iloc[9] if len(r) > 9 else ""),
                "unit": clean(r.iloc[11] if len(r) > 11 else ""),
                "buyer_price": to_num(r.iloc[12] if len(r) > 12 else 0),
                "aj_price": to_num(r.iloc[13] if len(r) > 13 else 0),
                "moq": to_num(r.iloc[15] if len(r) > 15 else 0),
                "usage_per_unit": usage,
            })

    return pd.DataFrame(rows, columns=cols)


def parse_inventory(raw: pd.DataFrame) -> pd.DataFrame:
    cols = ["source_row", "supplier", "item", "item_key", "item_key_strict", "material_color", "unit", "stock_qty", "unit_price"]

    if raw is None or raw.empty or raw.shape[1] < 14:
        return pd.DataFrame(columns=cols)

    rows = []
    for i in range(0, len(raw)):
        r = raw.iloc[i]
        item = clean(r.iloc[7] if len(r) > 7 else "")
        stock_qty = to_num(r.iloc[13] if len(r) > 13 else 0)
        if not item or stock_qty == 0:
            continue
        rows.append({
            "source_row": i + 1,
            "supplier": clean(r.iloc[6] if len(r) > 6 else ""),
            "item": item,
            "item_key": item_key(item),
            "item_key_strict": item_key_strict(item),
            "material_color": clean(r.iloc[9] if len(r) > 9 else ""),
            "unit": clean(r.iloc[11] if len(r) > 11 else ""),
            "stock_qty": stock_qty,
            "unit_price": to_num(r.iloc[12] if len(r) > 12 else 0),
        })

    return pd.DataFrame(rows, columns=cols)


def parse_flow(raw: pd.DataFrame, kind: str) -> pd.DataFrame:
    """
    입고/출고 시트는 구조가 상대적으로 불안정하므로,
    자재명 후보 + 수량 후보를 넓게 읽는다.
    """
    cols = ["kind", "source_row", "buyer", "pr_no", "po", "supplier", "item", "item_key", "item_key_strict", "color", "unit", "qty"]

    if raw is None or raw.empty:
        return pd.DataFrame(columns=cols)

    rows = []
    for i in range(0, len(raw)):
        r = raw.iloc[i]
        vals = [clean(v) for v in r.tolist()]
        lower = [norm(v) for v in vals]

        # 헤더/빈 행 제외
        if not any(vals):
            continue
        if any(h in " ".join(lower) for h in ["item", "description", "qty"]) and i < 10:
            continue

        item = ""
        # 입고/출고 시트에서 자재명은 보통 중간 열에 있음.
        # 명확한 자재 키워드가 있는 가장 긴 텍스트를 우선 선택.
        candidates = []
        keywords = [
            "ny", "poly", "zipper", "web", "fabric", "chain", "non woven",
            "elastic", "slider", "foam", "label", "laser", "kraft", "plain",
            "clay", "magic", "bar", "bio", "bag", "lining"
        ]
        for v in vals:
            nv = norm(v)
            if len(v) >= 4 and not re.fullmatch(r"[0-9.,/$\-]+", v):
                score = sum(1 for k in keywords if k in nv)
                if score > 0:
                    candidates.append((score, len(v), v))
        if candidates:
            candidates.sort(reverse=True)
            item = candidates[0][2]

        if not item:
            continue

        nums = [to_num(v, np.nan) for v in vals]
        nums = [x for x in nums if not pd.isna(x) and x > 0]
        qty = nums[-1] if nums else 0
        if qty <= 0:
            continue

        # buyer/pr/po/supplier는 보조 정보로만 사용
        rows.append({
            "kind": kind,
            "source_row": i + 1,
            "buyer": vals[1] if len(vals) > 1 else "",
            "pr_no": "",
            "po": vals[2] if len(vals) > 2 else "",
            "supplier": vals[4] if len(vals) > 4 else "",
            "item": item,
            "item_key": item_key(item),
            "item_key_strict": item_key_strict(item),
            "color": "",
            "unit": "",
            "qty": qty,
        })

    return pd.DataFrame(rows, columns=cols)


# =========================================================
# 4. Calculation
# =========================================================

def make_availability(inventory: pd.DataFrame, inbound: pd.DataFrame, outbound: pd.DataFrame) -> pd.DataFrame:
    """
    재고 + 입고 - 출고를 참고용 가용량으로 계산.
    단, 회사 시트에서는 출고가 이미 생산 투입을 의미할 수 있으므로
    실제 부족 판단은 MATERIAL 상태를 우선하고,
    이 값은 '보조 근거'로 사용한다.
    """
    keys = set()
    for df in [inventory, inbound, outbound]:
        if df is not None and not df.empty and "item_key" in df.columns:
            keys.update(df["item_key"].dropna().tolist())

    rows = []
    for k in keys:
        stock = inventory.loc[inventory["item_key"] == k, "stock_qty"].sum() if not inventory.empty else 0
        in_qty = inbound.loc[inbound["item_key"] == k, "qty"].sum() if not inbound.empty else 0
        out_qty = outbound.loc[outbound["item_key"] == k, "qty"].sum() if not outbound.empty else 0
        rows.append({
            "item_key": k,
            "stock_qty": float(stock),
            "inbound_qty": float(in_qty),
            "outbound_qty": float(out_qty),
            "available_reference_qty": float(stock + in_qty - out_qty),
        })

    return pd.DataFrame(rows, columns=["item_key", "stock_qty", "inbound_qty", "outbound_qty", "available_reference_qty"])


def filter_production(prod: pd.DataFrame, include_ok: bool, include_completed: bool) -> pd.DataFrame:
    if prod.empty:
        return prod.copy()

    out = prod.copy()

    if not include_completed:
        # 완료율 100%, balance 0, completed_qty >= prod_qty 등은 분석에서 제외
        completed_mask = (
            (out["completed_pct"] >= 100)
            | ((out["completed_qty"] >= out["prod_qty"]) & (out["prod_qty"] > 0))
            | ((out["balance"] == 0) & (out["completed_qty"] > 0))
        )
        out = out[~completed_mask].copy()

    if not include_ok:
        # MATERIAL OK는 회사 기준 자재 준비 완료로 보고 부족 산출 대상에서 제외
        out = out[~out["material_status"].apply(is_material_ok)].copy()

    return out


def build_analysis(
    prod: pd.DataFrame,
    bom: pd.DataFrame,
    inventory: pd.DataFrame,
    inbound: pd.DataFrame,
    outbound: pd.DataFrame,
    loss_rate: float,
    include_ok: bool = False,
    include_completed: bool = False,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    detail_cols = [
        "prod_key", "buyer_po", "buyer", "style", "style_color", "prod_qty",
        "need_date", "sewing_to", "ship_date", "material_status", "remark",
        "supplier", "item", "material_color", "unit", "usage_per_unit",
        "required_qty", "required_with_loss", "stock_qty", "inbound_qty",
        "outbound_qty", "available_reference_qty", "shortage_qty",
        "moq", "recommended_order_qty", "leadtime_days", "risk", "action",
        "source_row_prod", "source_row_bom", "basis"
    ]
    summary_cols = [
        "prod_key", "buyer_po", "buyer", "style", "style_color", "prod_qty",
        "need_date", "sewing_to", "ship_date", "material_status", "remark",
        "shortage_items", "urgent_items", "min_days_left", "risk", "main_action"
    ]

    prod_target = filter_production(prod, include_ok=include_ok, include_completed=include_completed)
    availability = make_availability(inventory, inbound, outbound)

    if prod_target.empty:
        return pd.DataFrame(columns=detail_cols), pd.DataFrame(columns=summary_cols), availability

    if bom.empty:
        summary = prod_target.copy()
        summary["shortage_items"] = 0
        summary["urgent_items"] = 0
        summary["min_days_left"] = np.nan
        summary["risk"] = "확인필요"
        summary["main_action"] = "BOM 데이터 매칭 필요"
        return pd.DataFrame(columns=detail_cols), summary[summary_cols], availability

    rows = []
    today = pd.Timestamp.today().normalize()

    for _, p in prod_target.iterrows():
        p_style = norm(p["style"])
        p_color = norm(p["color"])

        matched = bom[(bom["style"].apply(norm) == p_style) & (bom["style_color"].apply(norm) == p_color)].copy()
        if matched.empty:
            matched = bom[bom["style"].apply(norm) == p_style].copy()

        if matched.empty:
            rows.append({
                "prod_key": p["prod_key"],
                "buyer_po": p["buyer_po"],
                "buyer": p["buyer"],
                "style": p["style"],
                "style_color": p["color"],
                "prod_qty": p["prod_qty"],
                "need_date": p["sewing_from"],
                "sewing_to": p["sewing_to"],
                "ship_date": p["ship_date"],
                "material_status": p["material_status"],
                "remark": p["remark"],
                "supplier": "",
                "item": "(BOM 매칭 실패)",
                "material_color": "",
                "unit": "",
                "usage_per_unit": 0,
                "required_qty": 0,
                "required_with_loss": 0,
                "stock_qty": 0,
                "inbound_qty": 0,
                "outbound_qty": 0,
                "available_reference_qty": 0,
                "shortage_qty": 0,
                "moq": 0,
                "recommended_order_qty": 0,
                "leadtime_days": 0,
                "days_left": None,
                "risk": "확인필요",
                "urgency_reason": "BOM 매칭 실패",
                "action": "BOM에서 해당 STYLE/COLOR 매칭 필요",
                "source_row_prod": p["source_row"],
                "source_row_bom": "",
                "basis": "BOM 매칭 실패",
            })
            continue

        for _, b in matched.iterrows():
            required = float(p["prod_qty"]) * float(b["usage_per_unit"])
            required_loss = required * (1 + loss_rate)

            av = availability[availability["item_key"] == b["item_key"]]
            if av.empty:
                stock_qty = 0.0
                inbound_qty = 0.0
                outbound_qty = 0.0
                available_ref = 0.0
            else:
                stock_qty = float(av["stock_qty"].iloc[0])
                inbound_qty = float(av["inbound_qty"].iloc[0])
                outbound_qty = float(av["outbound_qty"].iloc[0])
                available_ref = float(av["available_reference_qty"].iloc[0])

            # 긴급 기준 수정:
            # 1순위: 부족량 존재 + 생산일까지 남은 일수 <= 리드타임
            # 2순위: MATERIAL/REMARK에 TBA, ETA, 지연 등 확인 필요 문구 존재
            # 주의: 부족은 있으나 리드타임상 조달 가능
            # 정상: MATERIAL=OK 또는 부족 없음
            need_date = p["sewing_from"]
            days_left = None
            if need_date is not None and not pd.isna(need_date):
                days_left = int((pd.Timestamp(need_date).normalize() - today).days)

            lead = int(b["leadtime_days"] or 30)
            risky_by_status = is_material_risky(p["material_status"], p["remark"])

            if is_material_ok(p["material_status"]):
                shortage = 0.0
                rec = 0.0
                risk = "정상"
                urgency_reason = "MATERIAL=OK"
                basis = "생산계획 MATERIAL=OK 기준"
                action = "생산계획상 MATERIAL=OK로 표시되어 추가 발주 불필요"
            else:
                shortage = max(required_loss - available_ref, 0)
                moq = float(b["moq"] or 0)
                rec = 0 if shortage <= 0 else max(shortage, moq)

                if shortage > 0 and days_left is not None and days_left <= lead:
                    risk = "긴급"
                    urgency_reason = f"생산일까지 {days_left}일 남음 / 리드타임 {lead}일 / 부족 {fmt_qty(shortage)}{b['unit']}"
                    basis = "부족량 존재 + 리드타임 내 조달 불가"
                    action = (
                        f"{b['supplier'] or '공급업체 확인'}에 {b['item']} "
                        f"{fmt_qty(rec)}{b['unit']}를 생산 전까지 입고 가능한지 즉시 확인 "
                        f"({urgency_reason})"
                    )
                elif risky_by_status:
                    risk = "긴급"
                    urgency_reason = f"MATERIAL/REMARK 확인 필요: {clean(p['material_status']) or clean(p['remark'])}"
                    basis = "MATERIAL/REMARK 위험 표시"
                    if rec > 0:
                        action = (
                            f"{b['supplier'] or '공급업체 확인'}에 {b['item']} "
                            f"{fmt_qty(rec)}{b['unit']} 발주/입고 가능 여부 즉시 확인 "
                            f"({urgency_reason})"
                        )
                    else:
                        action = (
                            f"{b['supplier'] or '공급업체 확인'}에 {b['item']} "
                            f"입고/ETA 상태 즉시 확인 ({urgency_reason})"
                        )
                elif shortage > 0:
                    risk = "주의"
                    urgency_reason = f"부족 {fmt_qty(shortage)}{b['unit']} / 생산일까지 {days_left if days_left is not None else '미확인'}일 / 리드타임 {lead}일"
                    basis = "부족은 있으나 리드타임상 검토 가능"
                    action = (
                        f"{b['supplier'] or '공급업체 확인'}에 {b['item']} "
                        f"{fmt_qty(rec)}{b['unit']} 발주 검토 ({urgency_reason})"
                    )
                else:
                    risk = "정상"
                    urgency_reason = "참고 가용량 기준 부족 없음"
                    basis = "참고 가용량 기준 커버 가능"
                    action = "재고/입고 흐름 기준으로 커버 가능"

            rows.append({
                "prod_key": p["prod_key"],
                "buyer_po": p["buyer_po"],
                "buyer": p["buyer"],
                "style": p["style"],
                "style_color": p["color"],
                "prod_qty": p["prod_qty"],
                "need_date": p["sewing_from"],
                "sewing_to": p["sewing_to"],
                "ship_date": p["ship_date"],
                "material_status": p["material_status"],
                "remark": p["remark"],
                "supplier": b["supplier"],
                "item": b["item"],
                "material_color": b["material_color"],
                "unit": b["unit"],
                "usage_per_unit": b["usage_per_unit"],
                "required_qty": required,
                "required_with_loss": required_loss,
                "stock_qty": stock_qty,
                "inbound_qty": inbound_qty,
                "outbound_qty": outbound_qty,
                "available_reference_qty": available_ref,
                "shortage_qty": shortage,
                "moq": float(b["moq"] or 0),
                "recommended_order_qty": rec,
                "leadtime_days": int(b["leadtime_days"] or 30),
                "days_left": days_left,
                "risk": risk,
                "urgency_reason": urgency_reason,
                "action": action,
                "source_row_prod": p["source_row"],
                "source_row_bom": b["source_row"],
                "basis": basis,
            })

    detail = pd.DataFrame(rows, columns=detail_cols)

    if detail.empty:
        summary = prod_target.copy()
        summary["shortage_items"] = 0
        summary["urgent_items"] = 0
        summary["min_days_left"] = np.nan
        summary["risk"] = "확인필요"
        summary["main_action"] = "분석 가능한 자재 없음"
        return detail, summary[summary_cols], availability

    summary = detail.groupby("prod_key", as_index=False).agg(
        buyer_po=("buyer_po", "first"),
        buyer=("buyer", "first"),
        style=("style", "first"),
        style_color=("style_color", "first"),
        prod_qty=("prod_qty", "first"),
        need_date=("need_date", "first"),
        sewing_to=("sewing_to", "first"),
        ship_date=("ship_date", "first"),
        material_status=("material_status", "first"),
        remark=("remark", "first"),
        shortage_items=("shortage_qty", lambda x: int((x > 0).sum())),
        urgent_items=("risk", lambda x: int((x == "긴급").sum())),
        min_days_left=("days_left", lambda x: pd.to_numeric(x, errors="coerce").min()),
    )

    def summary_risk(r):
        if is_material_ok(r["material_status"]):
            return "정상"
        if r["urgent_items"] > 0:
            return "긴급"
        if r["shortage_items"] > 0:
            return "주의"
        if is_material_risky(r["material_status"], r["remark"]):
            return "긴급"
        return "정상"

    summary["risk"] = summary.apply(summary_risk, axis=1)

    actions = []
    for key, g in detail.groupby("prod_key"):
        status = clean(g["material_status"].iloc[0]) if "material_status" in g.columns and len(g) else ""
        remark = clean(g["remark"].iloc[0]) if "remark" in g.columns and len(g) else ""

        if is_material_ok(status):
            actions.append({"prod_key": key, "main_action": "MATERIAL=OK: 추가 발주 불필요"})
            continue

        short = g[(g["shortage_qty"] > 0) | (g["risk"] == "긴급")].copy()
        if short.empty:
            actions.append({"prod_key": key, "main_action": "현재 기준 특이 부족 없음"})
        else:
            # 대표 액션 우선순위:
            # 1) 부족량 존재 + 생산일까지 남은 일수 <= 리드타임
            # 2) MATERIAL/REMARK 위험 표시
            # 3) 부족량 큰 항목
            short["days_left_num"] = pd.to_numeric(short["days_left"], errors="coerce")
            short["leadtime_num"] = pd.to_numeric(short["leadtime_days"], errors="coerce")
            short["is_leadtime_short"] = (
                (short["shortage_qty"] > 0)
                & short["days_left_num"].notna()
                & (short["days_left_num"] <= short["leadtime_num"])
            )
            short["priority"] = np.select(
                [
                    short["is_leadtime_short"],
                    short["risk"].eq("긴급"),
                    short["shortage_qty"].gt(0),
                ],
                [0, 1, 2],
                default=9,
            )
            top = short.sort_values(
                ["priority", "days_left_num", "shortage_qty"],
                ascending=[True, True, False],
                na_position="last",
            ).iloc[0]
            actions.append({"prod_key": key, "main_action": top["action"]})

    summary = summary.merge(pd.DataFrame(actions), on="prod_key", how="left")
    summary["main_action"] = summary["main_action"].fillna("확인 필요")
    rank2 = {"긴급": 0, "주의": 1, "정상": 2, "확인필요": 3}
    summary["rank"] = summary["risk"].map(rank2).fillna(9)
    # 긴급 안에서는 생산일까지 남은 기간이 짧은 건, 부족/긴급 자재가 많은 건을 우선 표시
    summary = summary.sort_values(
        ["rank", "min_days_left", "urgent_items", "shortage_items"],
        ascending=[True, True, False, False],
        na_position="last",
    )

    return detail, summary[summary_cols], availability


# =========================================================
# 5. UI
# =========================================================

def show_intro():
    st.markdown('<div class="big-title">📦 Material Action Dashboard</div>', unsafe_allow_html=True)
    st.info("왼쪽 사이드바에서 엑셀 파일을 업로드하면 분석이 시작됩니다.")
    st.markdown(
        """
        ### 이 앱이 보여주는 것
        - 생산계획상 MATERIAL 상태를 우선 반영한 자재 리스크
        - 어떤 PO가 자재 확인/발주 검토 대상인지
        - 어떤 자재를 어느 공급업체에 언제까지 확인해야 하는지
        - 상세 화면에서 BOM/재고/입고/출고 원본 데이터 확인
        """
    )


def show_dashboard(summary: pd.DataFrame, detail: pd.DataFrame):
    st.markdown('<div class="big-title">📦 Material Action Dashboard</div>', unsafe_allow_html=True)

    if summary.empty:
        st.warning("분석 가능한 생산계획 데이터가 없습니다. 데이터 점검 탭에서 시트 로드 상태를 확인해 주세요.")
        return

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("분석 대상 생산계획", f"{len(summary):,}")
    c2.metric("긴급 PO", f"{(summary['risk'] == '긴급').sum():,}")
    c3.metric("주의 PO", f"{(summary['risk'] == '주의').sum():,}")
    c4.metric("부족/확인 자재 항목", f"{((detail['shortage_qty'] > 0) | (detail['risk'] == '긴급')).sum() if not detail.empty else 0:,}")

    left, right = st.columns([0.9, 1.1])

    with left:
        st.subheader("위험도 분포")
        rc = summary["risk"].value_counts().reset_index()
        rc.columns = ["risk", "count"]
        fig = px.pie(
            rc,
            names="risk",
            values="count",
            hole=0.45,
            color="risk",
            color_discrete_map={"긴급": "#ef4444", "주의": "#f59e0b", "정상": "#10b981", "확인필요": "#9ca3af"},
        )
        fig.update_layout(height=330, margin=dict(l=10, r=10, t=20, b=10))
        st.plotly_chart(fig, use_container_width=True)

    with right:
        st.subheader("확인 필요 자재 TOP")
        short = detail[(detail["shortage_qty"] > 0) | (detail["risk"] == "긴급")].copy() if not detail.empty else pd.DataFrame()
        if short.empty:
            st.success("현재 표시 기준 확인 필요 자재가 없습니다.")
        else:
            top = short.groupby(["item", "unit"], as_index=False).agg(
                shortage_qty=("shortage_qty", "sum"),
                urgent_count=("risk", lambda x: int((x == "긴급").sum())),
            )
            top = top.sort_values(["urgent_count", "shortage_qty"], ascending=[False, False]).head(10)
            fig = px.bar(top, x="shortage_qty", y="item", orientation="h", text="shortage_qty")
            fig.update_traces(texttemplate="%{text:.1f}", textposition="outside")
            fig.update_layout(height=330, margin=dict(l=10, r=10, t=20, b=10), yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)

    st.subheader("PO별 요약")
    view = summary.copy()
    view["상태"] = view["risk"].apply(risk_label)
    view["생산시작"] = pd.to_datetime(view["need_date"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("-")
    view["생산종료"] = pd.to_datetime(view["sewing_to"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("-")
    st.dataframe(
        view[[
            "상태", "buyer_po", "style", "style_color", "prod_qty",
            "생산시작", "생산종료", "material_status", "min_days_left", "shortage_items", "urgent_items", "main_action"
        ]].rename(columns={
            "buyer_po": "PO No.",
            "style": "Style",
            "style_color": "Color",
            "prod_qty": "생산수량",
            "material_status": "MATERIAL",
            "min_days_left": "생산일까지 남은 일수",
            "shortage_items": "부족 자재 수",
            "urgent_items": "긴급 자재 수",
            "main_action": "우선 액션",
        }),
        use_container_width=True,
        hide_index=True,
    )


def show_detail(summary: pd.DataFrame, detail: pd.DataFrame, parsed: Dict[str, pd.DataFrame]):
    st.header("PO 상세 조회")

    if summary.empty:
        st.warning("상세 조회 가능한 PO가 없습니다.")
        return

    opts = summary.copy()
    opts["label"] = (
        opts["risk"].apply(risk_label)
        + " | "
        + opts["buyer_po"].astype(str)
        + " | "
        + opts["style"].astype(str)
        + " / "
        + opts["style_color"].astype(str)
        + " | "
        + opts["prod_qty"].apply(lambda x: f"{x:,.0f}PCS")
    )

    selected = st.selectbox("PO 선택", opts["label"].tolist())
    key = opts.loc[opts["label"] == selected, "prod_key"].iloc[0]
    srow = summary[summary["prod_key"] == key].iloc[0]
    d = detail[detail["prod_key"] == key].copy()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("상태", risk_label(srow["risk"]))
    c2.metric("생산수량", f"{srow['prod_qty']:,.0f} PCS")
    c3.metric("부족 자재", f"{int(srow['shortage_items']):,}")
    c4.metric("긴급 자재", f"{int(srow['urgent_items']):,}")

    st.subheader("1. 생산 정보")
    info = pd.DataFrame([
        ["PO No.", srow["buyer_po"]],
        ["Style / Color", f"{srow['style']} / {srow['style_color']}"],
        ["생산수량", f"{srow['prod_qty']:,.0f} PCS"],
        ["생산시작", pd.to_datetime(srow["need_date"]).strftime("%Y-%m-%d") if pd.notna(srow["need_date"]) else "-"],
        ["생산종료", pd.to_datetime(srow["sewing_to"]).strftime("%Y-%m-%d") if pd.notna(srow["sewing_to"]) else "-"],
        ["선적일", pd.to_datetime(srow["ship_date"]).strftime("%Y-%m-%d") if pd.notna(srow["ship_date"]) else "-"],
        ["MATERIAL", srow["material_status"]],
        ["REMARK", srow["remark"]],
    ], columns=["항목", "내용"])
    st.dataframe(info, use_container_width=True, hide_index=True)

    st.subheader("2. 자재 필요량 · 입출고 참고 · 발주 액션")
    if d.empty:
        st.info("이 PO는 BOM 매칭 결과가 없습니다.")
    else:
        dv = d.copy()
        dv["상태"] = dv["risk"].apply(risk_label)
        st.dataframe(
            dv[[
                "상태", "supplier", "item", "material_color", "unit",
                "usage_per_unit", "prod_qty", "required_with_loss",
                "stock_qty", "inbound_qty", "outbound_qty", "available_reference_qty",
                "shortage_qty", "moq", "recommended_order_qty",
                "leadtime_days", "days_left", "urgency_reason", "basis", "action"
            ]].rename(columns={
                "supplier": "Supplier",
                "item": "자재명",
                "material_color": "자재 Color",
                "unit": "단위",
                "usage_per_unit": "1개당 소요량",
                "prod_qty": "생산수량",
                "required_with_loss": "Loss 포함 필요량",
                "stock_qty": "현재 재고",
                "inbound_qty": "입고 참고량",
                "outbound_qty": "출고 참고량",
                "available_reference_qty": "참고 가용량",
                "shortage_qty": "부족 참고량",
                "moq": "MOQ",
                "recommended_order_qty": "권장 발주량",
                "leadtime_days": "리드타임(일)",
                "days_left": "생산일까지 남은 일수",
                "urgency_reason": "긴급/주의 사유",
                "basis": "판단 기준",
                "action": "요청 액션",
            }),
            use_container_width=True,
            hide_index=True,
        )

        st.subheader("3. 업체별 요청 문장")
        need = dv[(dv["recommended_order_qty"] > 0) | (dv["risk"] == "긴급")]
        if need.empty:
            st.success("현재 기준 추가 발주 또는 긴급 확인이 필요하지 않습니다.")
        else:
            for _, r in need.iterrows():
                if r["recommended_order_qty"] > 0:
                    st.markdown(
                        f"- **{r['supplier'] or '공급업체 확인'}**에 "
                        f"**{r['item']} / {r['material_color']}** "
                        f"**{fmt_qty(r['recommended_order_qty'])}{r['unit']}**를 "
                        f"생산 전까지 입고 가능한지 확인 필요. "
                        f"({r['basis']})"
                    )
                else:
                    st.markdown(
                        f"- **{r['supplier'] or '공급업체 확인'}**에 "
                        f"**{r['item']} / {r['material_color']}** 입고/ETA 상태 확인 필요. "
                        f"({r['basis']})"
                    )

    st.subheader("4. 원본 데이터 확인")
    with st.expander("파싱된 생산계획 원본"):
        st.dataframe(parsed["production"].head(300), use_container_width=True)
    with st.expander("파싱된 BOM 원본"):
        st.dataframe(parsed["bom"].head(300), use_container_width=True)
    with st.expander("파싱된 재고 원본"):
        st.dataframe(parsed["inventory"].head(300), use_container_width=True)
    with st.expander("파싱된 입고 원본"):
        st.dataframe(parsed["inbound"].head(300), use_container_width=True)
    with st.expander("파싱된 출고 원본"):
        st.dataframe(parsed["outbound"].head(300), use_container_width=True)


def show_quality(parsed: Dict[str, pd.DataFrame], availability: pd.DataFrame):
    st.header("데이터 점검")
    q = pd.DataFrame([
        {"시트/데이터": "생산계획", "행 수": len(parsed["production"]), "상태": "OK" if len(parsed["production"]) else "확인 필요"},
        {"시트/데이터": "자재품목리스트(BOM)", "행 수": len(parsed["bom"]), "상태": "OK" if len(parsed["bom"]) else "확인 필요"},
        {"시트/데이터": "자재 재고 내역", "행 수": len(parsed["inventory"]), "상태": "OK" if len(parsed["inventory"]) else "확인 필요"},
        {"시트/데이터": "입고내역", "행 수": len(parsed["inbound"]), "상태": "OK" if len(parsed["inbound"]) else "선택"},
        {"시트/데이터": "출고내역", "행 수": len(parsed["outbound"]), "상태": "OK" if len(parsed["outbound"]) else "선택"},
    ])
    st.dataframe(q, use_container_width=True, hide_index=True)

    st.subheader("입고/출고 참고 가용량")
    st.caption("참고 가용량 = 현재 재고 + 입고 참고량 - 출고 참고량. 실제 부족 판단은 생산계획 MATERIAL 상태를 우선합니다.")
    st.dataframe(availability.head(300), use_container_width=True, hide_index=True)


# =========================================================
# 6. Main
# =========================================================

def main():
    st.sidebar.title("설정")
    uploaded = st.sidebar.file_uploader("엑셀 업로드", type=["xlsx", "xlsm"])
    loss_rate = st.sidebar.number_input("Loss 반영률", min_value=0.0, max_value=0.2, value=0.03, step=0.005, format="%.3f")

    st.sidebar.markdown("---")
    st.sidebar.subheader("분석 범위")
    include_ok = st.sidebar.checkbox("MATERIAL=OK 생산건도 포함", value=False)
    include_completed = st.sidebar.checkbox("완료 생산건도 포함", value=False)
    st.sidebar.caption("기본값은 MATERIAL=OK와 완료건을 제외하여, 실제 확인/발주 검토가 필요한 건만 봅니다.")

    if uploaded is None:
        show_intro()
        return

    try:
        sheet_names, sheets = load_excel(uploaded.getvalue())

        sh_prod = find_sheet(sheet_names, ["생산계획"])
        sh_bom = find_sheet(sheet_names, ["자재품목"])
        sh_inv = find_sheet(sheet_names, ["자재 재고"])
        sh_in = find_sheet(sheet_names, ["자재입고"])
        sh_out = find_sheet(sheet_names, ["자재출고"])

        prod = parse_production(sheets.get(sh_prod, pd.DataFrame()))
        bom = parse_bom(sheets.get(sh_bom, pd.DataFrame()))
        inv = parse_inventory(sheets.get(sh_inv, pd.DataFrame()))
        inbound = parse_flow(sheets.get(sh_in, pd.DataFrame()), "입고")
        outbound = parse_flow(sheets.get(sh_out, pd.DataFrame()), "출고")

        parsed = {
            "production": prod,
            "bom": bom,
            "inventory": inv,
            "inbound": inbound,
            "outbound": outbound,
        }

        detail, summary, availability = build_analysis(
            prod=prod,
            bom=bom,
            inventory=inv,
            inbound=inbound,
            outbound=outbound,
            loss_rate=float(loss_rate),
            include_ok=bool(include_ok),
            include_completed=bool(include_completed),
        )

    except Exception as e:
        st.error("엑셀 분석 중 오류가 발생했습니다. 데이터 점검이 필요합니다.")
        st.exception(e)
        return

    tab1, tab2, tab3 = st.tabs(["대시보드", "PO 상세", "데이터 점검"])

    with tab1:
        show_dashboard(summary, detail)
    with tab2:
        show_detail(summary, detail, parsed)
    with tab3:
        show_quality(parsed, availability)


if __name__ == "__main__":
    main()
