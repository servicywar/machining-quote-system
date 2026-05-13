"""
engine/material.py — 재료비 계산, KOMIS API 연동, 폴백 처리

시세 소스 우선순위:
    1. KOMIS API (자동, 일별)
    2. data/market_prices.xlsx (수동 오버라이드)
    3. 하드코딩 기본값

재료비 공식 (변경 불가):
    중량(kg)  = 바운딩박스 부피(cm³) × 밀도(g/cm³) ÷ 1000
    재료비    = 중량(kg) × 시세(원/kg) × (1 + 로스율)
"""

import time
from pathlib import Path

import pandas as pd
import requests
import streamlit as st

BASE_DIR = Path(__file__).parent.parent
MATERIAL_MASTER_PATH = BASE_DIR / "data" / "material_master.xlsx"
MARKET_PRICES_PATH   = BASE_DIR / "data" / "market_prices.xlsx"

# 하드코딩 기본 시세 (원/kg) — KOMIS·Excel 모두 실패 시 폴백
_DEFAULT_PRICES: dict[str, float] = {
    "AL":  4800.0,   # 알루미늄
    "STS": 3200.0,   # 스테인리스
    "CU":  8500.0,   # 구리/황동
    "SM":  1200.0,   # 일반강
}

# 소재코드 → 시세 카테고리 매핑
_CODE_TO_CAT: dict[str, str] = {
    "AL6061": "AL", "AL7075": "AL",
    "STS304": "STS", "STS316": "STS",
    "C3604":  "CU",
    "SM45C":  "SM", "SCM440": "SM",
}

# KOMIS API 설정 (공개 API — 키 불필요, rate limit 있음)
_KOMIS_URL     = "https://www.komis.or.kr/openapi/lme_avg_price.do"
_KOMIS_TIMEOUT = 5  # 초


@st.cache_data(ttl=86400, show_spinner=False)
def get_market_prices(force_refresh: bool = False) -> dict:
    """
    시세 딕셔너리를 반환한다.
    반환 형식: {"AL": 4800.0, "STS": 3200.0, ..., "source": "KOMIS API"}

    force_refresh=True 이면 캐시를 무시하고 재조회한다.
    """
    if force_refresh:
        get_market_prices.clear()

    # 1순위: KOMIS API
    try:
        resp = requests.get(_KOMIS_URL, timeout=_KOMIS_TIMEOUT)
        if resp.status_code == 200:
            data = resp.json()
            prices = _parse_komis(data)
            if prices:
                prices["source"] = "KOMIS API"
                return prices
    except Exception:
        pass

    # 2순위: market_prices.xlsx
    try:
        if MARKET_PRICES_PATH.exists():
            df = pd.read_excel(MARKET_PRICES_PATH)
            prices: dict = {}
            for _, row in df.iterrows():
                cat = str(row.get("material_code", "")).strip().upper()
                if cat and row.get("price_per_kg"):
                    prices[cat] = float(row["price_per_kg"])
            if prices:
                prices["source"] = "market_prices.xlsx"
                return prices
    except Exception:
        pass

    # 3순위: 하드코딩 기본값
    result = dict(_DEFAULT_PRICES)
    result["source"] = "기본값 (시세 미연동)"
    return result


def _parse_komis(data) -> dict:
    """KOMIS API 응답을 {카테고리: 시세} 딕셔너리로 파싱한다."""
    # KOMIS 응답 구조가 바뀔 수 있으므로 안전하게 처리
    prices: dict = {}
    try:
        items = data.get("items") or data.get("data") or []
        for item in items:
            name = str(item.get("metalName", "") or item.get("name", "")).upper()
            price = item.get("avgPrice") or item.get("price")
            if price is None:
                continue
            price = float(price)
            if "알루미늄" in name or "ALUM" in name or "AL" in name:
                prices["AL"] = price
            elif "스테인리스" in name or "STS" in name or "STAINLESS" in name:
                prices["STS"] = price
            elif "구리" in name or "COPPER" in name or "CU" in name:
                prices["CU"] = price
            elif "강" in name or "STEEL" in name or "SM" in name:
                prices["SM"] = price
    except Exception:
        pass
    return prices


@st.cache_data(show_spinner=False)
def load_material_master() -> pd.DataFrame:
    """
    data/material_master.xlsx 를 읽어 DataFrame으로 반환한다.
    파일이 없으면 기본 샘플 데이터를 반환한다.
    """
    if MATERIAL_MASTER_PATH.exists():
        try:
            return pd.read_excel(MATERIAL_MASTER_PATH)
        except Exception:
            pass

    # 파일 없을 때 인메모리 기본값
    return pd.DataFrame([
        {"material_code": "AL6061", "material_name": "알루미늄 6061",
         "density": 2.70, "machinability": 1.0,
         "default_loss_bar": 0.15, "default_loss_plate": 0.08,
         "note": "일반 가공 기준"},
        {"material_code": "AL7075", "material_name": "알루미늄 7075",
         "density": 2.82, "machinability": 0.9,
         "default_loss_bar": 0.15, "default_loss_plate": 0.08,
         "note": "항공·고강도"},
        {"material_code": "STS304", "material_name": "스테인리스 304",
         "density": 7.93, "machinability": 0.5,
         "default_loss_bar": 0.12, "default_loss_plate": 0.07,
         "note": "내식성 범용"},
        {"material_code": "STS316", "material_name": "스테인리스 316",
         "density": 7.98, "machinability": 0.45,
         "default_loss_bar": 0.12, "default_loss_plate": 0.07,
         "note": "내산성·의료"},
        {"material_code": "C3604",  "material_name": "황동 C3604",
         "density": 8.50, "machinability": 1.5,
         "default_loss_bar": 0.10, "default_loss_plate": 0.07,
         "note": "쾌삭 황동"},
        {"material_code": "SM45C",  "material_name": "탄소강 SM45C",
         "density": 7.85, "machinability": 0.65,
         "default_loss_bar": 0.10, "default_loss_plate": 0.06,
         "note": "범용 기계구조용"},
        {"material_code": "SCM440", "material_name": "합금강 SCM440",
         "density": 7.85, "machinability": 0.55,
         "default_loss_bar": 0.10, "default_loss_plate": 0.06,
         "note": "고강도 구조용"},
    ])


def calc_material_cost(
    volume_cm3: float,
    material_code: str,
    form: str = "bar",
    loss_rate_override: float | None = None,
    price_override: float | None = None,
) -> dict:
    """
    재료비를 계산하고 결과 딕셔너리를 반환한다.

    Args:
        volume_cm3:          바운딩박스 부피 (cm³)
        material_code:       소재 코드 (예: AL6061)
        form:                원자재 형태 ("bar" 봉재 / "plate" 판재)
        loss_rate_override:  로스율 직접 지정 (None이면 마스터 기본값 사용)
        price_override:      시세 직접 지정 (None이면 자동 조회)

    Returns:
        {
            "material_cost": float,   # 재료비 (원)
            "weight_kg":     float,   # 중량 (kg)
            "price_used":    float,   # 사용된 시세 (원/kg)
            "loss_rate":     float,   # 로스율 (소수)
            "source":        str,     # 시세 출처
        }
    """
    mat_df = load_material_master()
    row = mat_df[mat_df["material_code"] == material_code]
    if row.empty:
        raise ValueError(f"소재 코드 '{material_code}'를 마스터에서 찾을 수 없습니다.")

    mat = row.iloc[0]
    density  = float(mat["density"])
    loss_key = f"default_loss_{form}"
    loss_rate = (
        float(loss_rate_override)
        if loss_rate_override is not None
        else float(mat.get(loss_key, 0.15))
    )

    weight_kg = (volume_cm3 * density) / 1000.0

    if price_override is not None:
        price_used = float(price_override)
        source     = "수동 입력"
    else:
        cat    = _CODE_TO_CAT.get(material_code, "SM")
        prices = get_market_prices()
        price_used = prices.get(cat, _DEFAULT_PRICES.get(cat, 1200.0))
        source     = prices.get("source", "기본값")

    material_cost = weight_kg * price_used * (1 + loss_rate)

    return {
        "material_cost": material_cost,
        "weight_kg":     weight_kg,
        "price_used":    price_used,
        "loss_rate":     loss_rate,
        "source":        source,
    }
