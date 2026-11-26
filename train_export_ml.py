#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LMH-Tomato-Planner v4.0 연동용 머신러닝 학습 및 예측 JSON(ml_pred.json) 내보내기 스크립트

지원 입력(HTML / Streamlit → Python):
  1) 다중 실측 파일 폴더(--calib-dir, 기본: data/calib_csv): .csv, .xlsx, .xls 자동 스캔·병합
  2) HTML [계획 저장(JSON)] / [캘리브레이션 프리셋 저장] / localStorage(LMH_TOMATO_PLANNER_V3_CALIB_STATE) JSON
  3) HTML [마스터플랜 CSV 다운로드] (토마토_마스터플랜_v3_*.csv)
  4) HTML [CSV 양식 받기] 실측 입력 CSV/Excel (토마토_실측데이터_입력양식.csv / .xlsx / .xls)
  5) 기상 관측 CSV (기상청 ASOS 일자료 EUC-KR/UTF-8 다중 지점 자동 분리 · Open-Meteo kWh→MJ 자동 변환)
  6) 기존 SQLite DB (smartfarm_yield_swing.db / weekly_yield_data) 및 원천 CSV (raw_yield_source.csv)

지원 출력(Python → public/simulator.html / Streamlit app.py):
  - Vite 정적 서빙 구조(public/ 디렉터리 존재 시)에서는 기본 'public/ml_pred.json'으로 저장되어
    public/simulator.html의 [ML JSON 불러오기](fetch("./ml_pred.json")) 및 Streamlit 임베디드 뷰와 즉시 연동됨
  - 단독 실행 환경(public/ 미존재 시)에서는 현재 경로 'ml_pred.json'으로 저장되며 --output으로 변경 가능
  - [JSON 파일 선택] 버튼으로도 수동 업로드 가능
    (week, week_no, start, date, y_pred, pred, yield, y_actual, orig_kg, adj_kg, k_total, dyn_lag, MAE/RMSE)
"""

from __future__ import annotations

import argparse
import csv
import glob
import json
import math
import os
import random
import re
import sqlite3
import xml.etree.ElementTree as ET
import zipfile
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

try:
    import numpy as np  # type: ignore
except ImportError:
    np = None  # type: ignore

try:
    import pandas as pd  # type: ignore
except ImportError:
    pd = None  # type: ignore

try:
    from sklearn.ensemble import RandomForestRegressor as SklearnRandomForestRegressor  # type: ignore
except ImportError:
    SklearnRandomForestRegressor = None  # type: ignore

try:
    import openpyxl  # type: ignore
except ImportError:
    openpyxl = None  # type: ignore

try:
    import xlrd  # type: ignore
except ImportError:
    xlrd = None  # type: ignore


# ==============================================================================
# 1. HTML(LMH-Tomato-Planner-v4.0)과 100% 동기화된 상수 · 기본 환경 프리셋
# ==============================================================================
DAY_SEC = 86400
PYEONG_TO_M2 = 3.305785
DEFAULT_BASE_EFF = 8.0
PAR_FACTOR_DEFAULT = 2.08
DEFAULT_CALIB_DIR = "data/calib_csv"
SUPPORTED_CALIB_EXTS = (".csv", ".xlsx", ".xls")

FEATURES = [
    "ma_3week",
    "dli_daily",
    "nat_dli",
    "add_dli",
    "mj_out",
    "eff_w",
    "t_day",
    "t_night",
    "t_mean",
    "dif",
    "rh_mean",
    "vpd_mean",
    "co2_mean",
    "rtr",
    "k_vpd",
    "k_temp",
    "k_co2",
    "k_light",
    "k_total",
    "dyn_lag",
    "sink_part",
    "ss_ratio",
    "fruit_load",
    "adj_kg_sim",
    "orig_kg_sim",
    "cum_p_n_8w",
    "tr_ratio",
]
TARGET = "y_actual"


def _make_env_row(temp: float, dif: float, rh: float, co2: float) -> Dict[str, float]:
    t24 = round(temp, 1)
    d = round(dif, 1)
    return {
        "temp": t24,
        "tDay": round(t24 + d / 2.0, 1),
        "tNight": round(t24 - d / 2.0, 1),
        "rh": float(rh),
        "co2": float(co2),
    }


STANDARD_ENV_MONTHS: List[Dict[str, float]] = [
    _make_env_row(18.5, 4.0, 76, 650),
    _make_env_row(19.5, 4.5, 74, 700),
    _make_env_row(21.0, 5.0, 72, 750),
    _make_env_row(22.0, 5.0, 70, 750),
    _make_env_row(23.5, 5.0, 68, 650),
    _make_env_row(26.0, 5.5, 75, 450),
    _make_env_row(29.0, 5.0, 82, 400),
    _make_env_row(30.0, 5.0, 80, 400),
    _make_env_row(25.0, 5.0, 73, 550),
    _make_env_row(22.5, 4.5, 70, 700),
    _make_env_row(20.5, 4.5, 74, 700),
    _make_env_row(19.0, 4.0, 78, 650),
]

# 논산(기본 지점) 월별 평년 일사량(MJ/㎡/일) 폴백값
DEFAULT_MONTHLY_MJ = [
    8.6, 11.4, 14.8, 17.5, 19.6, 18.2, 14.9, 15.6, 13.7, 11.9, 8.7, 7.6
]

DEFAULT_CONFIG: Dict[str, Any] = {
    "target": 55.0,
    "area": 3305.8,
    "mass": 50.0,
    "fruitsPerTruss": 12.0,
    "plantDensity": 2.5,
    "trans": 70.0,
    "factor": PAR_FACTOR_DEFAULT,
    "ppfd": 120.0,
    "maxHours": 10.0,
    "ceiling": 26.0,
    "firstSetDays": 28,
    "cropLag": 45,
    "toppingDays": 50,
    "eff": DEFAULT_BASE_EFF,
    "start": "2027-01-01",
    "harvestEnd": "2027-12-31",
    "useSinkCurve": True,
    "useDynamicLag": True,
    "useCalibWeather": True,
}


# ==============================================================================
# 2. 유틸리티 함수 (날짜 · 숫자 변환 · 인코딩 자동 감지)
# ==============================================================================
def clamp(val: float, lo: float, hi: float) -> float:
    return min(hi, max(lo, val))


def is_finite(val: Any) -> bool:
    if val is None or isinstance(val, bool):
        return False
    try:
        return math.isfinite(float(val))
    except (TypeError, ValueError):
        return False


def to_float(val: Any, default: Optional[float] = None) -> Optional[float]:
    if val is None or isinstance(val, bool):
        return default
    if isinstance(val, (int, float)):
        f = float(val)
        return f if math.isfinite(f) else default
    s = str(val).strip().replace(",", "")
    if not s or s in ("-", "—", "nan", "NaN", "null", "None"):
        return default
    try:
        f = float(s)
        return f if math.isfinite(f) else default
    except ValueError:
        return default


def parse_date(s: str) -> date:
    s_clean = str(s).strip().strip('"').strip("'")
    m = re.match(r"^(\d{4})[-./](\d{1,2})[-./](\d{1,2})", s_clean)
    if m:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m2 = re.match(r"^(\d{4})(\d{2})(\d{2})$", s_clean)
    if m2:
        return date(int(m2.group(1)), int(m2.group(2)), int(m2.group(3)))
    raise ValueError(f"날짜 형식(YYYY-MM-DD)을 해석할 수 없습니다: {s}")


def normalize_weather_date(s: Any) -> Optional[str]:
    if s is None:
        return None
    if isinstance(s, datetime):
        return s.date().isoformat()
    if isinstance(s, date):
        return s.isoformat()
    try:
        return parse_date(str(s)).isoformat()
    except ValueError:
        return None


def decode_text_auto(path: Path) -> str:
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "utf-8", "cp949", "euc-kr"):
        try:
            text = raw.decode(enc)
            if text.startswith("\ufeff"):
                text = text[1:]
            return text
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace").lstrip("\ufeff")


def plan_periods(start_str: str, harvest_end_str: str) -> List[Dict[str, Any]]:
    start_dt = parse_date(start_str)
    end_dt = parse_date(harvest_end_str)
    days = (end_dt - start_dt).days + 1
    if days < 1 or days > 731:
        raise ValueError(f"수확 종료일({harvest_end_str})은 정식일({start_str}) 이후 1~731일 이내여야 합니다.")
    n_weeks = math.ceil(days / 7)
    return [
        {
            "at": start_dt + timedelta(days=i * 7),
            "days": min(7, days - i * 7),
        }
        for i in range(n_weeks)
    ]


# ==============================================================================
# 3. HTML(calcPhysiology / calcDynamicLagFromLookup / computeWeekEff) 1:1 엔진
# ==============================================================================
def calc_physiology(
    temp: Optional[float],
    rh: Optional[float],
    co2: Optional[float],
    dif: Optional[float] = 4.5,
    dli: Optional[float] = 18.0,
) -> Dict[str, float]:
    """HTML calcPhysiology()와 수식·상수·클램프 100% 동일"""
    t = float(temp) if is_finite(temp) else 22.0
    rh_val = clamp(float(rh) if is_finite(rh) else 70.0, 5.0, 100.0)
    co2_val = max(150.0, float(co2) if is_finite(co2) else 400.0)
    dif_val = float(dif) if is_finite(dif) else 4.5
    dli_val = float(dli) if is_finite(dli) else 18.0

    # [1] Tetens SVP & VPD (kPa)
    svp = 0.61078 * math.exp((17.27 * t) / (t + 237.3))
    vpd = max(0.0, svp * (1.0 - rh_val / 100.0))

    # [2] VPD 양방향 가중치 (K_vpd)
    if vpd < 0.40:
        kvpd = max(0.50, 0.85 - (0.40 - vpd) * 0.90)
    elif vpd < 0.75:
        kvpd = 0.85 + 0.15 * ((vpd - 0.40) / 0.35)
    elif vpd <= 1.15:
        kvpd = 1.0 + 0.025 * max(0.0, 1.0 - abs(vpd - 0.95) / 0.20)
    else:
        kvpd = max(0.40, 1.0 - (vpd - 1.15) * 0.32)

    # [3] 온도 + DIF + RTR(적정온도 대비 밸런스)
    if t < 20.0:
        ktemp_base = max(0.45, 1.0 - (20.0 - t) * 0.045)
    elif t > 25.0:
        ktemp_base = max(0.40, 1.0 - (t - 25.0) * 0.055)
    else:
        ktemp_base = 1.0

    if dif_val < 2.0:
        kdif = max(0.92, 1.0 - (2.0 - dif_val) * 0.015)
    elif dif_val > 7.0:
        kdif = max(0.92, 1.0 - (dif_val - 7.0) * 0.015)
    else:
        kdif = 1.0

    t_opt = clamp(16.5 + 0.25 * dli_val, 17.5, 25.5)
    rtr = t_opt / max(10.0, t)
    if 0.90 <= rtr <= 1.10:
        krtr = 1.0 + 0.035 * max(0.0, 1.0 - abs(rtr - 1.0) / 0.10)
    else:
        dev = abs(rtr - 1.0) - 0.10
        krtr = max(0.82, 1.0 - dev * 0.35)

    ktemp = ktemp_base * kdif * krtr

    # [4] CO2 시비 효과 (K_co2)
    if co2_val <= 400.0:
        kco2 = max(0.68, 1.0 - (400.0 - co2_val) * 0.0012)
    else:
        kco2 = 1.0 + 0.14 * (1.0 - math.exp(-(co2_val - 400.0) / 340.0))

    # [5] 광포화 단일 보정 (K_light)
    klight = max(0.84, 1.0 - (dli_val - 26.0) * 0.007) if dli_val > 26.0 else 1.0

    # [6] 종합 생리 가중치 (K_total ∈ [0.35, 1.15])
    ktotal = clamp(kvpd * ktemp * kco2 * klight, 0.35, 1.15)
    return {
        "T": t,
        "RH": rh_val,
        "CO2": co2_val,
        "dif": dif_val,
        "dli": dli_val,
        "svp": svp,
        "vpd": vpd,
        "rtr": rtr,
        "kvpd": kvpd,
        "ktemp": ktemp,
        "kdif": kdif,
        "krtr": krtr,
        "kco2": kco2,
        "klight": klight,
        "ktotal": ktotal,
    }


def calc_dynamic_lag_from_lookup(
    set_dt: date,
    gdd_target: float,
    get_t24_fn: Callable[[date], float],
) -> int:
    """HTML calcDynamicLagFromLookup()와 동일: T_base=8.5°C 적산온도 기반 성숙일수(21~110일)"""
    accum = 0.0
    d = 0
    while d < 110:
        dt = set_dt + timedelta(days=d)
        t24 = get_t24_fn(dt)
        t_val = float(t24) if is_finite(t24) else 20.0
        daily_gdd = clamp(t_val - 8.5, 1.5, 28.5)
        accum += daily_gdd
        d += 1
        if accum >= gdd_target:
            break
    return int(clamp(d, 21, 110))


def compute_week_eff(
    base_eff: float,
    period_at: date,
    period_days: int,
    crop_start: date,
    profile: str = "flat",
    existing_eff: Optional[float] = None,
) -> float:
    """HTML computeWeekEff()와 동일: 정식 후 군락 LAI 램프업 또는 역산/고정 값"""
    eff0 = clamp(float(base_eff) if is_finite(base_eff) else DEFAULT_BASE_EFF, 0.1, 30.0)
    if profile == "calibrated" and is_finite(existing_eff):
        return round(float(existing_eff), 2)
    if profile in ("flat", "calibrated"):
        return round(eff0, 2)

    mid_at = period_at + timedelta(days=max(0, period_days - 1) // 2)
    days_after_plant = max(0.0, float((mid_at - crop_start).days))
    lai_factor = 0.55 + 0.45 * (1.0 - math.exp(-days_after_plant / 24.0))
    return clamp(round(eff0 * lai_factor, 2), 0.1, 30.0)


def simulate_masterplan(
    config: Dict[str, Any],
    week_rows: List[Dict[str, Any]],
    env_by_month: List[Dict[str, float]],
    obs_rows: Optional[List[Dict[str, Optional[float]]]] = None,
) -> Dict[str, Any]:
    """
    HTML calculateMasterplan()을 Python으로 100% 재현하여
    착과 주차(detail) 및 수확 주차(harvest) 물리·생리 시뮬레이션 결과와 ML 피처를 산출
    """
    c = {**DEFAULT_CONFIG, **config}
    start_dt = parse_date(str(c["start"]))
    harvest_end_dt = parse_date(str(c["harvestEnd"]))
    first_set_days = int(round(float(c["firstSetDays"])))
    crop_lag = int(round(float(c["cropLag"])))
    topping_days = int(round(float(c["toppingDays"])))

    first_set = start_dt + timedelta(days=first_set_days)
    topping_limit = min(
        harvest_end_dt - timedelta(days=topping_days),
        harvest_end_dt - timedelta(days=14),
    )
    if first_set > harvest_end_dt - timedelta(days=21):
        raise ValueError("수확 종료일이 너무 빠릅니다. 첫 개화일과 최소 성숙일수를 확보하세요.")

    periods = plan_periods(str(c["start"]), str(c["harvestEnd"]))
    if len(week_rows) != len(periods):
        raise ValueError(f"주별 행 수({len(week_rows)})가 재배 기간({len(periods)}주)과 일치하지 않습니다.")

    obs_list = obs_rows or []
    use_sink = c.get("useSinkCurve", True) is not False
    use_dyn_lag = c.get("useDynamicLag", True) is not False
    use_obs_weather = c.get("useCalibWeather", True) is not False

    gdd_target = 525.0 * (crop_lag / 45.0)
    fallback_eff = float(c["eff"]) if is_finite(c.get("eff")) and 0.1 <= float(c["eff"]) <= 30.0 else DEFAULT_BASE_EFF
    ppfd = float(c["ppfd"])
    max_hours = float(c["maxHours"])
    ceiling = float(c["ceiling"])
    factor = float(c["factor"])
    trans = float(c["trans"])
    target_yield = float(c["target"])
    mass_g = float(c["mass"])
    fruits_per_truss = float(c["fruitsPerTruss"])
    plant_density = float(c["plantDensity"])
    max_supp_dli = ppfd * max_hours * 0.0036

    def get_week_idx(dt: date) -> int:
        idx = (dt - start_dt).days // 7
        return max(0, min(len(periods) - 1, idx))

    def get_day_t24(dt: date) -> float:
        w_idx = get_week_idx(dt)
        if w_idx < len(obs_list):
            ob = obs_list[w_idx]
            if use_obs_weather and ob and is_finite(ob.get("obsT24")):
                return float(ob["obsT24"])
        m_idx = dt.month - 1
        em = env_by_month[m_idx] if 0 <= m_idx < len(env_by_month) else {}
        return float(em["temp"]) if is_finite(em.get("temp")) else 22.0

    def get_day_mj(dt: date, row_mj: float) -> float:
        w_idx = get_week_idx(dt)
        if w_idx < len(obs_list):
            ob = obs_list[w_idx]
            if use_obs_weather and ob and is_finite(ob.get("obsMj")):
                return float(ob["obsMj"])
        return float(row_mj)

    # 1차 패스: 일별 착과 가능일, 실외 일사, 자연 DLI, 동적 성숙일수(lag)
    pre_days: List[List[Dict[str, Any]]] = []
    first_harv_dt: Optional[date] = None
    detail: List[Dict[str, Any]] = []

    for i, p in enumerate(periods):
        r = week_rows[i]
        row_mj = clamp(float(r.get("mj", 12.0)), 0.0, 50.0)
        row_lag = int(clamp(round(float(r.get("lag", crop_lag))), 1, 180))
        week_eff = clamp(float(r["eff"]) if is_finite(r.get("eff")) else fallback_eff, 0.1, 30.0)
        week_eff_rate = week_eff / 1000.0

        mid_at = p["at"] + timedelta(days=p["days"] // 2)
        m_idx = mid_at.month - 1
        em = env_by_month[m_idx]
        week_t24 = get_day_t24(mid_at)
        t_day = float(em["tDay"]) if is_finite(em.get("tDay")) else round(week_t24 + 2.25, 1)
        t_night = float(em["tNight"]) if is_finite(em.get("tNight")) else round(week_t24 - 2.25, 1)
        week_dif = t_day - t_night
        eff_mj = get_day_mj(mid_at, row_mj)
        nat = eff_mj * factor * trans / 100.0
        q = min(nat, ceiling)

        eligible_recs: List[Dict[str, Any]] = []
        dyn_lag_sum = 0
        for j in range(p["days"]):
            dt = p["at"] + timedelta(days=j)
            if dt < first_set or dt > topping_limit:
                continue
            lag_days = calc_dynamic_lag_from_lookup(dt, gdd_target, get_day_t24) if use_dyn_lag else row_lag
            harv_dt = dt + timedelta(days=lag_days)
            if harv_dt <= harvest_end_dt:
                if first_harv_dt is None or harv_dt < first_harv_dt:
                    first_harv_dt = harv_dt
                eligible_recs.append({
                    "dt": dt,
                    "lagDays": lag_days,
                    "harvDt": harv_dt,
                    "weekIdx": i,
                    "monthIdx": dt.month - 1,
                })
                dyn_lag_sum += lag_days

        eligible_days = len(eligible_recs)
        if eligible_days > 0:
            dyn_lag = int(round(dyn_lag_sum / eligible_days))
        elif use_dyn_lag:
            dyn_lag = calc_dynamic_lag_from_lookup(mid_at, gdd_target, get_day_t24)
        else:
            dyn_lag = row_lag

        available = max(0.0, min(max_supp_dli, ceiling - nat)) if eligible_days > 0 else 0.0
        pre_days.append(eligible_recs)

        detail.append({
            "week": i + 1,
            "at": p["at"],
            "days": p["days"],
            "mj": row_mj,
            "effMj": eff_mj,
            "lag": row_lag,
            "eff": week_eff,
            "weekEffRate": week_eff_rate,
            "eligibleDays": eligible_days,
            "dynLag": dyn_lag,
            "nat": nat,
            "q": q,
            "available": available,
            "monthIdx": m_idx,
            "weekT24": week_t24,
            "weekTDay": t_day,
            "weekTNight": t_night,
            "weekDif": week_dif,
            "weekRh": float(em.get("rh", 70.0)),
            "weekCo2": float(em.get("co2", 550.0)),
            "eligibleStart": eligible_recs[0]["dt"] if eligible_days > 0 else p["at"],
        })

    # 2차 패스: 초기 1~3화방 Sink Partitioning 분배율 (0.30 -> 0.55 -> 0.80 -> 1.00)
    for i, r in enumerate(detail):
        recs = pre_days[i]
        sink_sum = 0.0
        for rec in recs:
            if not use_sink:
                sink_part = 1.0
            elif first_harv_dt is None:
                sink_part = 1.0
            else:
                h_offset = (rec["harvDt"] - first_harv_dt).days // 7
                if h_offset <= 0:
                    sink_part = 0.30
                elif h_offset == 1:
                    sink_part = 0.55
                elif h_offset == 2:
                    sink_part = 0.80
                else:
                    sink_part = 1.00
            rec["sinkPart"] = sink_part
            sink_sum += sink_part
        r["dayRecords"] = recs
        r["avgSinkPart"] = sink_sum / r["eligibleDays"] if r["eligibleDays"] > 0 else 1.0
        r["base"] = r["q"] * r["eligibleDays"] * r["weekEffRate"] * r["avgSinkPart"]

    # 3차 패스: 목표 수량 달성을 위한 필요 보광량 배분(이진 탐색)
    base_sum = sum(r["base"] for r in detail)
    max_sum = sum((r["q"] + r["available"]) * r["eligibleDays"] * r["weekEffRate"] * r["avgSinkPart"] for r in detail)
    goal = min(max(base_sum, target_yield), max_sum)

    lo, hi = 0.0, ceiling
    for _ in range(70):
        mid = (lo + hi) / 2.0
        kg_sum = sum(
            (r["q"] + min(r["available"], max(0.0, mid - r["q"]))) * r["eligibleDays"] * r["weekEffRate"] * r["avgSinkPart"]
            for r in detail
        )
        if kg_sum < goal:
            lo = mid
        else:
            hi = mid
    level = (lo + hi) / 2.0

    # 4차 패스: 생리 보정(K_total) & Source-Sink 부하율
    for r in detail:
        r["add"] = 0.0 if target_yield <= base_sum else min(r["available"], max(0.0, level - r["q"]))
        r["hours"] = r["add"] / (ppfd * 0.0036) if ppfd > 0 else 0.0
        r["totalDli"] = r["nat"] + r["add"]
        r["phys"] = calc_physiology(r["weekT24"], r["weekRh"], r["weekCo2"], dif=r["weekDif"], dli=r["q"] + r["add"])
        r["kg"] = (r["q"] + r["add"]) * r["eligibleDays"] * r["weekEffRate"] * r["avgSinkPart"]
        r["adjBase"] = r["base"] * r["phys"]["ktotal"]
        r["adjKg"] = r["kg"] * r["phys"]["ktotal"]
        r["deltaKg"] = r["adjKg"] - r["kg"]
        r["weekHours"] = r["hours"] * r["eligibleDays"]
        r["harvestAt"] = r["eligibleStart"] + timedelta(days=r["dynLag"])

        if r["eligibleDays"] > 0:
            source_grams = r["adjKg"] * 1000.0
            truss_per_week = clamp(0.95 + (r["weekT24"] - 20.0) * 0.035, 0.65, 1.25)
            sink_demand_grams = (
                fruits_per_truss * plant_density * truss_per_week * (r["eligibleDays"] / 7.0) * mass_g * r["avgSinkPart"]
            )
            r["ssRatio"] = source_grams / sink_demand_grams if sink_demand_grams > 0 else 1.0
            if r["ssRatio"] < 0.85:
                r["predMass"] = mass_g * max(0.55, math.pow(r["ssRatio"] / 0.85, 0.55))
            else:
                r["predMass"] = mass_g * min(1.12, 1.0 + (r["ssRatio"] - 0.85) * 0.10)
            r["trussPerWeek"] = truss_per_week
        else:
            r["ssRatio"] = 1.0
            r["predMass"] = mass_g
            r["trussPerWeek"] = 1.0

        for rec in r["dayRecords"]:
            m = rec["monthIdx"]
            em = env_by_month[m]
            t24_day = get_day_t24(rec["dt"])
            dif_day = (float(em["tDay"]) - float(em["tNight"])) if (is_finite(em.get("tDay")) and is_finite(em.get("tNight"))) else 4.5
            rec["physDay"] = calc_physiology(t24_day, em["rh"], em["co2"], dif=dif_day, dli=r["q"] + r["add"])

    # 수확 주차(harvest) 집계 + 착과->수확 기여 가중 평균 피처 산출
    harvest: List[Dict[str, Any]] = [
        {
            "week": hi + 1,
            "at": p["at"],
            "end": p["at"] + timedelta(days=p["days"] - 1),
            "days": p["days"],
            "origBase": 0.0,
            "origKg": 0.0,
            "adjKg": 0.0,
            "unitAdj": 0.0,
            "lagSum": 0.0,
            "sinkSum": 0.0,
            "massSum": 0.0,
            "dayCount": 0,
            "unitBySetWeek": {},
            "contribWeightSum": 0.0,
            "contribTotals": {
                "dli_daily": 0.0,
                "nat_dli": 0.0,
                "add_dli": 0.0,
                "mj_out": 0.0,
                "eff_w": 0.0,
                "t_day": 0.0,
                "t_night": 0.0,
                "t_mean": 0.0,
                "dif": 0.0,
                "rh_mean": 0.0,
                "vpd_mean": 0.0,
                "co2_mean": 0.0,
                "rtr": 0.0,
                "k_vpd": 0.0,
                "k_temp": 0.0,
                "k_co2": 0.0,
                "k_light": 0.0,
                "k_total": 0.0,
                "ss_ratio": 0.0,
                "fruit_load": 0.0,
            },
        }
        for hi, p in enumerate(periods)
    ]

    for wi, r in enumerate(detail):
        for rec in r["dayRecords"]:
            idx = (rec["harvDt"] - start_dt).days // 7
            if 0 <= idx < len(harvest):
                phys_day = rec.get("physDay") or r["phys"]
                unit_day_adj = ((r["q"] + r["add"]) / 1000.0) * rec["sinkPart"] * phys_day["ktotal"]
                d_base = r["q"] * r["weekEffRate"] * rec["sinkPart"]
                d_kg = (r["q"] + r["add"]) * r["weekEffRate"] * rec["sinkPart"]
                h = harvest[idx]
                h["origBase"] += d_base
                h["origKg"] += d_kg
                h["adjKg"] += d_kg * phys_day["ktotal"]
                h["unitAdj"] += unit_day_adj
                h["lagSum"] += rec["lagDays"]
                h["sinkSum"] += rec["sinkPart"]
                h["massSum"] += r["predMass"]
                h["dayCount"] += 1
                h["unitBySetWeek"][wi] = h["unitBySetWeek"].get(wi, 0.0) + unit_day_adj

                w_contrib = max(d_kg, 1e-6)
                h["contribWeightSum"] += w_contrib
                ct = h["contribTotals"]
                ct["dli_daily"] += (r["q"] + r["add"]) * w_contrib
                ct["nat_dli"] += r["nat"] * w_contrib
                ct["add_dli"] += r["add"] * w_contrib
                ct["mj_out"] += r["effMj"] * w_contrib
                ct["eff_w"] += r["eff"] * w_contrib
                ct["t_day"] += r["weekTDay"] * w_contrib
                ct["t_night"] += r["weekTNight"] * w_contrib
                ct["t_mean"] += phys_day["T"] * w_contrib
                ct["dif"] += phys_day["dif"] * w_contrib
                ct["rh_mean"] += phys_day["RH"] * w_contrib
                ct["vpd_mean"] += phys_day["vpd"] * w_contrib
                ct["co2_mean"] += phys_day["CO2"] * w_contrib
                ct["rtr"] += phys_day["rtr"] * w_contrib
                ct["k_vpd"] += phys_day["kvpd"] * w_contrib
                ct["k_temp"] += phys_day["ktemp"] * w_contrib
                ct["k_co2"] += phys_day["kco2"] * w_contrib
                ct["k_light"] += phys_day["klight"] * w_contrib
                ct["k_total"] += phys_day["ktotal"] * w_contrib
                ct["ss_ratio"] += r["ssRatio"] * w_contrib
                ct["fruit_load"] += (r["trussPerWeek"] * fruits_per_truss * plant_density * rec["sinkPart"]) * w_contrib

    for hi, h in enumerate(harvest):
        r_same = detail[hi]
        if h["dayCount"] > 0 and h["contribWeightSum"] > 0:
            w_sum = h["contribWeightSum"]
            h["avgLag"] = h["lagSum"] / h["dayCount"]
            h["avgSink"] = h["sinkSum"] / h["dayCount"]
            h["predMass"] = h["massSum"] / h["dayCount"]
            for k, v in h["contribTotals"].items():
                h[k] = v / w_sum
        else:
            p_same = r_same["phys"]
            h["avgLag"] = float(r_same["dynLag"])
            h["avgSink"] = float(r_same["avgSinkPart"])
            h["predMass"] = float(r_same["predMass"])
            h["dli_daily"] = float(r_same["q"] + r_same["add"])
            h["nat_dli"] = float(r_same["nat"])
            h["add_dli"] = float(r_same["add"])
            h["mj_out"] = float(r_same["effMj"])
            h["eff_w"] = float(r_same["eff"])
            h["t_day"] = float(r_same["weekTDay"])
            h["t_night"] = float(r_same["weekTNight"])
            h["t_mean"] = float(p_same["T"])
            h["dif"] = float(p_same["dif"])
            h["rh_mean"] = float(p_same["RH"])
            h["vpd_mean"] = float(p_same["vpd"])
            h["co2_mean"] = float(p_same["CO2"])
            h["rtr"] = float(p_same["rtr"])
            h["k_vpd"] = float(p_same["kvpd"])
            h["k_temp"] = float(p_same["ktemp"])
            h["k_co2"] = float(p_same["kco2"])
            h["k_light"] = float(p_same["klight"])
            h["k_total"] = float(p_same["ktotal"])
            h["ss_ratio"] = float(r_same["ssRatio"])
            h["fruit_load"] = float(r_same["trussPerWeek"] * fruits_per_truss * plant_density * r_same["avgSinkPart"])

    return {
        "config": c,
        "detail": detail,
        "harvest": harvest,
        "first_harv_dt": first_harv_dt,
    }


# ==============================================================================
# 4. 기상 관측 CSV 파서 (기상청 ASOS · Open-Meteo 자동 인식 및 다중 지점 분리)
# ==============================================================================
def load_weather_csv(path: Path) -> Dict[str, Dict[str, Any]]:
    """
    HTML의 parseWeatherCsvText()와 동일한 규격으로 기상 관측 CSV 파싱.
    - Open-Meteo: 상단 latitude/longitude 메타데이터 스킵, kWh -> MJ(*3.6) 자동 변환
    - 기상청 ASOS 일자료: 지점명별 자동 분리, UTF-8 / EUC-KR / CP949 자동 디코딩
    """
    text = decode_text_auto(path)
    lines = text.splitlines()
    meta_lat: Optional[float] = None
    meta_lon: Optional[float] = None
    header_idx = -1
    cols: Optional[Dict[str, Any]] = None

    for i in range(min(len(lines), 40)):
        raw_line = lines[i].strip()
        if not raw_line:
            continue
        cells = next(csv.reader([raw_line]))
        low = [c.lower().replace(" ", "") for c in cells]

        if "latitude" in low and "longitude" in low and i + 1 < len(lines):
            next_cells = next(csv.reader([lines[i + 1].strip()]))
            la = to_float(next_cells[low.index("latitude")]) if low.index("latitude") < len(next_cells) else None
            lo = to_float(next_cells[low.index("longitude")]) if low.index("longitude") < len(next_cells) else None
            if la is not None and lo is not None:
                meta_lat, meta_lon = la, lo
            continue

        def find_col(pred: Callable[[str, str], bool]) -> int:
            for idx, c in enumerate(cells):
                if pred(c.strip(), low[idx]):
                    return idx
            return -1

        date_idx = find_col(lambda c, l: l in ("time", "date", "날짜") or c == "일시" or l.startswith("일시("))
        mj_idx = find_col(lambda c, l: "shortwave_radiation_sum" in l or "일사량" in c or "radiation_sum" in l)
        if date_idx >= 0 and mj_idx >= 0:
            station_name_idx = find_col(lambda c, l: c == "지점명" or "지점명" in l or l == "station_name")
            station_code_idx = find_col(lambda c, l: c in ("지점", "지점번호") or l == "station")
            temp_idx = find_col(lambda c, l: "temperature_2m_mean" in l or "평균기온" in c or ("기온" in c and "최고" not in c and "최저" not in c))
            rh_idx = find_col(lambda c, l: "relative_humidity_2m_mean" in l or "평균상대습도" in c or "평균 상대습도" in c or "상대습도" in c)
            mj_header = cells[mj_idx]
            is_kwh = bool(re.search(r"kwh", mj_header, re.IGNORECASE))
            is_open_meteo = any("shortwave_radiation_sum" in l or l == "time" for l in low)
            header_idx = i
            cols = {
                "date_idx": date_idx,
                "mj_idx": mj_idx,
                "station_name_idx": station_name_idx,
                "station_code_idx": station_code_idx,
                "temp_idx": temp_idx,
                "rh_idx": rh_idx,
                "is_kwh": is_kwh,
                "is_open_meteo": is_open_meteo,
            }
            break

    if header_idx < 0 or cols is None:
        raise ValueError(f"[{path.name}] Open-Meteo 또는 기상청 ASOS 일자료 헤더를 찾지 못했습니다.")

    buckets: Dict[str, Dict[str, Any]] = {}
    raw_base = re.sub(r"\.(csv|txt)$", "", path.name, flags=re.IGNORECASE).strip() or "관측기상"
    short_om_base = re.sub(r"_?open-meteo-.*$", "", raw_base, flags=re.IGNORECASE).strip() or raw_base

    reader = csv.reader(lines[header_idx + 1 :])
    for cells in reader:
        if not cells or cols["date_idx"] >= len(cells):
            continue
        d_str = normalize_weather_date(cells[cols["date_idx"]])
        if not d_str:
            continue

        station_key = "__single__"
        station_label = short_om_base
        if not cols["is_open_meteo"] and (cols["station_name_idx"] >= 0 or cols["station_code_idx"] >= 0):
            s_name = cells[cols["station_name_idx"]].strip() if 0 <= cols["station_name_idx"] < len(cells) else ""
            s_code = cells[cols["station_code_idx"]].strip() if 0 <= cols["station_code_idx"] < len(cells) else ""
            station_label = s_name or (f"지점 {s_code}" if s_code else raw_base)
            station_key = station_label

        if station_key not in buckets:
            buckets[station_key] = {
                "station_label": station_label,
                "daily": {},
                "t_sum": [0.0] * 12,
                "t_cnt": [0] * 12,
                "rh_sum": [0.0] * 12,
                "rh_cnt": [0] * 12,
                "min_date": None,
                "max_date": None,
            }
        b = buckets[station_key]
        if b["min_date"] is None or d_str < b["min_date"]:
            b["min_date"] = d_str
        if b["max_date"] is None or d_str > b["max_date"]:
            b["max_date"] = d_str

        m_idx = int(d_str[5:7]) - 1
        if cols["mj_idx"] < len(cells):
            mj_val = to_float(cells[cols["mj_idx"]])
            if mj_val is not None and mj_val >= 0:
                if cols["is_kwh"]:
                    mj_val *= 3.6
                b["daily"][d_str] = round(mj_val, 3)

        if 0 <= cols["temp_idx"] < len(cells):
            tv = to_float(cells[cols["temp_idx"]])
            if tv is not None and 0 <= m_idx < 12:
                b["t_sum"][m_idx] += tv
                b["t_cnt"][m_idx] += 1

        if 0 <= cols["rh_idx"] < len(cells):
            hv = to_float(cells[cols["rh_idx"]])
            if hv is not None and 0.0 <= hv <= 100.0 and 0 <= m_idx < 12:
                b["rh_sum"][m_idx] += hv
                b["rh_cnt"][m_idx] += 1

    datasets: Dict[str, Dict[str, Any]] = {}
    for b in buckets.values():
        if not b["daily"]:
            continue
        hydrated = hydrate_custom_city({
            "id": b["station_label"],
            "name": b["station_label"],
            "stationName": b["station_label"],
            "sourceType": "Open-Meteo 실측" if cols["is_open_meteo"] else "기상청 ASOS 실측",
            "lat": meta_lat,
            "lon": meta_lon,
            "daily": b["daily"],
            "monthlyTemp": [round(b["t_sum"][i] / b["t_cnt"][i], 1) if b["t_cnt"][i] > 0 else None for i in range(12)],
            "monthlyRh": [round(b["rh_sum"][i] / b["rh_cnt"][i], 1) if b["rh_cnt"][i] > 0 else None for i in range(12)],
        })
        if hydrated:
            datasets[b["station_label"]] = hydrated

    if not datasets:
        raise ValueError(f"[{path.name}] 유효한 일사량(MJ/㎡) 행을 찾지 못했습니다.")
    return datasets


def hydrate_custom_city(raw: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    daily_raw = raw.get("daily")
    if not isinstance(daily_raw, dict):
        return None
    cal_sum: Dict[str, float] = {}
    cal_cnt: Dict[str, int] = {}
    m_sum = [0.0] * 12
    m_cnt = [0] * 12
    clean_daily: Dict[str, float] = {}
    total_sum = 0.0
    total_cnt = 0

    for d_str, val in daily_raw.items():
        mj = to_float(val)
        if mj is None or mj < 0 or not re.match(r"^\d{4}-\d{2}-\d{2}$", str(d_str)):
            continue
        clean_daily[str(d_str)] = mj
        mmdd = str(d_str)[5:]
        m_idx = int(str(d_str)[5:7]) - 1
        cal_sum[mmdd] = cal_sum.get(mmdd, 0.0) + mj
        cal_cnt[mmdd] = cal_cnt.get(mmdd, 0) + 1
        if 0 <= m_idx < 12:
            m_sum[m_idx] += mj
            m_cnt[m_idx] += 1
        total_sum += mj
        total_cnt += 1

    if total_cnt == 0:
        return None

    overall_avg = total_sum / total_cnt
    cal_fallback = {k: round(cal_sum[k] / cal_cnt[k], 3) for k in cal_sum}
    monthly = [
        round(m_sum[i] / m_cnt[i], 3) if m_cnt[i] > 0 else round(overall_avg, 3)
        for i in range(12)
    ]
    return {
        "id": str(raw.get("id", "custom")),
        "name": str(raw.get("name") or raw.get("stationName") or raw.get("id", "custom")),
        "daily": clean_daily,
        "calFallback": cal_fallback,
        "overallAvg": round(overall_avg, 3),
        "monthly": monthly,
        "monthlyTemp": raw.get("monthlyTemp") or [None] * 12,
        "monthlyRh": raw.get("monthlyRh") or [None] * 12,
    }


def apply_weather_dataset_to_plan(
    ds: Dict[str, Any],
    config: Dict[str, Any],
    week_rows: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    """기상 관측 DB(원시 날짜 1:1 매칭 → MM-DD 평균 폴백)를 주별 행의 mj에 반영"""
    periods = plan_periods(str(config["start"]), str(config["harvestEnd"]))
    out_rows = [dict(r) for r in week_rows]
    for idx, p in enumerate(periods):
        if idx >= len(out_rows):
            break
        mj_sum = 0.0
        for j in range(p["days"]):
            dt = p["at"] + timedelta(days=j)
            d_str = dt.isoformat()
            if d_str in ds["daily"]:
                mj_val = ds["daily"][d_str]
            elif d_str[5:] in ds["calFallback"]:
                mj_val = ds["calFallback"][d_str[5:]]
            else:
                mj_val = ds["monthly"][dt.month - 1]
            mj_sum += mj_val
        out_rows[idx]["mj"] = round(mj_sum / p["days"], 3)
    return out_rows


# ==============================================================================
# 5. HTML 내보내기 파일 로더 (Plan JSON / localStorage / 마스터플랜 CSV / 실측 CSV)
# ==============================================================================
def _normalize_env_months(raw_env: Any) -> List[Dict[str, float]]:
    if not isinstance(raw_env, list) or len(raw_env) != 12:
        return [dict(x) for x in STANDARD_ENV_MONTHS]
    out: List[Dict[str, float]] = []
    for i, item in enumerate(raw_env):
        fb = STANDARD_ENV_MONTHS[i]
        if not isinstance(item, dict):
            out.append(dict(fb))
            continue
        t24 = to_float(item.get("temp"), fb["temp"]) or fb["temp"]
        t_day = to_float(item.get("tDay"), round(t24 + 2.25, 1)) or round(t24 + 2.25, 1)
        t_night = to_float(item.get("tNight"), round(t24 - 2.25, 1)) or round(t24 - 2.25, 1)
        rh = to_float(item.get("rh"), fb["rh"]) or fb["rh"]
        co2 = to_float(item.get("co2"), fb["co2"]) or fb["co2"]
        out.append({"temp": t24, "tDay": t_day, "tNight": t_night, "rh": rh, "co2": co2})
    return out


def _normalize_calib_rows(raw_calib: Any, n_weeks: int) -> List[Dict[str, Optional[float]]]:
    rows_out: List[Dict[str, Optional[float]]] = []
    src = raw_calib if isinstance(raw_calib, list) else []
    for i in range(n_weeks):
        item = src[i] if i < len(src) and isinstance(src[i], dict) else {}
        rows_out.append({
            "obsKg": to_float(item.get("obsKg")),
            "obsMass": to_float(item.get("obsMass")),
            "obsMj": to_float(item.get("obsMj")),
            "obsT24": to_float(item.get("obsT24")),
        })
    return rows_out


def load_html_plan_json(
    path: Path,
    weather_db_path: Optional[Path] = None,
) -> Dict[str, Any]:
    """
    HTML [계획 저장(JSON)] / [캘리브레이션 프리셋 저장] 또는
    localStorage["LMH_TOMATO_PLANNER_V3_CALIB_STATE"] JSON 파싱
    """
    raw = json.loads(decode_text_auto(path))
    if not isinstance(raw, dict):
        raise ValueError(f"[{path.name}] 올바른 JSON 객체가 아닙니다.")

    # 래퍼 형태({ state: ..., weatherDb: ... }) 지원
    weather_db_items: Dict[str, Any] = {}
    if isinstance(raw.get("weatherDb"), dict):
        wb = raw["weatherDb"]
        weather_db_items = wb.get("items", wb) if isinstance(wb.get("items", wb), dict) else {}
    if "state" in raw and isinstance(raw["state"], dict):
        raw = raw["state"]

    if weather_db_path and weather_db_path.exists():
        wb_raw = json.loads(decode_text_auto(weather_db_path))
        if isinstance(wb_raw, dict):
            weather_db_items.update(wb_raw.get("items", wb_raw) if isinstance(wb_raw.get("items", wb_raw), dict) else {})

    active = raw.get("active") if isinstance(raw.get("active"), dict) else {}
    cfg_raw = raw.get("config") if isinstance(raw.get("config"), dict) else {}
    config = {**DEFAULT_CONFIG, **cfg_raw}
    if active.get("start") and not cfg_raw.get("start"):
        config["start"] = active["start"]
    if active.get("harvestEnd") and not cfg_raw.get("harvestEnd"):
        config["harvestEnd"] = active["harvestEnd"]
    for flag in ("useSinkCurve", "useDynamicLag", "useCalibWeather"):
        if flag in active and flag not in cfg_raw:
            config[flag] = bool(active[flag])

    periods = plan_periods(str(config["start"]), str(config["harvestEnd"]))
    env_months = _normalize_env_months(raw.get("envMonths"))

    eff_profile = str(active.get("effProfile") or "flat")
    crop_start_dt = parse_date(str(config["start"]))
    base_eff = float(config.get("eff") or DEFAULT_BASE_EFF)
    crop_lag = int(round(float(config.get("cropLag") or 45)))

    raw_rows = raw.get("rows")
    week_rows: List[Dict[str, Any]] = []
    if isinstance(raw_rows, list) and len(raw_rows) == len(periods):
        for i, rr in enumerate(raw_rows):
            rr = rr if isinstance(rr, dict) else {}
            p = periods[i]
            eff_val = to_float(rr.get("eff"))
            if eff_val is None:
                eff_val = compute_week_eff(base_eff, p["at"], p["days"], crop_start_dt, eff_profile)
            week_rows.append({
                "mj": clamp(to_float(rr.get("mj"), DEFAULT_MONTHLY_MJ[p["at"].month - 1]) or 12.0, 0.0, 50.0),
                "lag": int(clamp(round(to_float(rr.get("lag"), crop_lag) or crop_lag), 1, 180)),
                "eff": clamp(eff_val, 0.1, 30.0),
            })
    else:
        months_dli = raw.get("monthsDli") if isinstance(raw.get("monthsDli"), list) and len(raw.get("monthsDli")) == 12 else DEFAULT_MONTHLY_MJ
        for i, p in enumerate(periods):
            mj_avg = sum(float(months_dli[(p["at"] + timedelta(days=j)).month - 1]) for j in range(p["days"])) / p["days"]
            week_rows.append({
                "mj": round(mj_avg, 3),
                "lag": crop_lag,
                "eff": compute_week_eff(base_eff, p["at"], p["days"], crop_start_dt, eff_profile),
            })

    calib_rows = _normalize_calib_rows(raw.get("calibRows"), len(periods))

    # customCityData 또는 weatherDb에서 custom city 확인
    city_id = str(active.get("city") or "nonsan")
    custom_city = None
    if isinstance(raw.get("customCityData"), dict):
        custom_city = hydrate_custom_city(raw["customCityData"])
    elif city_id in weather_db_items and isinstance(weather_db_items[city_id], dict):
        custom_city = hydrate_custom_city({**weather_db_items[city_id], "id": city_id})

    scenario = str(active.get("solarScenario") or "normal")
    zone_label = f"{custom_city['name']}" if custom_city else f"{city_id} ({scenario})"

    return {
        "source": "html_plan_json",
        "zone_label": zone_label,
        "config": config,
        "week_rows": week_rows,
        "env_months": env_months,
        "calib_rows": calib_rows,
        "custom_city": custom_city,
    }


def load_masterplan_csv(path: Path) -> Dict[str, Any]:
    """HTML [마스터플랜 CSV 다운로드] (토마토_마스터플랜_v3_*.csv) 파싱"""
    text = decode_text_auto(path)
    lines = text.splitlines()
    config = dict(DEFAULT_CONFIG)
    env_months = [dict(x) for x in STANDARD_ENV_MONTHS]

    weekly_header_idx = -1
    monthly_header_idx = -1
    zone_label = "Masterplan_CSV"

    for idx, line in enumerate(lines):
        cells = [c.strip() for c in next(csv.reader([line]))] if line.strip() else []
        if not cells:
            continue
        if cells[0] == "지역" and len(cells) >= 6:
            city_name = cells[1] or "Unknown"
            scen_name = cells[3] if len(cells) > 3 else ""
            zone_label = f"{city_name} ({scen_name})" if scen_name else city_name
            crop_span = cells[5] if len(cells) > 5 else ""
            m_span = re.match(r"(\d{4}-\d{2}-\d{2})\s*~\s*(\d{4}-\d{2}-\d{2})", crop_span)
            if m_span:
                config["start"] = m_span.group(1)
                config["harvestEnd"] = m_span.group(2)
            for k_i in range(6, len(cells) - 1, 2):
                if "재배면적(㎡)" in cells[k_i] and to_float(cells[k_i + 1]) is not None:
                    config["area"] = to_float(cells[k_i + 1])
        elif cells[0].startswith("목표 수확량"):
            if len(cells) > 1 and to_float(cells[1]) is not None:
                config["target"] = to_float(cells[1])
        elif cells[0].startswith("광효율 프로파일"):
            for k_i in range(0, len(cells) - 1, 2):
                if "화방 착과수" in cells[k_i] and to_float(cells[k_i + 1]) is not None:
                    config["fruitsPerTruss"] = to_float(cells[k_i + 1])
                elif "재식밀도" in cells[k_i] and to_float(cells[k_i + 1]) is not None:
                    config["plantDensity"] = to_float(cells[k_i + 1])
                elif "평균 광효율" in cells[k_i] and to_float(cells[k_i + 1]) is not None:
                    config["eff"] = to_float(cells[k_i + 1])
        elif cells[0] == "[월별 마스터플랜 비교표]":
            monthly_header_idx = idx + 1
        elif cells[0].startswith("[주별 세부 광량"):
            weekly_header_idx = idx + 1
            break

    # 월별 환경 테이블 복원 (T24, RH, CO2)
    if monthly_header_idx >= 0:
        for row_line in lines[monthly_header_idx + 1 :]:
            if not row_line.strip():
                break
            cells = [c.strip() for c in next(csv.reader([row_line]))]
            if not cells or not cells[0].endswith("월"):
                break
            m_num = to_float(cells[0].replace("월", ""))
            if m_num is None or not (1 <= int(m_num) <= 12):
                continue
            m_idx = int(m_num) - 1
            t24 = to_float(cells[7]) if len(cells) > 7 else None
            rh = to_float(cells[8]) if len(cells) > 8 else None
            co2 = to_float(cells[9]) if len(cells) > 9 else None
            if t24 is not None:
                env_months[m_idx]["temp"] = t24
                env_months[m_idx]["tDay"] = round(t24 + 2.25, 1)
                env_months[m_idx]["tNight"] = round(t24 - 2.25, 1)
            if rh is not None:
                env_months[m_idx]["rh"] = rh
            if co2 is not None:
                env_months[m_idx]["co2"] = co2

    if weekly_header_idx < 0:
        raise ValueError(f"[{path.name}] '[주별 세부 광량 ...]' 섹션을 찾을 수 없습니다.")

    week_rows: List[Dict[str, Any]] = []
    calib_rows: List[Dict[str, Optional[float]]] = []
    start_dates: List[str] = []
    end_days_last = 7

    for row_line in lines[weekly_header_idx + 1 :]:
        if not row_line.strip():
            continue
        cells = [c.strip() for c in next(csv.reader([row_line]))]
        if len(cells) < 15:
            continue
        wk_no = to_float(cells[0])
        if wk_no is None:
            continue
        s_date = normalize_weather_date(cells[1])
        if s_date:
            start_dates.append(s_date)
        end_days_last = int(to_float(cells[2], 7) or 7)
        mj = clamp(to_float(cells[4], 12.0) or 12.0, 0.0, 50.0)
        dyn_lag = int(clamp(round(to_float(cells[6], 45) or 45), 1, 180))
        eff = clamp(to_float(cells[7], DEFAULT_BASE_EFF) or DEFAULT_BASE_EFF, 0.1, 30.0)
        obs_kg = to_float(cells[15]) if len(cells) > 15 else None
        t24_wk = to_float(cells[5])

        week_rows.append({"mj": mj, "lag": dyn_lag, "eff": eff})
        calib_rows.append({"obsKg": obs_kg, "obsMass": None, "obsMj": None, "obsT24": t24_wk})

    if start_dates:
        config["start"] = start_dates[0]
        last_dt = parse_date(start_dates[-1]) + timedelta(days=max(0, end_days_last - 1))
        config["harvestEnd"] = last_dt.isoformat()

    return {
        "source": "html_masterplan_csv",
        "zone_label": zone_label,
        "config": config,
        "week_rows": week_rows,
        "env_months": env_months,
        "calib_rows": calib_rows,
        "custom_city": None,
    }


def parse_calib_csv_into_rows(
    calib_csv_path: Path,
    config: Dict[str, Any],
    base_calib_rows: Optional[List[Dict[str, Optional[float]]]] = None,
) -> List[Dict[str, Optional[float]]]:
    """
    HTML parseCalibCsvText() 및 [CSV 양식 받기](토마토_실측데이터_입력양식.csv) 형식 모두 지원:
      - 6컬럼(양식): 주차, 시작일, 실측수확량(kg/m2), 실측단과중(g), 실측외부일사(MJ/m2/일), 실측T24(C)
      - 2~5컬럼(붙여넣기): 주차/날짜, 실측수확량(kg/㎡), 단과중(g), [선택]실측일사, [선택]실측T24
    """
    periods = plan_periods(str(config["start"]), str(config["harvestEnd"]))
    out = _normalize_calib_rows(base_calib_rows, len(periods))
    text = decode_text_auto(calib_csv_path)
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    if not lines:
        raise ValueError(f"[{calib_csv_path.name}] 실측 CSV 내용이 비어 있습니다.")

    is_template_6col = False
    for line_idx, line in enumerate(lines):
        if "\t" in line or ";" in line:
            parts = [p.strip().strip('"') for p in re.split(r"[\t,;]+", line)]
        else:
            parts = [p.strip().strip('"') for p in next(csv.reader([line]))]
        if len(parts) < 2:
            continue

        p0 = parts[0]
        if line_idx == 0:
            joined_head = "".join(parts).lower()
            if "시작일" in joined_head or ("실측수확량" in joined_head and len(parts) >= 6):
                is_template_6col = True
                continue
            if to_float(re.sub(r"[^0-9.-]", "", p0)) is None and not re.match(r"^\d{4}[-./]\d{1,2}[-./]\d{1,2}$", p0):
                continue

        w_idx = -1
        d_norm = normalize_weather_date(p0)
        if d_norm:
            dt = parse_date(d_norm)
            for idx, p in enumerate(periods):
                if p["at"] <= dt < p["at"] + timedelta(days=p["days"]):
                    w_idx = idx
                    break
        else:
            wn_str = re.sub(r"[^0-9-]", "", p0)
            if wn_str:
                wn = int(wn_str)
                if 1 <= wn <= len(periods):
                    w_idx = wn - 1
        if w_idx < 0 or w_idx >= len(periods):
            continue

        # 6컬럼 템플릿(주차, 시작일, obsKg, obsMass, obsMj, obsT24) vs 5컬럼 붙여넣기(주차, obsKg, obsMass, obsMj, obsT24)
        offset = 1 if (is_template_6col or (len(parts) >= 3 and normalize_weather_date(parts[1]) is not None)) else 0
        obs_kg = to_float(parts[1 + offset]) if len(parts) > 1 + offset else None
        obs_mass = to_float(parts[2 + offset]) if len(parts) > 2 + offset else None
        obs_mj = to_float(parts[3 + offset]) if len(parts) > 3 + offset else None
        obs_t24 = to_float(parts[4 + offset]) if len(parts) > 4 + offset else None

        if obs_kg is not None:
            out[w_idx]["obsKg"] = obs_kg
        if obs_mass is not None:
            out[w_idx]["obsMass"] = obs_mass
        if obs_mj is not None:
            out[w_idx]["obsMj"] = obs_mj
        if obs_t24 is not None:
            out[w_idx]["obsT24"] = obs_t24

    return out


# ==============================================================================
# 5-B. 로컬 다중 실측 CSV 폴더 스캔 · 유연한 컬럼 매핑 · 중복 제거 · 병합 로더
# ==============================================================================
CALIB_EXACT_ALIASES: Dict[str, str] = {
    # 구역/작기/실험 식별자
    "zone_id": "zone_id",
    "zoneid": "zone_id",
    "zone": "zone_id",
    "구역": "zone_id",
    "구역id": "zone_id",
    "구역명": "zone_id",
    "작기": "zone_id",
    "season": "zone_id",
    "plot": "zone_id",
    "house": "zone_id",
    "온실": "zone_id",
    "실험구": "zone_id",
    # 주차/시점
    "week_no": "week_no",
    "weekno": "week_no",
    "week": "week_no",
    "주차": "week_no",
    "주": "week_no",
    "wk": "week_no",
    "date": "date",
    "start": "date",
    "startdate": "date",
    "일자": "date",
    "날짜": "date",
    "시작일": "date",
    "수확일": "date",
    # 타깃: 실측 수확량
    "y_actual": "y_actual",
    "yactual": "y_actual",
    "obs_kg": "y_actual",
    "obskg": "y_actual",
    "실측수확량": "y_actual",
    "실측수확량(kg/m2)": "y_actual",
    "실측수확량(kg/㎡)": "y_actual",
    "실측수확": "y_actual",
    "실측수확(kg/㎡)": "y_actual",
    "yield_actual": "y_actual",
    "yieldactual": "y_actual",
    "actual_yield": "y_actual",
    "actualyield": "y_actual",
    "obs_yield": "y_actual",
    "수확량": "y_actual",
    "실측": "y_actual",
    # 실측 보조 관측치
    "obs_mass": "obs_mass",
    "obsmass": "obs_mass",
    "실측단과중": "obs_mass",
    "실측단과중(g)": "obs_mass",
    "단과중": "obs_mass",
    "obs_mj": "mj_out",
    "obsmj": "mj_out",
    "실측외부일사": "mj_out",
    "실측외부일사(mj/m2/일)": "mj_out",
    "실측일사": "mj_out",
    "obs_t24": "t_mean",
    "obst24": "t_mean",
    "실측t24": "t_mean",
    "실측t24(c)": "t_mean",
    "실측t24(℃)": "t_mean",
    # 주요 환경/생육 변수
    "dli_daily": "dli_daily",
    "dlidaily": "dli_daily",
    "dli": "dli_daily",
    "solar_dli": "dli_daily",
    "solardli": "dli_daily",
    "총dli": "dli_daily",
    "평균총dli": "dli_daily",
    "nat_dli": "nat_dli",
    "natdli": "nat_dli",
    "자연광dli": "nat_dli",
    "자연dli": "nat_dli",
    "add_dli": "add_dli",
    "adddli": "add_dli",
    "보광dli": "add_dli",
    "mj_out": "mj_out",
    "mjout": "mj_out",
    "rad_mj": "mj_out",
    "radmj": "mj_out",
    "외부일사": "mj_out",
    "실외일사": "mj_out",
    "실외일사(mj/㎡/일)": "mj_out",
    "eff_w": "eff_w",
    "effw": "eff_w",
    "eff": "eff_w",
    "광효율": "eff_w",
    "주별광효율": "eff_w",
    "t_mean": "t_mean",
    "tmean": "t_mean",
    "t24": "t_mean",
    "t24(℃)": "t_mean",
    "평균온도": "t_mean",
    "일평균온도": "t_mean",
    "평균기온": "t_mean",
    "t_day": "t_day",
    "tday": "t_day",
    "주간온도": "t_day",
    "t_night": "t_night",
    "tnight": "t_night",
    "야간온도": "t_night",
    "dif": "dif",
    "주야간온도차": "dif",
    "vpd_mean": "vpd_mean",
    "vpdmean": "vpd_mean",
    "vpd": "vpd_mean",
    "vpd(kpa)": "vpd_mean",
    "rh_mean": "rh_mean",
    "rhmean": "rh_mean",
    "rh": "rh_mean",
    "상대습도": "rh_mean",
    "상대습도(%)": "rh_mean",
    "습도": "rh_mean",
    "co2_mean": "co2_mean",
    "co2mean": "co2_mean",
    "co2": "co2_mean",
    "co2(ppm)": "co2_mean",
    "rtr": "rtr",
    "k_vpd": "k_vpd",
    "kvpd": "k_vpd",
    "k_temp": "k_temp",
    "ktemp": "k_temp",
    "k_co2": "k_co2",
    "kco2": "k_co2",
    "k_light": "k_light",
    "klight": "k_light",
    "k_total": "k_total",
    "ktotal": "k_total",
    "dyn_lag": "dyn_lag",
    "dynlag": "dyn_lag",
    "lag": "dyn_lag",
    "동적성숙lag": "dyn_lag",
    "sink_part": "sink_part",
    "sinkpart": "sink_part",
    "초기분배율": "sink_part",
    "ss_ratio": "ss_ratio",
    "ssratio": "ss_ratio",
    "fruit_load": "fruit_load",
    "fruitload": "fruit_load",
    "착과부하": "fruit_load",
    "cum_p_n_8w": "cum_p_n_8w",
    "cumpn8w": "cum_p_n_8w",
    "tr_ratio": "tr_ratio",
    "trratio": "tr_ratio",
    "ma_3week": "ma_3week",
    "ma3week": "ma_3week",
    "3주이동평균": "ma_3week",
    "orig_kg_sim": "orig_kg_sim",
    "orig_kg": "orig_kg_sim",
    "원본수확": "orig_kg_sim",
    "adj_kg_sim": "adj_kg_sim",
    "adj_kg": "adj_kg_sim",
    "보정수확": "adj_kg_sim",
}

CALIB_BASE_ALIASES: Dict[str, str] = {
    re.sub(r"\([^)]*\)", "", k): v for k, v in CALIB_EXACT_ALIASES.items()
}

RICH_FEATURE_KEYS = {
    "dli_daily", "nat_dli", "add_dli", "t_day", "t_night", "dif",
    "vpd_mean", "rh_mean", "co2_mean", "cum_p_n_8w", "tr_ratio",
    "fruit_load", "ma_3week", "k_total", "ss_ratio", "orig_kg_sim", "adj_kg_sim",
}


def _normalize_header_token(name: str) -> str:
    return re.sub(r"[\s_\-]+", "", str(name).strip().lstrip("\ufeff").lower())


def _resolve_calib_column_name(header: str) -> Optional[str]:
    exact_tok = _normalize_header_token(header)
    if not exact_tok:
        return None
    if exact_tok in CALIB_EXACT_ALIASES:
        return CALIB_EXACT_ALIASES[exact_tok]
    base_tok = _normalize_header_token(re.sub(r"\([^)]*\)", "", header))
    if base_tok in CALIB_BASE_ALIASES:
        return CALIB_BASE_ALIASES[base_tok]
    return None


def _split_csv_row(line: str) -> List[str]:
    if "\t" in line or (";" in line and "," not in line):
        return [p.strip().strip('"') for p in re.split(r"[\t;]+", line)]
    return [p.strip().strip('"') for p in next(csv.reader([line]))]


def _sanitize_calib_numeric(field: str, val: Optional[float]) -> Optional[float]:
    if val is None or not math.isfinite(val):
        return None
    if field in ("y_actual", "obs_mass", "mj_out", "dli_daily", "nat_dli", "add_dli", "eff_w", "vpd_mean", "orig_kg_sim", "adj_kg_sim", "ma_3week", "cum_p_n_8w", "fruit_load"):
        if val < 0.0:
            return None
    if field == "rh_mean" and not (0.0 <= val <= 100.0):
        return None
    if field == "co2_mean" and val <= 0.0:
        return None
    if field == "t_mean" and not (-20.0 <= val <= 60.0):
        return None
    return float(val)


def _excel_cell_to_str(val: Any) -> str:
    if val is None:
        return ""
    if isinstance(val, datetime):
        return val.date().isoformat()
    if isinstance(val, date):
        return val.isoformat()
    if isinstance(val, float):
        if not math.isfinite(val):
            return ""
        if val.is_integer():
            return str(int(val))
        return str(val)
    s = str(val).strip()
    if s.lower() in ("nan", "nat", "none", "null"):
        return ""
    return s


def _trim_grid_rows(raw_grid: List[List[Any]]) -> List[List[str]]:
    cleaned: List[List[str]] = []
    for row in raw_grid:
        cells = [_excel_cell_to_str(c) for c in row]
        while cells and cells[-1] == "":
            cells.pop()
        if any(c != "" for c in cells):
            cleaned.append(cells)
    return cleaned


def _read_xlsx_with_openpyxl(excel_path: Path) -> List[Tuple[str, List[List[str]]]]:
    if openpyxl is None:
        raise ImportError("openpyxl 패키지가 설치되어 있지 않습니다.")
    wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    sheets: List[Tuple[str, List[List[str]]]] = []
    try:
        for name in wb.sheetnames:
            ws = wb[name]
            raw_rows = [list(r) for r in ws.iter_rows(values_only=True)]
            grid = _trim_grid_rows(raw_rows)
            if grid:
                sheets.append((name, grid))
    finally:
        wb.close()
    return sheets


def _col_letters_to_idx(ref: str) -> int:
    m = re.match(r"^([A-Za-z]+)", ref)
    if not m:
        return 0
    idx = 0
    for ch in m.group(1).upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return max(0, idx - 1)


def _read_xlsx_stdlib_fallback(excel_path: Path) -> List[Tuple[str, List[List[str]]]]:
    """openpyxl 미설치 환경에서도 .xlsx(OOXML)를 읽기 위한 경량 표준 라이브러리 폴백"""
    ns_main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ns_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    sheets_out: List[Tuple[str, List[List[str]]]] = []

    with zipfile.ZipFile(excel_path, "r") as zf:
        shared_strings: List[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            ss_root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in ss_root.findall(f"{{{ns_main}}}si"):
                txt = "".join(t.text or "" for t in si.iter(f"{{{ns_main}}}t"))
                shared_strings.append(txt)

        wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
        rels_map: Dict[str, str] = {}
        if "xl/_rels/workbook.xml.rels" in zf.namelist():
            rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
            for rel in rels_root:
                rid = rel.attrib.get("Id", "")
                target = rel.attrib.get("Target", "").lstrip("/")
                if target and not target.startswith("xl/"):
                    target = f"xl/{target}"
                rels_map[rid] = target

        sheet_entries: List[Tuple[str, str]] = []
        for idx, sh in enumerate(wb_root.findall(f".//{{{ns_main}}}sheet"), start=1):
            name = sh.attrib.get("name", f"Sheet{idx}")
            rid = sh.attrib.get(f"{{{ns_rel}}}id", "")
            target = rels_map.get(rid, f"xl/worksheets/sheet{idx}.xml")
            if target in zf.namelist():
                sheet_entries.append((name, target))

        for name, target in sheet_entries:
            ws_root = ET.fromstring(zf.read(target))
            rows_parsed: List[List[str]] = []
            for row_el in ws_root.findall(f".//{{{ns_main}}}sheetData/{{{ns_main}}}row"):
                row_vals: Dict[int, str] = {}
                max_col = -1
                seq_col = 0
                for c_el in row_el.findall(f"{{{ns_main}}}c"):
                    ref = c_el.attrib.get("r", "")
                    col_idx = _col_letters_to_idx(ref) if ref else seq_col
                    seq_col = col_idx + 1
                    c_type = c_el.attrib.get("t", "")
                    val_str = ""
                    if c_type == "inlineStr":
                        val_str = "".join(t.text or "" for t in c_el.iter(f"{{{ns_main}}}t"))
                    else:
                        v_el = c_el.find(f"{{{ns_main}}}v")
                        if v_el is not None and v_el.text is not None:
                            if c_type == "s":
                                s_idx = int(v_el.text)
                                val_str = shared_strings[s_idx] if 0 <= s_idx < len(shared_strings) else ""
                            else:
                                val_str = v_el.text
                    row_vals[col_idx] = val_str.strip()
                    if col_idx > max_col:
                        max_col = col_idx
                if max_col >= 0:
                    rows_parsed.append([row_vals.get(i, "") for i in range(max_col + 1)])
            grid = _trim_grid_rows(rows_parsed)
            if grid:
                sheets_out.append((name, grid))
    return sheets_out


def _read_xls_with_xlrd(excel_path: Path) -> List[Tuple[str, List[List[str]]]]:
    if xlrd is None:
        raise ImportError("'.xls' 구형 엑셀 파일을 읽으려면 xlrd>=2.0.1 패키지가 필요합니다.")
    wb = xlrd.open_workbook(str(excel_path))
    sheets: List[Tuple[str, List[List[str]]]] = []
    for sheet in wb.sheets():
        raw_rows: List[List[Any]] = []
        for r_idx in range(sheet.nrows):
            row_vals: List[Any] = []
            for c_idx in range(sheet.ncols):
                cell = sheet.cell(r_idx, c_idx)
                val = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        val = xlrd.xldate_as_datetime(val, wb.datemode).date().isoformat()
                    except Exception:
                        pass
                row_vals.append(val)
            raw_rows.append(row_vals)
        grid = _trim_grid_rows(raw_rows)
        if grid:
            sheets.append((sheet.name, grid))
    return sheets


def _read_excel_sheets(excel_path: Path) -> List[Tuple[str, List[List[str]]]]:
    ext = excel_path.suffix.lower()
    if ext == ".xls":
        return _read_xls_with_xlrd(excel_path)
    if openpyxl is not None:
        return _read_xlsx_with_openpyxl(excel_path)
    if pd is not None:
        xls = pd.ExcelFile(excel_path)
        out: List[Tuple[str, List[List[str]]]] = []
        for name in xls.sheet_names:
            df = xls.parse(name, header=None, dtype=object)
            grid = _trim_grid_rows(df.values.tolist())
            if grid:
                out.append((str(name), grid))
        return out
    return _read_xlsx_stdlib_fallback(excel_path)


def _parse_masterplan_grid(
    grid: List[List[str]],
    src_name: str,
    default_zone: str = "Zone_A",
) -> List[Dict[str, Any]]:
    weekly_header_idx = -1
    zone_label = default_zone
    start_str = DEFAULT_CONFIG["start"]

    for idx, cells in enumerate(grid):
        if not cells:
            continue
        if cells[0] == "지역" and len(cells) >= 4:
            city_name = cells[1] or "Unknown"
            scen_name = cells[3] if len(cells) > 3 else ""
            zone_label = f"{city_name} ({scen_name})" if scen_name else city_name
            crop_span = cells[5] if len(cells) > 5 else ""
            m_span = re.match(r"(\d{4}-\d{2}-\d{2})\s*~\s*(\d{4}-\d{2}-\d{2})", crop_span)
            if m_span:
                start_str = m_span.group(1)
        elif cells[0].startswith("[주별 세부 광량"):
            weekly_header_idx = idx + 1
            break

    if weekly_header_idx < 0:
        return []

    records: List[Dict[str, Any]] = []
    for cells in grid[weekly_header_idx + 1 :]:
        if len(cells) < 16:
            continue
        wk_no = to_float(cells[0])
        if wk_no is None:
            continue
        y_act = _sanitize_calib_numeric("y_actual", to_float(cells[15]))
        if y_act is None:
            continue
        s_date = normalize_weather_date(cells[1])
        if not s_date:
            s_date = (parse_date(str(start_str)) + timedelta(days=(int(wk_no) - 1) * 7)).isoformat()
        records.append({
            "zone_id": zone_label or default_zone,
            "week_no": int(round(wk_no)),
            "date": s_date,
            "y_actual": y_act,
            "mj_out": _sanitize_calib_numeric("mj_out", to_float(cells[4])),
            "t_mean": _sanitize_calib_numeric("t_mean", to_float(cells[5])),
            "dyn_lag": _sanitize_calib_numeric("dyn_lag", to_float(cells[6])),
            "eff_w": _sanitize_calib_numeric("eff_w", to_float(cells[7])),
            "_source_file": src_name,
            "_has_rich_features": False,
        })
    return records


def _parse_flexible_calib_grid(
    grid: List[List[str]],
    src_name: str,
    default_zone: str = "Zone_A",
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    stats = {"source_file": src_name, "raw_rows": 0, "skipped_invalid": 0, "kind": "unknown"}
    if not grid:
        return [], stats

    head_preview = "\n".join(",".join(r) for r in grid[:35])
    if "[주별 세부 광량" in head_preview and "온실환경연구실 토마토" in head_preview:
        mp_records = _parse_masterplan_grid(grid, src_name=src_name, default_zone=default_zone)
        if mp_records:
            stats["kind"] = "masterplan_table"
            stats["raw_rows"] = len(mp_records)
            return mp_records, stats

    # Open-Meteo / ASOS 기상 관측 파일은 실측 수확량 폴더에서 오인 파싱되지 않도록 스킵
    first_cells = grid[0]
    first_cols_joined = ",".join(first_cells).lower()
    first_line_text = " ".join(first_cells)
    if ("latitude" in first_cols_joined and "longitude" in first_cols_joined) or (
        ("지점명" in first_line_text or "shortwave_radiation_sum" in first_cols_joined)
        and "수확" not in head_preview
        and "y_actual" not in head_preview.lower()
        and "obs_kg" not in head_preview.lower()
    ):
        stats["kind"] = "weather_file_skipped"
        return [], stats

    mapped_headers: List[Optional[str]] = [_resolve_calib_column_name(c) for c in first_cells]
    has_header = any(m is not None for m in mapped_headers)

    records_out: List[Dict[str, Any]] = []
    start_row_idx = 1 if has_header else 0
    if not has_header:
        p0 = first_cells[0] if first_cells else ""
        if to_float(re.sub(r"[^0-9.-]", "", p0)) is None and normalize_weather_date(p0) is None:
            start_row_idx = 1

    mapped_rich = set(m for m in mapped_headers if m in RICH_FEATURE_KEYS)
    stats["kind"] = "header_mapped" if has_header else "positional"

    for cells in grid[start_row_idx:]:
        if not cells or all(c == "" for c in cells):
            continue
        stats["raw_rows"] += 1
        rec: Dict[str, Any] = {
            "zone_id": None,
            "week_no": None,
            "date": None,
            "y_actual": None,
            "_source_file": src_name,
            "_has_rich_features": bool(mapped_rich),
        }

        if has_header:
            for col_idx, canon in enumerate(mapped_headers):
                if canon is None or col_idx >= len(cells):
                    continue
                raw_val = cells[col_idx].strip()
                if not raw_val:
                    continue
                if canon == "zone_id":
                    rec["zone_id"] = raw_val
                elif canon == "date":
                    d_norm = normalize_weather_date(raw_val)
                    if d_norm:
                        rec["date"] = d_norm
                elif canon == "week_no":
                    wn_num = to_float(re.sub(r"[^0-9.-]", "", raw_val))
                    if wn_num is not None and int(round(wn_num)) >= 1:
                        rec["week_no"] = int(round(wn_num))
                else:
                    num_val = _sanitize_calib_numeric(canon, to_float(raw_val))
                    if num_val is not None:
                        rec[canon] = num_val
        else:
            if len(cells) < 2:
                stats["skipped_invalid"] += 1
                continue
            p0 = cells[0]
            d_norm0 = normalize_weather_date(p0)
            if d_norm0:
                rec["date"] = d_norm0
            else:
                wn_str = re.sub(r"[^0-9-]", "", p0)
                wn_val = to_float(wn_str)
                if wn_val is not None and int(round(wn_val)) >= 1:
                    rec["week_no"] = int(round(wn_val))
            offset = 1 if (len(cells) >= 3 and normalize_weather_date(cells[1]) is not None) else 0
            if offset == 1 and rec["date"] is None:
                rec["date"] = normalize_weather_date(cells[1])
            pos_fields = ("y_actual", "obs_mass", "mj_out", "t_mean")
            for k_offset, fname in enumerate(pos_fields):
                c_idx = 1 + offset + k_offset
                if c_idx < len(cells):
                    val = _sanitize_calib_numeric(fname, to_float(cells[c_idx]))
                    if val is not None:
                        rec[fname] = val

        if rec["week_no"] is None and rec["date"] is None:
            stats["skipped_invalid"] += 1
            continue
        has_useful_val = any(
            rec.get(k) is not None
            for k in ("y_actual", "obs_mass", "mj_out", "t_mean", "dli_daily", "adj_kg_sim", "orig_kg_sim")
        )
        if not has_useful_val:
            stats["skipped_invalid"] += 1
            continue

        if not rec["zone_id"]:
            rec["zone_id"] = default_zone
        records_out.append(rec)

    return records_out, stats


def parse_flexible_calib_csv(
    csv_path: Path,
    source_label: Optional[str] = None,
    default_zone: str = "Zone_A",
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """
    다중 실측 CSV 폴더 내 개별 CSV 파일을 유연하게 파싱한다.
    지원 형식:
      1) 헤더가 있는 실측/환경 CSV (영문/한글 별칭 컬럼 자동 매핑)
      2) HTML [CSV 양식 받기] 6컬럼 및 붙여넣기 2~5컬럼(헤더 유무 모두 지원)
      3) HTML [마스터플랜 CSV 다운로드] 파일
    기상청 ASOS / Open-Meteo 전용 파일처럼 수확 실측·주차 컬럼이 없는 파일은 안전하게 건너뛴다.
    """
    src_name = source_label or csv_path.name
    text = decode_text_auto(csv_path)
    raw_lines = [ln for ln in text.splitlines() if ln.strip()]
    if not raw_lines:
        return [], {"source_file": src_name, "raw_rows": 0, "skipped_invalid": 0, "kind": "empty"}

    grid = [_split_csv_row(ln) for ln in raw_lines]
    return _parse_flexible_calib_grid(grid, src_name=src_name, default_zone=default_zone)


def parse_flexible_calib_excel(
    excel_path: Path,
    source_label: Optional[str] = None,
    default_zone: str = "Zone_A",
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """
    다중 시트 엑셀(.xlsx, .xls) 파일을 읽어 첫 번째 시트부터 순서대로 검사해
    유효 데이터가 포함된 시트를 안전하게 파싱한다.
    """
    src_name = source_label or excel_path.name
    sheets = _read_excel_sheets(excel_path)
    if not sheets:
        return [], {"source_file": src_name, "raw_rows": 0, "skipped_invalid": 0, "kind": "excel_empty"}

    last_stats = {"source_file": src_name, "raw_rows": 0, "skipped_invalid": 0, "kind": "excel_no_valid_sheet"}
    for sheet_idx, (sheet_name, grid) in enumerate(sheets):
        sheet_label = src_name if sheet_idx == 0 else f"{src_name}#{sheet_name}"
        recs, st = _parse_flexible_calib_grid(grid, src_name=sheet_label, default_zone=default_zone)
        last_stats = st
        if recs:
            st["sheet_name"] = sheet_name
            st["kind"] = f"excel_{st['kind']}"
            return recs, st
        if st.get("kind") == "weather_file_skipped" and len(sheets) == 1:
            return [], st

    return [], last_stats


def parse_flexible_calib_file(
    file_path: Path,
    source_label: Optional[str] = None,
    default_zone: str = "Zone_A",
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    ext = file_path.suffix.lower()
    if ext in (".xlsx", ".xlsm", ".xls"):
        return parse_flexible_calib_excel(file_path, source_label=source_label, default_zone=default_zone)
    return parse_flexible_calib_csv(file_path, source_label=source_label, default_zone=default_zone)


def merge_and_deduplicate_calib_records(
    records: List[Dict[str, Any]],
) -> Tuple[List[Dict[str, Any]], int]:
    """
    다중 CSV/Excel에서 수집된 레코드를 (zone_id, date 또는 week_no) 키로 중복 정리한다.
    정렬 순서(파일명 순 -> 행 순) 기준 뒤쪽(최신) 파일 값으로 병합/갱신한다.
    """
    merged_map: Dict[Tuple[str, str], Dict[str, Any]] = {}
    ordered_keys: List[Tuple[str, str]] = []
    duplicate_count = 0

    for rec in records:
        z_key = str(rec.get("zone_id") or "Zone_A").strip()
        if rec.get("date"):
            t_key = f"d:{rec['date']}"
        elif rec.get("week_no") is not None:
            t_key = f"w:{int(rec['week_no'])}"
        else:
            continue

        key = (z_key, t_key)
        if key not in merged_map:
            merged_map[key] = dict(rec)
            ordered_keys.append(key)
        else:
            duplicate_count += 1
            existing = merged_map[key]
            for fld, val in rec.items():
                if fld.startswith("_"):
                    continue
                if val is not None:
                    existing[fld] = val
            existing["_source_file"] = rec.get("_source_file", existing.get("_source_file"))
            existing["_has_rich_features"] = bool(existing.get("_has_rich_features") or rec.get("_has_rich_features"))

    merged_list = [merged_map[k] for k in ordered_keys]

    # 구역별로 날짜/주차 오름차순 정렬 및 누락 week_no 보완
    by_zone: Dict[str, List[Dict[str, Any]]] = {}
    for item in merged_list:
        by_zone.setdefault(str(item["zone_id"]), []).append(item)

    final_sorted: List[Dict[str, Any]] = []
    for z_name in sorted(by_zone.keys()):
        z_items = by_zone[z_name]
        z_items.sort(key=lambda r: (
            str(r.get("date") or "9999-99-99"),
            int(r.get("week_no") if r.get("week_no") is not None else 10**9),
        ))
        dated_vals = [parse_date(r["date"]) for r in z_items if r.get("date")]
        min_dt = min(dated_vals) if dated_vals else None
        for seq_idx, r in enumerate(z_items):
            if r.get("week_no") is None:
                if r.get("date") and min_dt is not None:
                    r["week_no"] = ((parse_date(r["date"]) - min_dt).days // 7) + 1
                else:
                    r["week_no"] = seq_idx + 1
            final_sorted.append(r)

    return final_sorted, duplicate_count


def load_calib_dir_records(
    calib_dir: Path,
    recursive: bool = False,
    default_zone: str = "Zone_A",
) -> Dict[str, Any]:
    """지정 폴더(기본: data/calib_csv/)의 모든 CSV 및 엑셀(.xlsx/.xls)을 정렬 순서로 스캔·병합한다."""
    if not calib_dir.exists() or not calib_dir.is_dir():
        return {
            "exists": False,
            "calib_dir": str(calib_dir),
            "files_found": [],
            "files_used": [],
            "csv_count": 0,
            "excel_count": 0,
            "records": [],
            "raw_rows": 0,
            "merged_rows": 0,
            "duplicates_merged": 0,
            "valid_obs_count": 0,
            "zones": [],
        }

    pattern_iter = calib_dir.rglob("*") if recursive else calib_dir.iterdir()
    data_files = sorted(
        [
            p for p in pattern_iter
            if p.is_file()
            and not p.name.startswith("~$")
            and p.suffix.lower() in SUPPORTED_CALIB_EXTS
        ],
        key=lambda p: p.relative_to(calib_dir).as_posix().lower(),
    )

    all_raw_records: List[Dict[str, Any]] = []
    files_used: List[str] = []
    csv_count = 0
    excel_count = 0
    raw_rows_total = 0

    for p in data_files:
        rel_name = p.relative_to(calib_dir).as_posix()
        try:
            recs, st = parse_flexible_calib_file(p, source_label=rel_name, default_zone=default_zone)
            raw_rows_total += st["raw_rows"]
            if recs:
                files_used.append(rel_name)
                if p.suffix.lower() == ".csv":
                    csv_count += 1
                else:
                    excel_count += 1
                all_raw_records.extend(recs)
        except Exception as exc:
            print(f"[경고] '{rel_name}' 파싱 중 오류가 발생하여 건너뜁니다: {exc}")

    merged_records, dup_cnt = merge_and_deduplicate_calib_records(all_raw_records)
    valid_obs_cnt = sum(1 for r in merged_records if r.get("y_actual") is not None)
    zones = sorted({str(r["zone_id"]) for r in merged_records if r.get("zone_id")})

    return {
        "exists": True,
        "calib_dir": str(calib_dir),
        "files_found": [p.relative_to(calib_dir).as_posix() for p in data_files],
        "files_used": files_used,
        "csv_count": csv_count,
        "excel_count": excel_count,
        "records": merged_records,
        "raw_rows": raw_rows_total,
        "merged_rows": len(merged_records),
        "duplicates_merged": dup_cnt,
        "valid_obs_count": valid_obs_cnt,
        "zones": zones,
    }


def apply_calib_records_to_periods(
    records: List[Dict[str, Any]],
    config: Dict[str, Any],
    base_calib_rows: Optional[List[Dict[str, Optional[float]]]] = None,
) -> Tuple[List[Dict[str, Optional[float]]], List[Dict[str, Any]]]:
    """
    수집·병합된 실측 레코드 중 현재 시뮬레이션 작기(periods)에 매칭되는 행을 calib_rows에 주입하고,
    매칭되지 않은 나머지 타 구역/타 시점 레코드는 별도 리스트(unmatched)로 반환한다.
    """
    periods = plan_periods(str(config["start"]), str(config["harvestEnd"]))
    out = _normalize_calib_rows(base_calib_rows, len(periods))
    unmatched: List[Dict[str, Any]] = []

    for rec in records:
        w_idx = -1
        if rec.get("date"):
            dt = parse_date(str(rec["date"]))
            for idx, p in enumerate(periods):
                if p["at"] <= dt < p["at"] + timedelta(days=p["days"]):
                    w_idx = idx
                    break
        elif rec.get("week_no") is not None:
            wn = int(rec["week_no"])
            if 1 <= wn <= len(periods):
                w_idx = wn - 1

        if w_idx < 0 or w_idx >= len(periods):
            unmatched.append(rec)
            continue

        if rec.get("y_actual") is not None:
            out[w_idx]["obsKg"] = float(rec["y_actual"])
        if rec.get("obs_mass") is not None:
            out[w_idx]["obsMass"] = float(rec["obs_mass"])
        if rec.get("mj_out") is not None:
            out[w_idx]["obsMj"] = float(rec["mj_out"])
        if rec.get("t_mean") is not None:
            out[w_idx]["obsT24"] = float(rec["t_mean"])

    return out, unmatched


def build_tabular_ml_rows(
    records: List[Dict[str, Any]],
    default_start_date: str = "2027-01-01",
    weather_ds: Optional[Dict[str, Any]] = None,
) -> List[Dict[str, Any]]:
    """
    다중 실측 CSV(환경·수확량 로그) 레코드를 구역(zone_id)별로 그룹핑하여
    HTML 27개 피처(calc_physiology · 이동평균 ma_3week · 누적광합성 cum_p_n_8w) 규격으로 변환한다.
    """
    by_zone: Dict[str, List[Dict[str, Any]]] = {}
    for rec in records:
        z_name = str(rec.get("zone_id") or "Zone_A")
        by_zone.setdefault(z_name, []).append(rec)

    all_out: List[Dict[str, Any]] = []
    for z_name in sorted(by_zone.keys()):
        z_recs = sorted(
            by_zone[z_name],
            key=lambda r: (str(r.get("date") or "9999-99-99"), int(r.get("week_no") or 0)),
        )
        start_dt = parse_date(str(z_recs[0]["date"])) if z_recs and z_recs[0].get("date") else parse_date(default_start_date)
        y_hist: List[float] = []
        pn_hist: List[float] = []

        for idx, r in enumerate(z_recs):
            wk = int(r.get("week_no") or (idx + 1))
            wk_date = str(r.get("date") or (start_dt + timedelta(days=max(wk - 1, 0) * 7)).isoformat())
            dt_obj = parse_date(wk_date)
            m_idx = dt_obj.month - 1
            env_fb = STANDARD_ENV_MONTHS[m_idx] if 0 <= m_idx < 12 else STANDARD_ENV_MONTHS[0]

            mj_out = to_float(r.get("mj_out"))
            if mj_out is None and weather_ds is not None:
                if wk_date in weather_ds["daily"]:
                    mj_out = weather_ds["daily"][wk_date]
                elif wk_date[5:] in weather_ds["calFallback"]:
                    mj_out = weather_ds["calFallback"][wk_date[5:]]
                else:
                    mj_out = weather_ds["monthly"][m_idx]

            dli_daily = to_float(r.get("dli_daily"))
            if dli_daily is None and mj_out is not None:
                dli_daily = mj_out * PAR_FACTOR_DEFAULT * 0.70
            if dli_daily is None:
                mj_fb = DEFAULT_MONTHLY_MJ[m_idx] if 0 <= m_idx < 12 else 12.0
                mj_out = mj_fb
                dli_daily = mj_fb * PAR_FACTOR_DEFAULT * 0.70
            if mj_out is None:
                mj_out = dli_daily / (PAR_FACTOR_DEFAULT * 0.70)

            t_mean = to_float(r.get("t_mean"), env_fb["temp"]) or env_fb["temp"]
            t_day = to_float(r.get("t_day"), round(t_mean + 2.25, 1)) or round(t_mean + 2.25, 1)
            t_night = to_float(r.get("t_night"), round(t_mean - 2.25, 1)) or round(t_mean - 2.25, 1)
            dif = to_float(r.get("dif"), t_day - t_night) or (t_day - t_night)
            rh_mean = to_float(r.get("rh_mean"), env_fb["rh"]) or env_fb["rh"]
            co2_mean = to_float(r.get("co2_mean"), env_fb["co2"]) or env_fb["co2"]

            phys = calc_physiology(t_mean, rh_mean, co2_mean, dif=dif, dli=dli_daily)
            vpd_mean = to_float(r.get("vpd_mean"), phys["vpd"]) or phys["vpd"]
            eff_w = to_float(r.get("eff_w"), DEFAULT_BASE_EFF) or DEFAULT_BASE_EFF
            sink_part = to_float(r.get("sink_part"), 1.0) or 1.0
            orig_sim = to_float(r.get("orig_kg_sim"), dli_daily * 7.0 * (eff_w / 1000.0) * sink_part) or 0.0
            adj_sim = to_float(r.get("adj_kg_sim"), orig_sim * phys["ktotal"]) or 0.0

            y_act = to_float(r.get("y_actual"))
            ma_3 = to_float(r.get("ma_3week"))
            if ma_3 is None:
                ma_3 = sum(y_hist[-3:]) / len(y_hist[-3:]) if y_hist else float(adj_sim)

            pn_hist.append(dli_daily * 0.65)
            cum_pn_8w = to_float(r.get("cum_p_n_8w"), sum(pn_hist[-8:])) or sum(pn_hist[-8:])
            y_hist.append(y_act if y_act is not None else float(adj_sim))

            tr_default = clamp(0.75 + (vpd_mean - 0.95) * 0.08, 0.50, 1.00)
            tr_ratio = to_float(r.get("tr_ratio"), tr_default) or tr_default

            all_out.append({
                "zone_id": z_name,
                "_source_file": r.get("_source_file", ""),
                "week": wk,
                "week_no": wk,
                "start": wk_date,
                "date": wk_date,
                "y_actual": y_act,
                "orig_kg_sim": float(orig_sim),
                "adj_kg_sim": float(adj_sim),
                "has_harvest_potential": True,
                "ma_3week": float(ma_3),
                "dli_daily": float(dli_daily),
                "nat_dli": float(to_float(r.get("nat_dli"), dli_daily) or dli_daily),
                "add_dli": float(to_float(r.get("add_dli"), 0.0) or 0.0),
                "mj_out": float(mj_out),
                "eff_w": float(eff_w),
                "t_day": float(t_day),
                "t_night": float(t_night),
                "t_mean": float(t_mean),
                "dif": float(dif),
                "rh_mean": float(rh_mean),
                "vpd_mean": float(vpd_mean),
                "co2_mean": float(co2_mean),
                "rtr": float(to_float(r.get("rtr"), phys["rtr"]) or phys["rtr"]),
                "k_vpd": float(to_float(r.get("k_vpd"), phys["kvpd"]) or phys["kvpd"]),
                "k_temp": float(to_float(r.get("k_temp"), phys["ktemp"]) or phys["ktemp"]),
                "k_co2": float(to_float(r.get("k_co2"), phys["kco2"]) or phys["kco2"]),
                "k_light": float(to_float(r.get("k_light"), phys["klight"]) or phys["klight"]),
                "k_total": float(to_float(r.get("k_total"), phys["ktotal"]) or phys["ktotal"]),
                "dyn_lag": float(to_float(r.get("dyn_lag"), 45.0) or 45.0),
                "sink_part": float(sink_part),
                "ss_ratio": float(to_float(r.get("ss_ratio"), 1.0) or 1.0),
                "fruit_load": float(to_float(r.get("fruit_load"), 1.0) or 1.0),
                "cum_p_n_8w": float(cum_pn_8w),
                "tr_ratio": float(tr_ratio),
            })

    return all_out


# ==============================================================================
# 6. HTML 시뮬레이터 결과 -> ML 피처 테이블 구축 (+ 기존 SQLite/CSV 폴백 지원)
# ==============================================================================
def build_ml_rows_from_simulation(
    sim: Dict[str, Any],
    calib_rows: List[Dict[str, Optional[float]]],
) -> List[Dict[str, Any]]:
    harvest = sim["harvest"]
    obs_history: List[float] = []
    pn_history: List[float] = []
    rows_ml: List[Dict[str, Any]] = []

    for hi, h in enumerate(harvest):
        ob = calib_rows[hi] if hi < len(calib_rows) else {}
        y_actual = to_float(ob.get("obsKg")) if isinstance(ob, dict) else None

        ma_3 = sum(obs_history[-3:]) / len(obs_history[-3:]) if obs_history else float(h["adjKg"])
        pn_week = float(h["dli_daily"]) * 0.65
        pn_history.append(pn_week)
        cum_pn_8w = sum(pn_history[-8:])

        # 과거 실측치가 있으면 이동평균 이력에 반영, 없으면 물리 모델 기반 값으로 연속성 유지
        obs_history.append(y_actual if y_actual is not None else float(h["adjKg"]))

        has_harvest_potential = (float(h["adjKg"]) > 1e-6) or (float(h["origKg"]) > 1e-6) or (y_actual is not None)

        rows_ml.append({
            "week": int(h["week"]),
            "week_no": int(h["week"]),
            "start": h["at"].isoformat(),
            "date": h["at"].isoformat(),
            "y_actual": y_actual,
            "orig_kg_sim": float(h["origKg"]),
            "adj_kg_sim": float(h["adjKg"]),
            "has_harvest_potential": has_harvest_potential,
            # HTML-동기화 피처
            "ma_3week": float(ma_3),
            "dli_daily": float(h["dli_daily"]),
            "nat_dli": float(h["nat_dli"]),
            "add_dli": float(h["add_dli"]),
            "mj_out": float(h["mj_out"]),
            "eff_w": float(h["eff_w"]),
            "t_day": float(h["t_day"]),
            "t_night": float(h["t_night"]),
            "t_mean": float(h["t_mean"]),
            "dif": float(h["dif"]),
            "rh_mean": float(h["rh_mean"]),
            "vpd_mean": float(h["vpd_mean"]),
            "co2_mean": float(h["co2_mean"]),
            "rtr": float(h["rtr"]),
            "k_vpd": float(h["k_vpd"]),
            "k_temp": float(h["k_temp"]),
            "k_co2": float(h["k_co2"]),
            "k_light": float(h["k_light"]),
            "k_total": float(h["k_total"]),
            "dyn_lag": float(h["avgLag"]),
            "sink_part": float(h["avgSink"]),
            "ss_ratio": float(h["ss_ratio"]),
            "fruit_load": float(h["fruit_load"]),
            "cum_p_n_8w": float(cum_pn_8w),
            "tr_ratio": float(clamp(0.75 + (float(h["vpd_mean"]) - 0.95) * 0.08, 0.50, 1.00)),
        })

    return rows_ml


def load_legacy_tabular_records(
    db_path: str = "smartfarm_yield_swing.db",
    csv_path: str = "raw_yield_source.csv",
    zone_id: str = "Zone_A",
    season_start_date: str = "2026-09-01",
) -> Tuple[List[Dict[str, Any]], str]:
    """기존 SQLite DB 또는 raw_yield_source.csv를 로드하고 누락된 HTML 생리 피처를 보완"""
    raw_records: List[Dict[str, Any]] = []
    src_name = ""

    if Path(db_path).exists():
        conn = sqlite3.connect(db_path)
        try:
            conn.row_factory = sqlite3.Row
            cur = conn.execute("SELECT * FROM weekly_yield_data")
            raw_records = [dict(r) for r in cur.fetchall()]
            src_name = f"sqlite:{db_path}"
        finally:
            conn.close()
    elif Path(csv_path).exists():
        text = decode_text_auto(Path(csv_path))
        reader = csv.DictReader(text.splitlines())
        raw_records = [dict(r) for r in reader]
        src_name = f"csv:{csv_path}"
    else:
        raise FileNotFoundError(
            "입력 데이터를 찾을 수 없습니다. HTML에서 내보낸 JSON/CSV(--plan, --masterplan-csv, --calib-csv) "
            f"또는 '{db_path}' / '{csv_path}' 파일을 준비해 주세요."
        )

    zone_rows = [r for r in raw_records if str(r.get("zone_id", zone_id)) == zone_id]
    if not zone_rows and raw_records:
        zone_rows = raw_records
    zone_rows.sort(key=lambda r: int(to_float(r.get("week_no") or r.get("week"), 0) or 0))

    start_dt = parse_date(season_start_date)
    y_hist: List[float] = []
    pn_hist: List[float] = []
    out: List[Dict[str, Any]] = []

    for idx, r in enumerate(zone_rows):
        wk = int(to_float(r.get("week_no") or r.get("week"), idx + 1) or (idx + 1))
        wk_date = (start_dt + timedelta(days=max(wk - 1, 0) * 7)).isoformat()
        y_act = to_float(r.get(TARGET) or r.get("obsKg"))

        mj_out = to_float(r.get("rad_mj") or r.get("mj_out"))
        dli_daily = to_float(r.get("dli_daily") or r.get("dli") or r.get("solar_dli"))
        if dli_daily is None and mj_out is not None:
            dli_daily = mj_out * PAR_FACTOR_DEFAULT * 0.70
        if dli_daily is None:
            dli_daily = 18.0
        if mj_out is None:
            mj_out = dli_daily / (PAR_FACTOR_DEFAULT * 0.70)

        t_mean = to_float(r.get("t_mean") or r.get("T24"), 21.5) or 21.5
        t_day = to_float(r.get("t_day"), round(t_mean + 2.25, 1)) or round(t_mean + 2.25, 1)
        t_night = to_float(r.get("t_night"), round(t_mean - 2.25, 1)) or round(t_mean - 2.25, 1)
        dif = to_float(r.get("dif"), t_day - t_night) or (t_day - t_night)
        rh_mean = to_float(r.get("rh_mean") or r.get("rh"), 72.0) or 72.0
        co2_mean = to_float(r.get("co2_mean") or r.get("co2"), 600.0) or 600.0

        phys = calc_physiology(t_mean, rh_mean, co2_mean, dif=dif, dli=dli_daily)
        vpd_mean = to_float(r.get("vpd_mean"), phys["vpd"]) or phys["vpd"]
        eff_w = to_float(r.get("eff_w"), DEFAULT_BASE_EFF) or DEFAULT_BASE_EFF
        sink_part = to_float(r.get("sink_part"), 1.0) or 1.0
        orig_sim = to_float(r.get("orig_kg_sim"), dli_daily * 7.0 * (eff_w / 1000.0) * sink_part) or 0.0
        adj_sim = to_float(r.get("adj_kg_sim"), orig_sim * phys["ktotal"]) or 0.0

        ma_3 = to_float(r.get("ma_3week"))
        if ma_3 is None:
            ma_3 = sum(y_hist[-3:]) / len(y_hist[-3:]) if y_hist else 0.0

        pn_hist.append(dli_daily * 0.65)
        cum_pn_8w = to_float(r.get("cum_p_n_8w"), sum(pn_hist[-8:])) or sum(pn_hist[-8:])
        if y_act is not None:
            y_hist.append(y_act)

        out.append({
            "week": wk,
            "week_no": wk,
            "start": str(r.get("date") or wk_date),
            "date": str(r.get("date") or wk_date),
            "y_actual": y_act,
            "orig_kg_sim": float(orig_sim),
            "adj_kg_sim": float(adj_sim),
            "has_harvest_potential": True,
            "ma_3week": float(ma_3),
            "dli_daily": float(dli_daily),
            "nat_dli": float(to_float(r.get("nat_dli"), dli_daily) or dli_daily),
            "add_dli": float(to_float(r.get("add_dli"), 0.0) or 0.0),
            "mj_out": float(mj_out),
            "eff_w": float(eff_w),
            "t_day": float(t_day),
            "t_night": float(t_night),
            "t_mean": float(t_mean),
            "dif": float(dif),
            "rh_mean": float(rh_mean),
            "vpd_mean": float(vpd_mean),
            "co2_mean": float(co2_mean),
            "rtr": float(phys["rtr"]),
            "k_vpd": float(phys["kvpd"]),
            "k_temp": float(phys["ktemp"]),
            "k_co2": float(phys["kco2"]),
            "k_light": float(phys["klight"]),
            "k_total": float(phys["ktotal"]),
            "dyn_lag": float(to_float(r.get("dyn_lag"), 45.0) or 45.0),
            "sink_part": float(sink_part),
            "ss_ratio": float(to_float(r.get("ss_ratio"), 1.0) or 1.0),
            "fruit_load": float(to_float(r.get("fruit_load"), 1.0) or 1.0),
            "cum_p_n_8w": float(cum_pn_8w),
            "tr_ratio": float(to_float(r.get("tr_ratio"), 0.75) or 0.75),
        })

    return out, src_name


# ==============================================================================
# 7. 머신러닝 모델 (scikit-learn RandomForest 우선 · 미설치 환경용 순수 Python 폴백)
# ==============================================================================
class _PureDecisionTree:
    """scikit-learn 미설치 환경에서도 동작하는 경량 회귀 의사결정나무"""

    def __init__(self, max_depth: int = 5, min_samples_leaf: int = 2, max_features: Optional[int] = None, rng: Optional[random.Random] = None):
        self.max_depth = max_depth
        self.min_samples_leaf = min_samples_leaf
        self.max_features = max_features
        self.rng = rng or random.Random(42)
        self.tree: Any = 0.0

    def fit(self, X: List[List[float]], y: List[float]) -> None:
        indices = list(range(len(X)))
        self.tree = self._build(X, y, indices, depth=0)

    def _build(self, X: List[List[float]], y: List[float], idxs: List[int], depth: int) -> Any:
        n = len(idxs)
        mean_y = sum(y[i] for i in idxs) / n
        if depth >= self.max_depth or n < 2 * self.min_samples_leaf:
            return mean_y
        first_y = y[idxs[0]]
        if all(abs(y[i] - first_y) < 1e-12 for i in idxs):
            return mean_y

        n_feats = len(X[0])
        feat_pool = list(range(n_feats))
        if self.max_features and self.max_features < n_feats:
            feat_pool = self.rng.sample(feat_pool, self.max_features)

        total_sum = sum(y[i] for i in idxs)
        total_sq = sum(y[i] * y[i] for i in idxs)
        best_sse = total_sq - (total_sum * total_sum) / n
        best_feat = -1
        best_thr = 0.0
        best_left: List[int] = []
        best_right: List[int] = []

        for f in feat_pool:
            sorted_idx = sorted(idxs, key=lambda i: X[i][f])
            left_sum = 0.0
            left_sq = 0.0
            for pos in range(n - 1):
                idx_cur = sorted_idx[pos]
                yv = y[idx_cur]
                left_sum += yv
                left_sq += yv * yv
                left_cnt = pos + 1
                right_cnt = n - left_cnt
                if left_cnt < self.min_samples_leaf or right_cnt < self.min_samples_leaf:
                    continue
                v_cur = X[idx_cur][f]
                v_nxt = X[sorted_idx[pos + 1]][f]
                if v_cur == v_nxt:
                    continue
                right_sum = total_sum - left_sum
                right_sq = total_sq - left_sq
                sse = (left_sq - (left_sum * left_sum) / left_cnt) + (right_sq - (right_sum * right_sum) / right_cnt)
                if sse < best_sse - 1e-12:
                    best_sse = sse
                    best_feat = f
                    best_thr = 0.5 * (v_cur + v_nxt)
                    best_left = sorted_idx[:left_cnt]
                    best_right = sorted_idx[left_cnt:]

        if best_feat < 0:
            return mean_y
        return (
            best_feat,
            best_thr,
            self._build(X, y, best_left, depth + 1),
            self._build(X, y, best_right, depth + 1),
        )

    def predict_one(self, row: List[float]) -> float:
        node = self.tree
        while isinstance(node, tuple):
            feat, thr, left, right = node
            node = left if row[feat] <= thr else right
        return float(node)


class PureRandomForestRegressor:
    """scikit-learn 미설치 환경용 Bagging 회귀 포레스트 폴백"""

    def __init__(self, n_estimators: int = 80, max_depth: int = 5, min_samples_leaf: int = 2, random_state: int = 42):
        self.n_estimators = n_estimators
        self.max_depth = max_depth
        self.min_samples_leaf = min_samples_leaf
        self.random_state = random_state
        self.trees: List[_PureDecisionTree] = []

    def fit(self, X: List[List[float]], y: List[float]) -> None:
        rng = random.Random(self.random_state)
        n = len(X)
        m_feat = max(1, int(math.sqrt(len(X[0])) * 1.5)) if X else 1
        self.trees = []
        for _ in range(self.n_estimators):
            sample_idx = [rng.randrange(n) for _ in range(n)]
            Xs = [X[i] for i in sample_idx]
            ys = [y[i] for i in sample_idx]
            t = _PureDecisionTree(
                max_depth=self.max_depth,
                min_samples_leaf=self.min_samples_leaf,
                max_features=m_feat,
                rng=rng,
            )
            t.fit(Xs, ys)
            self.trees.append(t)

    def predict(self, X: List[List[float]]) -> List[float]:
        return [sum(t.predict_one(row) for t in self.trees) / len(self.trees) for row in X]


def fit_and_predict_ml(
    rows_ml: List[Dict[str, Any]],
    base_week: Optional[int],
    min_weeks: int,
    extra_train_rows: Optional[List[Dict[str, Any]]] = None,
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """
    실측 주차(+ 다중 폴더의 타 구역/타 시점 학습 데이터)로 RandomForestRegressor 학습 후
    전체 작기(과거 적합 + 미래 순차 롤아웃) 예측 및 오차 지표 산출
    """
    obs_target = [r for r in rows_ml if r["y_actual"] is not None]
    extra_obs = [r for r in (extra_train_rows or []) if r.get("y_actual") is not None]
    if not obs_target and not extra_obs:
        raise ValueError("학습에 사용할 실측 수확량(y_actual / obsKg / 실측수확량) 데이터가 없습니다.")

    if base_week is not None:
        resolved_base_week = base_week
    elif obs_target:
        resolved_base_week = max(r["week"] for r in obs_target)
    else:
        resolved_base_week = max(r["week"] for r in extra_obs)

    train_target = [r for r in obs_target if r["week"] <= resolved_base_week]
    train_rows = train_target + extra_obs
    if len(train_rows) < min_weeks:
        raise ValueError(
            f"학습 데이터가 부족합니다 ({len(train_rows)}건, 기준 주차 {resolved_base_week}주 이하). "
            f"최소 {min_weeks}개 이상의 실측 데이터가 필요합니다 (--min-weeks 옵션으로 조정 가능)."
        )

    def extract_matrix(items: List[Dict[str, Any]]) -> List[List[float]]:
        return [[float(it.get(f, 0.0) or 0.0) for f in FEATURES] for it in items]

    X_train_list = extract_matrix(train_rows)
    y_train_list = [float(r["y_actual"]) for r in train_rows]

    use_sklearn = SklearnRandomForestRegressor is not None and pd is not None and np is not None
    if use_sklearn:
        X_train_df = pd.DataFrame(X_train_list, columns=FEATURES).fillna(0.0)
        y_train_ser = pd.Series(y_train_list, dtype=float)
        model = SklearnRandomForestRegressor(
            n_estimators=200,
            max_depth=6,
            min_samples_leaf=2,
            random_state=42,
        )
        model.fit(X_train_df, y_train_ser)
        model_name = "RandomForestRegressor (HTML-Physics Sync)"

        def predict_row(feat_dict: Dict[str, float]) -> float:
            x_df = pd.DataFrame([[feat_dict[f] for f in FEATURES]], columns=FEATURES).fillna(0.0)
            return float(model.predict(x_df)[0])
    else:
        model = PureRandomForestRegressor(
            n_estimators=80,
            max_depth=5,
            min_samples_leaf=2,
            random_state=42,
        )
        model.fit(X_train_list, y_train_list)
        model_name = "RandomForestRegressor-PurePy (HTML-Physics Sync)"

        def predict_row(feat_dict: Dict[str, float]) -> float:
            return float(model.predict([[feat_dict[f] for f in FEATURES]])[0])

    # 전체 주차 예측 생성 (기준 주차 이전: 직접 피처 예측 / 이후: 물리 시뮬레이터 피처 + 이동평균 순차 롤아웃)
    last_train = train_rows[-1]
    history_y = [float(r["y_actual"]) for r in (train_target if train_target else train_rows)]
    predictions: List[Dict[str, Any]] = []

    abs_err_sum = 0.0
    sq_err_sum = 0.0
    bias_sum = 0.0
    obs_eval_cnt = 0

    for r in rows_ml:
        wk = int(r["week"])
        if not r.get("has_harvest_potential", True) and r["y_actual"] is None:
            y_hat = 0.0
        elif wk <= resolved_base_week or not obs_target:
            feat_map = {f: float(r.get(f, 0.0) or 0.0) for f in FEATURES}
            y_hat = max(0.0, predict_row(feat_map))
        else:
            steps = wk - resolved_base_week
            decay = math.exp(-0.25 * steps)
            ma3 = sum(history_y[-3:]) / len(history_y[-3:]) if history_y else float(r["adj_kg_sim"])
            feat_map = {f: float(r.get(f, 0.0) or 0.0) for f in FEATURES}
            feat_map["ma_3week"] = ma3
            # 단기 외삽 시 최근 관측 말단 기상과 중장기 시뮬레이터 기상을 부드럽게 블렌딩
            for blend_col in ("dli_daily", "t_mean", "vpd_mean"):
                feat_map[blend_col] = float(r[blend_col]) * (1.0 - 0.35 * decay) + float(last_train[blend_col]) * (0.35 * decay)
            y_hat = max(0.0, predict_row(feat_map))
            history_y.append(y_hat)

        y_hat_round = round(y_hat, 3)
        y_act = r["y_actual"]
        if y_act is not None:
            err = y_hat_round - float(y_act)
            abs_err_sum += abs(err)
            sq_err_sum += err * err
            bias_sum += err
            obs_eval_cnt += 1

        predictions.append({
            "week": wk,
            "week_no": wk,
            "start": r["start"],
            "date": r["date"],
            "y_pred": y_hat_round,
            "pred": y_hat_round,
            "yield": y_hat_round,
            "y_actual": round(float(y_act), 3) if y_act is not None else None,
            "orig_kg": round(float(r["orig_kg_sim"]), 3),
            "adj_kg": round(float(r["adj_kg_sim"]), 3),
            "k_total": round(float(r["k_total"]), 4),
            "dyn_lag": round(float(r["dyn_lag"]), 1),
        })

    # 대상 타임라인에 실측이 없는 미래 계획 예측인 경우 학습셋 오차 지표로 폴백
    if obs_eval_cnt > 0:
        mae = round(abs_err_sum / obs_eval_cnt, 4)
        rmse = round(math.sqrt(sq_err_sum / obs_eval_cnt), 4)
        bias = round(bias_sum / obs_eval_cnt, 4)
    else:
        tr_abs, tr_sq, tr_bias = 0.0, 0.0, 0.0
        for tr_r in train_rows:
            tr_pred = max(0.0, predict_row({f: float(tr_r.get(f, 0.0) or 0.0) for f in FEATURES}))
            tr_err = round(tr_pred, 3) - float(tr_r["y_actual"])
            tr_abs += abs(tr_err)
            tr_sq += tr_err * tr_err
            tr_bias += tr_err
        mae = round(tr_abs / len(train_rows), 4)
        rmse = round(math.sqrt(tr_sq / len(train_rows)), 4)
        bias = round(tr_bias / len(train_rows), 4)

    metrics = {
        "model_name": model_name,
        "base_week": resolved_base_week,
        "matched_obs_weeks": obs_eval_cnt if obs_eval_cnt > 0 else len(train_rows),
        "train_samples": len(train_rows),
        "mae": mae,
        "rmse": rmse,
        "bias": bias,
    }
    return predictions, metrics


# ==============================================================================
# 8. 통합 실행 파이프라인 & CLI
# ==============================================================================
def resolve_default_output_path() -> str:
    """Vite 정적 서빙 구조(public/simulator.html) 감지 시 public/ml_pred.json 기본 반환"""
    if Path("public").is_dir() and Path("public/simulator.html").exists():
        return "public/ml_pred.json"
    if Path("public").is_dir():
        return "public/ml_pred.json"
    return "ml_pred.json"


def export_ml_predictions_json(payload: Dict[str, Any], output_json: Optional[str] = None) -> Path:
    """
    public/simulator.html의 parseMlPredictionsJson() / mlPredState 규격에 맞춰
    ML 예측 결과를 JSON 파일로 저장한다.
    """
    target_path = Path(output_json or resolve_default_output_path())
    target_path.parent.mkdir(parents=True, exist_ok=True)
    target_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return target_path


def auto_discover_input() -> Tuple[Optional[str], Optional[Path]]:
    search_dirs = [Path("."), Path("public")]
    for base in search_dirs:
        if not base.is_dir():
            continue
        for cand in (
            "LMH-Tomato-Calibrated-Preset.json",
            "LMH-Tomato-Physiology-Plan-v3.json",
            "plan.json",
        ):
            p = base / cand
            if p.exists():
                return "plan", p

    for base in search_dirs:
        mp_csvs = sorted(glob.glob(str(base / "토마토_마스터플랜_v3_*.csv")))
        if mp_csvs:
            return "masterplan_csv", Path(mp_csvs[0])

    if Path("smartfarm_yield_swing.db").exists() or Path("raw_yield_source.csv").exists():
        return "legacy", None

    for base in search_dirs:
        p = base / "토마토_실측데이터_입력양식.csv"
        if p.exists():
            return "calib_csv", p

    return None, None


def _select_target_zone_from_records(records: List[Dict[str, Any]], requested_zone: Optional[str]) -> str:
    zones = sorted({str(r.get("zone_id") or "Zone_A") for r in records})
    if not zones:
        return requested_zone or "Zone_A"
    if requested_zone and requested_zone in zones:
        return requested_zone
    if len(zones) == 1:
        return zones[0]

    def zone_rank(z: str) -> Tuple[str, int, int, str]:
        zr = [r for r in records if str(r.get("zone_id") or "") == z]
        max_dt = max((str(r.get("date") or "") for r in zr), default="")
        obs_cnt = sum(1 for r in zr if r.get("y_actual") is not None)
        return (max_dt, obs_cnt, len(zr), z)

    return max(zones, key=zone_rank)


def run_pipeline(
    plan_path: Optional[str] = None,
    masterplan_csv_path: Optional[str] = None,
    calib_csv_path: Optional[str] = None,
    calib_dir: Optional[str] = DEFAULT_CALIB_DIR,
    no_calib_dir: bool = False,
    calib_dir_recursive: bool = False,
    weather_csv_path: Optional[str] = None,
    weather_db_path: Optional[str] = None,
    station: Optional[str] = None,
    db_path: str = "smartfarm_yield_swing.db",
    csv_path: str = "raw_yield_source.csv",
    zone_id: Optional[str] = None,
    base_week: Optional[int] = None,
    min_weeks: int = 4,
    output_json: Optional[str] = None,
    season_start_date: str = "2027-01-01",
    harvest_end_date: str = "2027-12-31",
) -> Dict[str, Any]:
    plan_data: Optional[Dict[str, Any]] = None
    extra_train_rows: List[Dict[str, Any]] = []

    # 1) 다중 실측 CSV/Excel 폴더 스캔 (--no-calib-dir 미지정 시)
    dir_scan: Dict[str, Any] = {
        "exists": False,
        "calib_dir": calib_dir or "",
        "files_found": [],
        "files_used": [],
        "csv_count": 0,
        "excel_count": 0,
        "records": [],
        "raw_rows": 0,
        "merged_rows": 0,
        "duplicates_merged": 0,
        "valid_obs_count": 0,
        "zones": [],
    }
    if not no_calib_dir and calib_dir:
        dir_scan = load_calib_dir_records(
            Path(calib_dir),
            recursive=calib_dir_recursive,
            default_zone=zone_id or "Zone_A",
        )
        if dir_scan["files_used"]:
            print(
                f"[실측 폴더 스캔] '{dir_scan['calib_dir']}' | 사용 파일 {len(dir_scan['files_used'])}개 "
                f"(CSV {dir_scan['csv_count']}개, Excel {dir_scan['excel_count']}개: {', '.join(dir_scan['files_used'])}) | "
                f"원본 {dir_scan['raw_rows']}행 -> 병합 {dir_scan['merged_rows']}행 "
                f"(중복 정리 {dir_scan['duplicates_merged']}건, 유효 실측 {dir_scan['valid_obs_count']}건, 구역: {', '.join(dir_scan['zones'])})"
            )
        elif dir_scan["exists"] and not dir_scan["files_found"]:
            print(f"[안내] 지정된 실측 폴더 '{dir_scan['calib_dir']}'가 비어 있어 기존 입력/폴백 경로를 확인합니다.")

    # 2) 명시 입력이 없고 다중 폴더에도 데이터가 없으면 기존 자동 탐색 수행
    has_dir_records = bool(dir_scan["records"])
    if not any([plan_path, masterplan_csv_path, calib_csv_path]) and not has_dir_records and not (
        Path(db_path).exists() or Path(csv_path).exists()
    ):
        kind, found = auto_discover_input()
        if kind == "plan" and found:
            plan_path = str(found)
        elif kind == "masterplan_csv" and found:
            masterplan_csv_path = str(found)
        elif kind == "calib_csv" and found:
            calib_csv_path = str(found)

    # plan / masterplan 이 있는 경우 자동 탐색 보조 로드 (폴더 파일과 병합)
    if not any([plan_path, masterplan_csv_path]) and has_dir_records:
        kind, found = auto_discover_input()
        if kind == "plan" and found:
            plan_path = str(found)
        elif kind == "masterplan_csv" and found:
            masterplan_csv_path = str(found)

    if plan_path:
        plan_data = load_html_plan_json(
            Path(plan_path),
            weather_db_path=Path(weather_db_path) if weather_db_path else None,
        )
    elif masterplan_csv_path:
        plan_data = load_masterplan_csv(Path(masterplan_csv_path))

    weather_ds: Optional[Dict[str, Any]] = None
    if weather_csv_path:
        stations = load_weather_csv(Path(weather_csv_path))
        weather_ds = stations.get(station) if (station and station in stations) else next(iter(stations.values()))

    # 3) 파이프라인 분기
    if plan_data is not None:
        config = plan_data["config"]
        week_rows = plan_data["week_rows"]
        env_months = plan_data["env_months"]
        calib_rows = plan_data["calib_rows"]
        resolved_zone = zone_id or plan_data["zone_label"]
        source_kind = plan_data["source"]

        if weather_ds is not None:
            week_rows = apply_weather_dataset_to_plan(weather_ds, config, week_rows)
            resolved_zone = zone_id or weather_ds["name"]
            source_kind += "+weather_csv"

        if has_dir_records:
            target_dir_recs = [
                r for r in dir_scan["records"]
                if r.get("zone_id") in (resolved_zone, zone_id, "Zone_A", plan_data["zone_label"])
            ]
            other_dir_recs = [r for r in dir_scan["records"] if r not in target_dir_recs]
            calib_rows, unmatched_target = apply_calib_records_to_periods(target_dir_recs, config, calib_rows)
            extra_raw = other_dir_recs + unmatched_target
            if extra_raw:
                extra_train_rows = build_tabular_ml_rows(extra_raw, default_start_date=str(config["start"]), weather_ds=weather_ds)
            source_kind += f"+calib_dir({len(dir_scan['files_used'])}files)"

        if calib_csv_path:
            c_path = Path(calib_csv_path)
            if c_path.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
                c_recs, _ = parse_flexible_calib_file(c_path, source_label=c_path.name, default_zone=resolved_zone)
                calib_rows, c_unmatched = apply_calib_records_to_periods(c_recs, config, calib_rows)
                if c_unmatched:
                    extra_train_rows.extend(
                        build_tabular_ml_rows(c_unmatched, default_start_date=str(config["start"]), weather_ds=weather_ds)
                    )
            else:
                calib_rows = parse_calib_csv_into_rows(c_path, config, calib_rows)
            source_kind += "+calib_file"

        sim = simulate_masterplan(config, week_rows, env_months, calib_rows)
        rows_ml = build_ml_rows_from_simulation(sim, calib_rows)

    elif has_dir_records or calib_csv_path:
        combined_recs: List[Dict[str, Any]] = list(dir_scan["records"])
        used_files_cnt = len(dir_scan["files_used"])
        if calib_csv_path:
            explicit_recs, _ = parse_flexible_calib_file(
                Path(calib_csv_path),
                source_label=Path(calib_csv_path).name,
                default_zone=zone_id or "Zone_A",
            )
            combined_recs.extend(explicit_recs)
            combined_recs, _ = merge_and_deduplicate_calib_records(combined_recs)
            used_files_cnt += 1

        if not combined_recs:
            raise ValueError(
                f"'{calib_dir}' 폴더와 지정 파일에서 유효한 실측 수확량 데이터를 찾지 못했습니다."
            )

        resolved_zone = _select_target_zone_from_records(combined_recs, zone_id)
        target_recs = [r for r in combined_recs if str(r.get("zone_id") or "") == resolved_zone]
        other_recs = [r for r in combined_recs if str(r.get("zone_id") or "") != resolved_zone]

        # 풍부한 환경 피처가 있거나 1~5주차부터 즉시 수확 실측이 시작되는 표 형식 데이터 여부 판별
        obs_wks = [int(r["week_no"]) for r in target_recs if r.get("y_actual") is not None and r.get("week_no") is not None]
        is_tabular_style = any(r.get("_has_rich_features") for r in target_recs) or (bool(obs_wks) and min(obs_wks) <= 5)

        if is_tabular_style:
            rows_ml = build_tabular_ml_rows(target_recs, default_start_date=season_start_date, weather_ds=weather_ds)
            if other_recs:
                extra_train_rows = build_tabular_ml_rows(other_recs, default_start_date=season_start_date, weather_ds=weather_ds)
            source_kind = f"calib_dir_tabular({used_files_cnt}files)" if has_dir_records else "calib_csv_tabular"
        else:
            # HTML [CSV 양식 받기] 형식: 날짜가 기본 작기와 다르면 관측 기간에 맞춰 시뮬레이션 작기 자동 정렬
            dated = sorted(r["date"] for r in target_recs if r.get("date"))
            sim_start = season_start_date
            sim_end = harvest_end_date
            if dated and season_start_date == "2027-01-01" and harvest_end_date == "2027-12-31":
                min_dt = parse_date(dated[0])
                max_dt = parse_date(dated[-1])
                if min_dt < parse_date("2027-01-01") or max_dt > parse_date("2027-12-31"):
                    sim_start = min_dt.isoformat()
                    span_end = max(max_dt, min_dt + timedelta(days=364))
                    if (span_end - min_dt).days > 730:
                        span_end = min_dt + timedelta(days=730)
                    sim_end = span_end.isoformat()

            config = {**DEFAULT_CONFIG, "start": sim_start, "harvestEnd": sim_end}
            periods = plan_periods(config["start"], config["harvestEnd"])
            env_months = [dict(x) for x in STANDARD_ENV_MONTHS]
            week_rows = [
                {
                    "mj": DEFAULT_MONTHLY_MJ[p["at"].month - 1],
                    "lag": int(config["cropLag"]),
                    "eff": float(config["eff"]),
                }
                for p in periods
            ]
            if weather_ds is not None:
                week_rows = apply_weather_dataset_to_plan(weather_ds, config, week_rows)
                if not zone_id and resolved_zone == "Zone_A":
                    resolved_zone = weather_ds["name"]

            calib_rows, unmatched_target = apply_calib_records_to_periods(target_recs, config, None)
            extra_raw = other_recs + unmatched_target
            if extra_raw:
                extra_train_rows = build_tabular_ml_rows(extra_raw, default_start_date=sim_start, weather_ds=weather_ds)

            sim = simulate_masterplan(config, week_rows, env_months, calib_rows)
            rows_ml = build_ml_rows_from_simulation(sim, calib_rows)
            source_kind = f"calib_dir_sim({used_files_cnt}files)" if has_dir_records else "html_calib_csv"
            if weather_ds is not None:
                source_kind += "+weather_csv"

        if calib_csv_path and has_dir_records:
            source_kind += "+calib_csv"

    else:
        resolved_zone = zone_id or "Zone_A"
        try:
            rows_ml, source_kind = load_legacy_tabular_records(
                db_path=db_path,
                csv_path=csv_path,
                zone_id=resolved_zone,
                season_start_date=season_start_date,
            )
        except FileNotFoundError as exc:
            calib_dir_note = f"'{calib_dir}' 폴더 내 *.csv/*.xlsx/*.xls 파일, " if not no_calib_dir and calib_dir else ""
            raise FileNotFoundError(
                f"학습할 실측 데이터를 찾을 수 없습니다. {calib_dir_note}"
                f"HTML 내보내기 파일(--plan, --masterplan-csv, --calib-csv) 또는 '{db_path}' / '{csv_path}'를 확인해 주세요."
            ) from exc

    predictions, metrics = fit_and_predict_ml(
        rows_ml,
        base_week=base_week,
        min_weeks=min_weeks,
        extra_train_rows=extra_train_rows,
    )

    payload = {
        "zone": resolved_zone,
        "model": metrics["model_name"],
        "base_week": metrics["base_week"],
        "matched_obs_weeks": metrics["matched_obs_weeks"],
        "train_samples": metrics["train_samples"],
        "calib_files_used": dir_scan["files_used"],
        "mae": metrics["mae"],
        "rmse": metrics["rmse"],
        "bias": metrics["bias"],
        "source": source_kind,
        "features": FEATURES,
        "predictions": predictions,
    }

    export_ml_predictions_json(payload, output_json=output_json)
    return payload


def train_and_export_json(
    zone_id: str = "Zone_A",
    base_week: Optional[int] = None,
    min_weeks: int = 4,
    output_json: Optional[str] = None,
    season_start_date: str = "2027-01-01",
    **kwargs: Any,
) -> Dict[str, Any]:
    """기존 호출부 호환용 래퍼 함수"""
    return run_pipeline(
        zone_id=zone_id,
        base_week=base_week,
        min_weeks=min_weeks,
        output_json=output_json,
        season_start_date=season_start_date,
        **kwargs,
    )


def build_arg_parser() -> argparse.ArgumentParser:
    default_out = resolve_default_output_path()
    parser = argparse.ArgumentParser(
        description="LMH-Tomato-Planner v4.0 (public/simulator.html) 연동 ML 학습 및 ml_pred.json 생성기",
    )
    parser.add_argument("--plan", help="HTML [계획 저장(JSON)] / [캘리브레이션 프리셋 저장] / localStorage JSON 파일 경로")
    parser.add_argument("--masterplan-csv", help="HTML [마스터플랜 CSV 다운로드] (토마토_마스터플랜_v3_*.csv) 경로")
    parser.add_argument("--calib-csv", help="HTML [CSV 양식 받기] 실측 데이터 CSV/Excel(.xlsx/.xls) 경로")
    parser.add_argument(
        "--calib-dir",
        default=DEFAULT_CALIB_DIR,
        help=f"다중 실측 CSV/Excel(.csv, .xlsx, .xls) 파일 자동 수집·병합 폴더 경로 (기본: {DEFAULT_CALIB_DIR})",
    )
    parser.add_argument(
        "--no-calib-dir",
        action="store_true",
        help="실측 파일 폴더(--calib-dir) 자동 스캔 비활성화",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="--calib-dir 하위 디렉터리의 *.csv, *.xlsx, *.xls 파일까지 재귀 스캔",
    )
    parser.add_argument("--weather-csv", help="기상청 ASOS 일자료 또는 Open-Meteo 기상 관측 CSV 경로")
    parser.add_argument("--weather-db", help="localStorage(LMH_TOMATO_PLANNER_V3_WEATHER_DB) JSON 파일 경로")
    parser.add_argument("--station", help="다중 지점 ASOS CSV에서 사용할 지점명 (예: 논산, 대전, 수원)")
    parser.add_argument("--db", default="smartfarm_yield_swing.db", help="레거시 SQLite DB 경로")
    parser.add_argument("--csv", default="raw_yield_source.csv", help="레거시 원천 CSV 경로")
    parser.add_argument("--zone", default=None, help="출력 JSON에 기록할 구역/품종 라벨")
    parser.add_argument("--base-week", type=int, default=None, help="학습 기준 마지막 주차 (미지정 시 실측 마지막 주차 자동 사용)")
    parser.add_argument("--min-weeks", type=int, default=4, help="학습에 필요한 최소 실측 샘플 수 (기본: 4)")
    parser.add_argument("--start", default="2027-01-01", help="작기 시작일 YYYY-MM-DD (단독 CSV 실행 시)")
    parser.add_argument("--harvest-end", default="2027-12-31", help="수확 종료일 YYYY-MM-DD (단독 CSV 실행 시)")
    parser.add_argument(
        "-o",
        "--output",
        default=default_out,
        help=f"출력 ML 예측 JSON 경로 (기본: {default_out})",
    )
    return parser


def main() -> None:
    parser = build_arg_parser()
    args = parser.parse_args()

    result = run_pipeline(
        plan_path=args.plan,
        masterplan_csv_path=args.masterplan_csv,
        calib_csv_path=args.calib_csv,
        calib_dir=args.calib_dir,
        no_calib_dir=args.no_calib_dir,
        calib_dir_recursive=args.recursive,
        weather_csv_path=args.weather_csv,
        weather_db_path=args.weather_db,
        station=args.station,
        db_path=args.db,
        csv_path=args.csv,
        zone_id=args.zone,
        base_week=args.base_week,
        min_weeks=args.min_weeks,
        output_json=args.output,
        season_start_date=args.start,
        harvest_end_date=args.harvest_end,
    )

    out_saved = args.output or resolve_default_output_path()
    mae_str = f"{result['mae']:.4f}" if result["mae"] is not None else "-"
    rmse_str = f"{result['rmse']:.4f}" if result["rmse"] is not None else "-"
    files_cnt = len(result.get("calib_files_used") or [])
    train_samples = result.get("train_samples", result["matched_obs_weeks"])
    print(
        f"[완료] {out_saved} 저장됨 | 구역: {result['zone']} | 모델: {result['model']} | "
        f"폴더 파일: {files_cnt}개 | 학습 샘플: {train_samples}건 | "
        f"총 {len(result['predictions'])}주 예측 (실측 매칭 {result['matched_obs_weeks']}주, "
        f"MAE={mae_str} kg/㎡, RMSE={rmse_str} kg/㎡)"
    )


if __name__ == "__main__":
    main()
