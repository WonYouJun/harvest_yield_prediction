#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
토마토 스마트팜 마스터플랜 · 머신러닝 학습 통합 Streamlit 웹 앱 (app.py)

- 다중 CSV(.csv) 및 엑셀(.xlsx, .xls) 실측 파일 업로드·미리보기·개별 제외/초기화
- 버튼 클릭으로 train_export_ml.py 실행 및 학습 로그·MAE/RMSE·예측 차트·ml_pred.json 확인/다운로드
- 기존 public/simulator.html을 Streamlit 내부에 임베드하고 생성된 ml_pred.json을 자동 연동
"""

from __future__ import annotations

import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

ROOT_DIR = Path(__file__).resolve().parent
TRAIN_SCRIPT = ROOT_DIR / "train_export_ml.py"
SIMULATOR_HTML = ROOT_DIR / "public" / "simulator.html"
REPO_CALIB_DIR = ROOT_DIR / "data" / "calib_csv"
PUBLIC_ML_PRED = ROOT_DIR / "public" / "ml_pred.json"
SUPPORTED_EXTS = (".csv", ".xlsx", ".xls")


# ==============================================================================
# 유틸리티: 인코딩 감지 CSV/Excel 미리보기 · 작업 폴더 관리 · HTML ML 주입
# ==============================================================================
def decode_bytes_auto(raw: bytes) -> str:
    for enc in ("utf-8-sig", "utf-8", "cp949", "euc-kr"):
        try:
            return raw.decode(enc).lstrip("\ufeff")
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace").lstrip("\ufeff")


def inspect_uploaded_table(filename: str, raw: bytes) -> Dict[str, Any]:
    """업로드된 CSV/Excel 파일의 행 수, 컬럼 목록, 시트 정보, 상위 미리보기를 안전하게 추출"""
    ext = Path(filename).suffix.lower()
    info: Dict[str, Any] = {
        "file_type": "CSV" if ext == ".csv" else f"Excel ({ext})",
        "sheet_names": [],
        "active_sheet": None,
        "row_count": 0,
        "col_count": 0,
        "columns": [],
        "preview_df": None,
        "error": None,
    }
    try:
        if ext == ".csv":
            text = decode_bytes_auto(raw)
            lines = [ln for ln in text.splitlines() if ln.strip()]
            if not lines:
                info["error"] = "파일이 비어 있습니다."
                return info
            # 구분자 추정(탭/세미콜론/쉼표)
            sample_head = lines[0]
            sep = "\t" if "\t" in sample_head else (";" if ";" in sample_head and "," not in sample_head else ",")
            try:
                df = pd.read_csv(io.StringIO(text), sep=sep)
            except Exception:
                # 마스터플랜 CSV 등 다단 헤더 fallback
                df = pd.DataFrame([ln.split(sep) for ln in lines[:20]])
            info["row_count"] = int(len(df))
            info["col_count"] = int(len(df.columns))
            info["columns"] = [str(c) for c in df.columns]
            info["preview_df"] = df.head(5)
            return info

        # Excel (.xlsx / .xls)
        engine = "xlrd" if ext == ".xls" else "openpyxl"
        bio = io.BytesIO(raw)
        xls = pd.ExcelFile(bio, engine=engine)
        info["sheet_names"] = [str(s) for s in xls.sheet_names]
        chosen_df: Optional[pd.DataFrame] = None
        chosen_sheet: Optional[str] = None
        for s_name in xls.sheet_names:
            df_s = xls.parse(s_name)
            if not df_s.dropna(how="all").empty:
                chosen_df = df_s.dropna(how="all")
                chosen_sheet = str(s_name)
                break
        if chosen_df is None and xls.sheet_names:
            chosen_sheet = str(xls.sheet_names[0])
            chosen_df = xls.parse(chosen_sheet)

        if chosen_df is not None:
            info["active_sheet"] = chosen_sheet
            info["row_count"] = int(len(chosen_df))
            info["col_count"] = int(len(chosen_df.columns))
            info["columns"] = [str(c) for c in chosen_df.columns]
            info["preview_df"] = chosen_df.head(5)
        else:
            info["error"] = "유효한 시트 데이터를 찾지 못했습니다."
    except Exception as exc:
        info["error"] = f"미리보기 파싱 실패: {exc}"
    return info


def sanitize_filename(name: str) -> str:
    safe = Path(name).name.strip()
    safe = re.sub(r'[\\/:*?"<>|]+', "_", safe)
    return safe or "uploaded_data.csv"


def ensure_session_state() -> None:
    if "uploader_key" not in st.session_state:
        st.session_state["uploader_key"] = 0
    if "managed_files" not in st.session_state:
        st.session_state["managed_files"] = {}
    if "session_work_dir" not in st.session_state:
        st.session_state["session_work_dir"] = tempfile.mkdtemp(prefix="tomato_ml_")
    if "seeded_repo_files" not in st.session_state:
        st.session_state["seeded_repo_files"] = True
        if REPO_CALIB_DIR.is_dir():
            for p in sorted(REPO_CALIB_DIR.iterdir(), key=lambda x: x.name.lower()):
                if p.is_file() and not p.name.startswith("~$") and p.suffix.lower() in SUPPORTED_EXTS:
                    raw = p.read_bytes()
                    meta = inspect_uploaded_table(p.name, raw)
                    st.session_state["managed_files"][p.name] = {
                        "name": p.name,
                        "bytes": raw,
                        "size_kb": round(len(raw) / 1024.0, 1),
                        "origin": "저장소 기본 파일",
                        "added_at": datetime.now().strftime("%H:%M:%S"),
                        "meta": meta,
                    }
    if "ml_pred_payload" not in st.session_state:
        if PUBLIC_ML_PRED.exists():
            try:
                st.session_state["ml_pred_payload"] = json.loads(PUBLIC_ML_PRED.read_text(encoding="utf-8"))
            except Exception:
                st.session_state["ml_pred_payload"] = None
        else:
            st.session_state["ml_pred_payload"] = None
    if "last_run_log" not in st.session_state:
        st.session_state["last_run_log"] = ""
    if "last_run_error" not in st.session_state:
        st.session_state["last_run_error"] = ""
    if "last_run_cmd" not in st.session_state:
        st.session_state["last_run_cmd"] = ""


def sync_managed_files_to_workdir() -> Path:
    """세션에 관리 중인 파일들을 학습 전용 작업 폴더(calib_csv)로 깨끗하게 동기화"""
    work_calib_dir = Path(st.session_state["session_work_dir"]) / "calib_csv"
    if work_calib_dir.exists():
        shutil.rmtree(work_calib_dir, ignore_errors=True)
    work_calib_dir.mkdir(parents=True, exist_ok=True)

    for fname, item in st.session_state["managed_files"].items():
        target = work_calib_dir / sanitize_filename(fname)
        target.write_bytes(item["bytes"])
    return work_calib_dir


def run_training_subprocess(
    calib_dir: Path,
    out_json_path: Path,
    min_weeks: int,
    zone_opt: str,
    base_week_opt: int,
    plan_file: Optional[Any],
    weather_file: Optional[Any],
) -> Tuple[bool, str, str, Optional[Dict[str, Any]]]:
    """train_export_ml.py를 서브프로세스로 실행하고 로그 및 결과 JSON을 반환"""
    aux_dir = Path(st.session_state["session_work_dir"]) / "aux_inputs"
    aux_dir.mkdir(parents=True, exist_ok=True)

    cmd: List[str] = [
        sys.executable,
        str(TRAIN_SCRIPT),
        "--calib-dir",
        str(calib_dir),
        "--min-weeks",
        str(int(min_weeks)),
        "--output",
        str(out_json_path),
    ]
    if zone_opt.strip():
        cmd.extend(["--zone", zone_opt.strip()])
    if base_week_opt > 0:
        cmd.extend(["--base-week", str(int(base_week_opt))])

    if plan_file is not None:
        plan_name = sanitize_filename(plan_file.name)
        plan_path = aux_dir / plan_name
        plan_path.write_bytes(plan_file.getvalue())
        if plan_name.lower().endswith(".csv"):
            cmd.extend(["--masterplan-csv", str(plan_path)])
        else:
            cmd.extend(["--plan", str(plan_path)])

    if weather_file is not None:
        w_path = aux_dir / sanitize_filename(weather_file.name)
        w_path.write_bytes(weather_file.getvalue())
        cmd.extend(["--weather-csv", str(w_path)])

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"

    proc = subprocess.run(
        cmd,
        cwd=str(ROOT_DIR),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
    )
    combined_log = "\n".join(part for part in (proc.stdout.strip(), proc.stderr.strip()) if part)
    cmd_str = " ".join(cmd)

    if proc.returncode != 0:
        return False, cmd_str, combined_log or "알 수 없는 실행 오류가 발생했습니다.", None

    if not out_json_path.exists():
        return False, cmd_str, combined_log + "\n결과 파일(ml_pred.json)이 생성되지 않았습니다.", None

    payload = json.loads(out_json_path.read_text(encoding="utf-8"))
    # 로컬/단일 실행 환경 호환을 위해 public/ml_pred.json에도 동기화 시도
    try:
        PUBLIC_ML_PRED.parent.mkdir(parents=True, exist_ok=True)
        PUBLIC_ML_PRED.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass
    return True, cmd_str, combined_log, payload


def build_simulator_embed_html(ml_payload: Optional[Dict[str, Any]]) -> str:
    """public/simulator.html을 읽고 생성된 ml_pred.json을 window.__PRELOADED_ML_PRED__로 안전 주입"""
    if not SIMULATOR_HTML.exists():
        return "<h3>public/simulator.html 파일을 찾을 수 없습니다.</h3>"
    html_text = SIMULATOR_HTML.read_text(encoding="utf-8")
    if ml_payload is None:
        return html_text

    safe_json = json.dumps(ml_payload, ensure_ascii=False).replace("</", "<\\/")
    inject_script = (
        f"<script>\n"
        f"window.__PRELOADED_ML_PRED__ = {safe_json};\n"
        f'window.__PRELOADED_ML_PRED_NAME__ = "ml_pred.json (Streamlit 학습 결과)";\n'
        f"</script>\n"
    )
    if "</head>" in html_text:
        return html_text.replace("</head>", inject_script + "</head>", 1)
    return inject_script + html_text


def apply_sage_theme_css() -> None:
    """public/simulator.html의 연회색·세이지(#eef2ef ~ #f3f6f2) 톤과 Streamlit UI를 자연스럽게 통일"""
    st.markdown(
        """
        <style>
        :root {
            --sage-bg: #eef2ef;
            --sage-bg-soft: #f1f4f1;
            --sage-surface: #f6f8f6;
            --sage-surface-alt: #e6ede8;
            --sage-border: #dce3de;
            --sage-border-strong: #c6d3ca;
            --sage-ink: #1f2924;
            --sage-muted: #55665d;
            --sage-forest: #143f33;
            --sage-emerald: #1c674e;
            --sage-emerald-hover: #14503c;
        }

        /* 1. 전체 메인 배경 및 상단 헤더/데코레이션 */
        html, body, .stApp,
        [data-testid="stApp"],
        [data-testid="stAppViewContainer"],
        [data-testid="stAppViewBlockContainer"],
        [data-testid="stMain"],
        .main {
            background-color: var(--sage-bg) !important;
            color: var(--sage-ink);
        }

        header[data-testid="stHeader"] {
            background-color: rgba(238, 242, 239, 0.92) !important;
            backdrop-filter: blur(6px);
            border-bottom: 1px solid var(--sage-border);
        }

        [data-testid="stDecoration"] {
            background-image: linear-gradient(90deg, #143f33, #1c674e, #2c8055) !important;
            height: 3px !important;
        }

        .main .block-container {
            padding-top: 2.1rem;
            padding-bottom: 1.75rem;
        }

        /* 2. 타이포그래피 및 구분선 */
        .stApp h1, .stApp h2, .stApp h3, .stApp h4 {
            color: var(--sage-forest);
            letter-spacing: -0.015em;
        }
        .stApp p, .stApp li, .stApp label,
        [data-testid="stMarkdownContainer"] {
            color: var(--sage-ink);
        }
        [data-testid="stCaptionContainer"], .stApp small {
            color: var(--sage-muted) !important;
        }
        .stApp hr {
            border-color: var(--sage-border) !important;
        }

        /* 3. 사이드바 영역 */
        section[data-testid="stSidebar"],
        section[data-testid="stSidebar"] > div:first-child,
        [data-testid="stSidebarContent"] {
            background-color: var(--sage-surface) !important;
            border-right: 1px solid var(--sage-border);
            color: var(--sage-ink);
        }

        /* 4. 입력 위젯 (number_input, text_input 등) */
        [data-baseweb="input"],
        [data-baseweb="base-input"],
        [data-testid="stNumberInputContainer"] {
            background-color: #f8faf8 !important;
            border: 1px solid var(--sage-border-strong) !important;
            border-radius: 8px !important;
            color: var(--sage-ink) !important;
        }
        [data-baseweb="input"] input,
        [data-baseweb="base-input"] input {
            background-color: transparent !important;
            color: var(--sage-ink) !important;
        }
        [data-testid="stNumberInputStepDown"],
        [data-testid="stNumberInputStepUp"] {
            background-color: var(--sage-bg) !important;
            color: var(--sage-ink) !important;
            border-left: 1px solid var(--sage-border) !important;
        }
        [data-testid="stNumberInputStepDown"]:hover,
        [data-testid="stNumberInputStepUp"]:hover {
            background-color: var(--sage-surface-alt) !important;
        }

        /* 5. 탭(st.tabs) */
        [data-baseweb="tab-list"] {
            gap: 8px;
            border-bottom: 1px solid var(--sage-border-strong);
            padding-bottom: 0;
        }
        [data-baseweb="tab"] {
            background-color: var(--sage-surface-alt) !important;
            color: #36473e !important;
            border: 1px solid var(--sage-border) !important;
            border-bottom: none !important;
            border-radius: 10px 10px 0 0 !important;
            padding: 8px 16px !important;
            font-weight: 600 !important;
            transition: background-color 0.15s ease, color 0.15s ease;
        }
        [data-baseweb="tab"]:hover {
            background-color: #dde6df !important;
            color: var(--sage-forest) !important;
        }
        [data-baseweb="tab"][aria-selected="true"] {
            background-color: var(--sage-surface) !important;
            color: var(--sage-forest) !important;
            border-color: var(--sage-border-strong) !important;
            font-weight: 700 !important;
        }
        [data-baseweb="tab-highlight"] {
            background-color: var(--sage-emerald) !important;
            height: 3px !important;
        }
        [data-baseweb="tab-border"] {
            background-color: var(--sage-border) !important;
        }

        /* 6. Expander (안내 박스, 상세 미리보기, 로그, 테이블) */
        [data-testid="stExpander"],
        [data-testid="stExpander"] details {
            background-color: var(--sage-surface) !important;
            border: 1px solid var(--sage-border) !important;
            border-radius: 12px !important;
            box-shadow: 0 1px 2px rgba(20, 63, 51, 0.03);
        }
        [data-testid="stExpander"] summary {
            background-color: var(--sage-surface) !important;
            color: var(--sage-ink) !important;
            border-radius: 12px !important;
            font-weight: 600;
        }
        [data-testid="stExpander"] summary:hover {
            background-color: var(--sage-bg-soft) !important;
            color: var(--sage-forest) !important;
        }

        /* 7. 파일 업로더 (st.file_uploader) */
        [data-testid="stFileUploaderDropzone"] {
            background-color: var(--sage-surface) !important;
            border: 1.5px dashed #b4c4b9 !important;
            border-radius: 12px !important;
            color: var(--sage-ink) !important;
            transition: background-color 0.15s ease, border-color 0.15s ease;
        }
        [data-testid="stFileUploaderDropzone"]:hover {
            background-color: #eaf1ec !important;
            border-color: var(--sage-emerald) !important;
        }
        [data-testid="stFileUploaderDropzone"] small,
        [data-testid="stFileUploaderDropzone"] span {
            color: var(--sage-muted) !important;
        }
        [data-testid="stFileUploaderDropzone"] button {
            background-color: var(--sage-bg) !important;
            color: var(--sage-ink) !important;
            border: 1px solid var(--sage-border-strong) !important;
            border-radius: 8px !important;
            font-weight: 600 !important;
        }
        [data-testid="stFileUploaderDropzone"] button:hover {
            background-color: var(--sage-surface-alt) !important;
            border-color: var(--sage-emerald) !important;
            color: var(--sage-forest) !important;
        }
        [data-testid="stFileUploaderFile"] {
            background-color: var(--sage-surface) !important;
            border: 1px solid var(--sage-border) !important;
            border-radius: 8px !important;
            padding: 4px 10px !important;
        }

        /* 8. 버튼 (Primary / Secondary / Download) */
        button[kind="primary"],
        [data-testid="baseButton-primary"] {
            background-color: var(--sage-emerald) !important;
            border: 1px solid var(--sage-forest) !important;
            color: #ffffff !important;
            border-radius: 9px !important;
            font-weight: 700 !important;
        }
        button[kind="primary"] *,
        [data-testid="baseButton-primary"] * {
            color: #ffffff !important;
        }
        button[kind="primary"]:hover:not(:disabled),
        [data-testid="baseButton-primary"]:hover:not(:disabled) {
            background-color: var(--sage-emerald-hover) !important;
            border-color: #0e2e25 !important;
        }

        button[kind="secondary"],
        [data-testid="baseButton-secondary"],
        [data-testid="stDownloadButton"] button {
            background-color: var(--sage-surface) !important;
            border: 1px solid var(--sage-border-strong) !important;
            color: var(--sage-ink) !important;
            border-radius: 9px !important;
            font-weight: 600 !important;
        }
        button[kind="secondary"]:hover:not(:disabled),
        [data-testid="baseButton-secondary"]:hover:not(:disabled),
        [data-testid="stDownloadButton"] button:hover:not(:disabled) {
            background-color: var(--sage-surface-alt) !important;
            border-color: var(--sage-emerald) !important;
            color: var(--sage-forest) !important;
        }

        /* 9. 지표 카드(st.metric), 알림(st.alert), 데이터프레임, 코드 블록 */
        [data-testid="stMetric"] {
            background-color: var(--sage-surface) !important;
            border: 1px solid var(--sage-border) !important;
            border-radius: 10px !important;
            padding: 12px 14px !important;
            box-shadow: 0 1px 2px rgba(20, 63, 51, 0.03);
        }
        [data-testid="stMetricLabel"] {
            color: var(--sage-muted) !important;
        }
        [data-testid="stMetricValue"] {
            color: var(--sage-forest) !important;
        }

        [data-testid="stAlert"] {
            border-radius: 10px !important;
            border: 1px solid var(--sage-border) !important;
        }

        [data-testid="stDataFrame"],
        [data-testid="stCode"],
        [data-testid="stJson"] {
            background-color: var(--sage-surface) !important;
            border: 1px solid var(--sage-border) !important;
            border-radius: 10px !important;
            overflow: hidden;
        }

        /* 10. 시뮬레이터(simulator.html) 임베드 프레임 조화 */
        [data-testid="stIFrame"],
        [data-testid="stCustomComponentV1"] iframe,
        iframe[title*="streamlit_components"],
        iframe[srcdoc] {
            background-color: #f3f6f2 !important;
            border: 1px solid #cfd9d2 !important;
            border-radius: 14px !important;
            box-shadow: 0 2px 10px rgba(20, 63, 51, 0.06);
        }

        /* 11. 상단 타이틀 배너 (simulator.html <header>와 동일 톤) */
        .app-hero-banner {
            background-color: var(--sage-forest) !important;
            color: #ffffff !important;
            padding: 18px 24px;
            border-radius: 14px;
            border-bottom: 4px solid #2c8055;
            margin-bottom: 14px;
            box-shadow: 0 2px 10px rgba(20, 63, 51, 0.12);
        }
        .app-hero-banner .app-hero-sub {
            color: #c2d7cb !important;
            font-size: 12.5px;
            line-height: 1.45;
            margin: 0 0 4px 0;
        }
        .app-hero-banner .app-hero-title {
            color: #ffffff !important;
            font-size: clamp(1.35rem, 2.2vw, 1.85rem);
            font-weight: 700;
            line-height: 1.3;
            letter-spacing: -0.3px;
            margin: 0;
            padding: 0;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ==============================================================================
# Streamlit UI 구성
# ==============================================================================
st.set_page_config(
    page_title="토마토 스마트팜 ML 학습 · 시뮬레이터",
    page_icon="🍅",
    layout="wide",
)
apply_sage_theme_css()
ensure_session_state()

st.markdown(
    """
    <div class="app-hero-banner">
        <div class="app-hero-sub">여러 개의 CSV · 엑셀(.xlsx, .xls) 실측 파일을 업로드하고 버튼 한 번으로 머신러닝 학습 및 시뮬레이터 연동을 실행합니다.</div>
        <h1 class="app-hero-title">🍅 토마토 스마트팜 ML 학습 및 수확량 예측 시뮬레이터</h1>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.expander("ℹ️ Streamlit Cloud 이용 안내 (파일 유지 및 저장 방법)", expanded=False):
    st.markdown(
        """
        - **다중 파일 지원**: CSV(`.csv`, UTF-8/CP949/EUC-KR 자동 인식)와 엑셀(`.xlsx`, `.xls`) 파일을 여러 개 동시에 올릴 수 있습니다.
        - **임시 저장소 주의**: Streamlit Cloud는 앱이 절전(Sleep)되거나 재시작되면 서버의 임시 파일이 초기화될 수 있습니다.
          학습 완료 후 아래 **「ml_pred.json 다운로드」** 버튼으로 결과 파일을 PC에 보관하시는 것을 권장합니다.
        - **시뮬레이터 연동**: 학습이 완료되면 **「🌱 토마토 스마트팜 시뮬레이터」** 탭의 주별 그래프에 방금 생성된 ML 예측선이 자동으로 겹쳐 표시됩니다.
        """
    )

with st.sidebar:
    st.header("⚙️ ML 학습 설정")
    min_weeks = st.number_input(
        "최소 학습 실측 샘플 수 (--min-weeks)",
        min_value=1,
        max_value=52,
        value=4,
        step=1,
        help="업로드된 실측 데이터 합계가 이 수보다 적으면 학습이 중단됩니다.",
    )
    zone_opt = st.text_input(
        "대상 구역/작기 지정 (선택, --zone)",
        value="",
        placeholder="비워두면 자동 선택 (예: Zone_A)",
    )
    base_week_opt = st.number_input(
        "학습 기준 마지막 주차 (--base-week, 0=자동)",
        min_value=0,
        max_value=104,
        value=0,
        step=1,
    )
    st.divider()
    st.subheader("📎 선택 입력 (마스터플랜 · 기상 CSV)")
    plan_upload = st.file_uploader(
        "시뮬레이터 계획 JSON 또는 마스터플랜 CSV",
        type=["json", "csv"],
        accept_multiple_files=False,
        help="simulator.html에서 [계획 저장(JSON)] 또는 [마스터플랜 CSV 다운로드]로 저장한 파일이 있으면 함께 연결할 수 있습니다.",
    )
    weather_upload = st.file_uploader(
        "기상 관측 CSV (기상청 ASOS / Open-Meteo)",
        type=["csv"],
        accept_multiple_files=False,
        help="외부 일사량/기온 보정이 필요할 때 선택적으로 업로드하세요.",
    )

tab_train, tab_sim = st.tabs(["📤 1. 실측 파일 업로드 · ML 학습", "🌱 2. 토마토 스마트팜 시뮬레이터 (simulator.html)"])

# ------------------------------------------------------------------------------
# TAB 1: 다중 CSV/엑셀 업로드 · 파일 관리 · ML 학습 실행 · 결과 확인/다운로드
# ------------------------------------------------------------------------------
with tab_train:
    col_up, col_ctrl = st.columns([3, 1])
    with col_up:
        new_uploads = st.file_uploader(
            "실측 수확량·환경 데이터 파일 업로드 (CSV / Excel 여러 개 선택 가능)",
            type=["csv", "xlsx", "xls"],
            accept_multiple_files=True,
            key=f"calib_uploader_{st.session_state['uploader_key']}",
        )
        if new_uploads:
            added_cnt = 0
            for uf in new_uploads:
                safe_name = sanitize_filename(uf.name)
                raw_bytes = uf.getvalue()
                prev = st.session_state["managed_files"].get(safe_name)
                if prev is None or prev["bytes"] != raw_bytes:
                    meta = inspect_uploaded_table(safe_name, raw_bytes)
                    st.session_state["managed_files"][safe_name] = {
                        "name": safe_name,
                        "bytes": raw_bytes,
                        "size_kb": round(len(raw_bytes) / 1024.0, 1),
                        "origin": "사용자 업로드",
                        "added_at": datetime.now().strftime("%H:%M:%S"),
                        "meta": meta,
                    }
                    added_cnt += 1
            if added_cnt > 0:
                st.success(f"{added_cnt}개 파일이 학습 목록에 반영되었습니다.")

    with col_ctrl:
        st.write("")
        st.write("")
        if st.button("🧹 업로드 목록 전체 초기화", use_container_width=True):
            st.session_state["managed_files"] = {}
            st.session_state["uploader_key"] += 1
            st.rerun()

    managed: Dict[str, Dict[str, Any]] = st.session_state["managed_files"]
    st.subheader(f"📁 현재 학습 대상 파일 목록 ({len(managed)}개)")

    if not managed:
        st.info("아직 등록된 실측 파일이 없습니다. 위 업로더에서 `.csv`, `.xlsx`, `.xls` 파일을 하나 이상 올려 주세요.")
    else:
        summary_rows = []
        for fname, item in managed.items():
            meta = item["meta"]
            cols_preview = ", ".join(meta["columns"][:6])
            if len(meta["columns"]) > 6:
                cols_preview += f" 외 {len(meta['columns']) - 6}개"
            summary_rows.append({
                "파일명": fname,
                "구분": item["origin"],
                "형식": meta["file_type"],
                "시트": meta["active_sheet"] or "-",
                "행 수": meta["row_count"],
                "컬럼 수": meta["col_count"],
                "주요 컬럼": cols_preview or "-",
                "크기(KB)": item["size_kb"],
            })
        st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

        # 개별 파일 미리보기 및 제외 버튼
        with st.expander("🔍 업로드 파일별 상세 미리보기 및 개별 제외", expanded=False):
            for fname in list(managed.keys()):
                item = managed[fname]
                meta = item["meta"]
                c1, c2 = st.columns([5, 1])
                with c1:
                    sheet_note = f" · 시트: `{meta['active_sheet']}`" if meta["active_sheet"] else ""
                    st.markdown(f"**{fname}** ({meta['file_type']}{sheet_note} · {meta['row_count']}행 × {meta['col_count']}열)")
                with c2:
                    if st.button("🗑️ 제외", key=f"del_{fname}", use_container_width=True):
                        del st.session_state["managed_files"][fname]
                        st.session_state["uploader_key"] += 1
                        st.rerun()
                if meta.get("error"):
                    st.warning(meta["error"])
                elif meta.get("preview_df") is not None:
                    st.dataframe(meta["preview_df"], use_container_width=True, hide_index=True)
                st.divider()

    st.divider()
    run_col1, run_col2 = st.columns([2, 3])
    with run_col1:
        run_clicked = st.button(
            "🚀 ML 학습 실행 (train_export_ml.py)",
            type="primary",
            use_container_width=True,
            disabled=(len(managed) == 0 and plan_upload is None),
        )
    with run_col2:
        st.caption("버튼을 누르면 목록의 모든 CSV/엑셀 파일을 병합·중복제거한 뒤 RandomForest 모델로 학습하여 `ml_pred.json`을 생성합니다.")

    if run_clicked:
        with st.spinner("실측 파일 병합 및 머신러닝(RandomForest) 학습 실행 중..."):
            work_calib_dir = sync_managed_files_to_workdir()
            session_out_json = Path(st.session_state["session_work_dir"]) / "ml_pred.json"
            ok, cmd_str, log_text, payload = run_training_subprocess(
                calib_dir=work_calib_dir,
                out_json_path=session_out_json,
                min_weeks=int(min_weeks),
                zone_opt=zone_opt,
                base_week_opt=int(base_week_opt),
                plan_file=plan_upload,
                weather_file=weather_upload,
            )
            st.session_state["last_run_cmd"] = cmd_str
            st.session_state["last_run_log"] = log_text
            if ok and payload is not None:
                st.session_state["ml_pred_payload"] = payload
                st.session_state["last_run_error"] = ""
                st.success("✅ ML 학습 및 `ml_pred.json` 생성이 완료되었습니다! 「2. 토마토 스마트팜 시뮬레이터」 탭에서도 바로 확인할 수 있습니다.")
            else:
                st.session_state["last_run_error"] = log_text
                st.error("❌ 학습 실행 중 오류가 발생했습니다. 아래 실행 로그를 확인해 주세요.")

    if st.session_state["last_run_log"] or st.session_state["last_run_error"]:
        with st.expander("📋 실행 명령어 및 콘솔 로그 (stdout / stderr)", expanded=bool(st.session_state["last_run_error"])):
            if st.session_state["last_run_cmd"]:
                st.code(st.session_state["last_run_cmd"], language="bash")
            st.code(st.session_state["last_run_log"] or st.session_state["last_run_error"], language="text")

    # 학습 결과 요약 및 다운로드
    payload_cur = st.session_state.get("ml_pred_payload")
    if isinstance(payload_cur, dict) and payload_cur.get("predictions"):
        st.subheader("📊 ML 학습 결과 요약 (ml_pred.json)")
        files_used = payload_cur.get("calib_files_used") or []
        preds = payload_cur.get("predictions") or []
        mae_v = payload_cur.get("mae")
        rmse_v = payload_cur.get("rmse")

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("구역 / 작기", str(payload_cur.get("zone") or "-"))
        m2.metric("사용된 파일 수", f"{len(files_used)}개")
        m3.metric("학습 샘플 수", f"{payload_cur.get('train_samples', payload_cur.get('matched_obs_weeks', 0))}건")
        m4.metric("MAE (kg/㎡)", f"{mae_v:.4f}" if isinstance(mae_v, (int, float)) else "-")
        m5.metric("RMSE (kg/㎡)", f"{rmse_v:.4f}" if isinstance(rmse_v, (int, float)) else "-")

        dl_col1, dl_col2 = st.columns([1, 2])
        with dl_col1:
            json_str = json.dumps(payload_cur, ensure_ascii=False, indent=2)
            st.download_button(
                label="📥 ml_pred.json 다운로드",
                data=json_str.encode("utf-8"),
                file_name="ml_pred.json",
                mime="application/json",
                use_container_width=True,
            )
        with dl_col2:
            if files_used:
                st.caption(f"반영된 파일: {', '.join(files_used)}")

        pred_df = pd.DataFrame(preds)
        chart_cols = [c for c in ("y_pred", "y_actual", "adj_kg") if c in pred_df.columns]
        if chart_cols and "week" in pred_df.columns:
            chart_df = pred_df.set_index("week")[chart_cols].rename(
                columns={"y_pred": "ML 예측 (y_pred)", "y_actual": "실측 (y_actual)", "adj_kg": "물리 생리보정 (adj_kg)"}
            )
            st.line_chart(chart_df, use_container_width=True)

        with st.expander("📑 주차별 예측 테이블 및 JSON 원문 보기", expanded=False):
            st.dataframe(pred_df, use_container_width=True, hide_index=True)
            st.json(payload_cur)


# ------------------------------------------------------------------------------
# TAB 2: 기존 public/simulator.html 임베드 및 생성된 ml_pred.json 자동 연동
# ------------------------------------------------------------------------------
with tab_sim:
    payload_cur = st.session_state.get("ml_pred_payload")
    if payload_cur:
        st.success(
            f"✅ 학습 결과(`ml_pred.json`, 구역: {payload_cur.get('zone', '-')}, "
            f"{len(payload_cur.get('predictions', []))}주)가 아래 시뮬레이터 그래프에 자동 연동되었습니다. "
            "(필요 시 시뮬레이터 내부 [ML JSON 불러오기] / [JSON 파일 선택] 버튼으로도 다시 불러올 수 있습니다.)"
        )
    else:
        st.info(
            "💡 아직 학습된 `ml_pred.json`이 없습니다. 먼저 **「1. 실측 파일 업로드 · ML 학습」** 탭에서 학습을 실행하거나, "
            "시뮬레이터 내부에서 직접 실측 CSV 또는 `ml_pred.json` 파일을 불러와 보세요."
        )

    embedded_html = build_simulator_embed_html(payload_cur)
    components.html(embedded_html, height=1380, scrolling=True)
