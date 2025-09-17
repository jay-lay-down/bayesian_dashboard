# -*- coding: utf-8 -*-

# app.py — Bayesian Journey Dashboard (Colab-friendly, robust Excel + plots)
# Fixes:
#  - ✅ 정규화 유틸(_as_all, _ensure_key_cols 등) 포함
#  - ✅ pick_row_for 포함
#  - ✅ Plotly 축 그리드 속성 정리(유효하지 않은 prop 제거)
#  - ✅ 포트 충돌 시 자동 대체 포트로 재시도

import os, json, re, traceback
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from dash import Dash, html, dcc, dash_table, Input, Output, State
from dash.dash_table import FormatTemplate
from dash.dash_table.Format import Format, Scheme
import dash  # (NEW) 인터랙션 로그용

# (파일 상단 import 근처에 추가)
import io

import hashlib

FLOW_SALT = os.getenv("FLOW_SALT", "phi-v1-2025-01")  # 필요시 환경변수로 바꿔치기 가능
FLOW_SALT = os.getenv("FLOW_SALT", "phi-v1-2025-01")
FLOW_GLOBAL = True        # True면 전역 고정, False면 해시 기반
GLOBAL_K = 11.3

def _flow_scale(seg, mod, loy):
    if FLOW_GLOBAL:
        return GLOBAL_K
    key = f"{seg}|{mod}|{loy}|{FLOW_SALT}"
    h = int(hashlib.sha256(key.encode("utf-8")).hexdigest()[:8], 16)
    return 7.5 + (h % 1100) / 100.0

# ======== 인터랙션 공용 설정 ========
GRAPH_CONFIG = {
    "displayModeBar": True,
    "scrollZoom": True,          # 휠로 줌
    "doubleClick": "reset",      # 더블클릭 리셋
    "modeBarButtonsToAdd": ["lasso2d", "select2d"],
    "showTips": True,
}

# ===================== 기본 경로 =====================
DEFAULT_PATH = r"/content/drive/MyDrive/baye_dash/bayesian_analysis_total_v1.xlsx"

# ===================== 레벨 상수 =====================
LEVEL_OVERALL="전체"; LEVEL_SEGMENT="세그먼트"; LEVEL_MODEL="모델"
LEVEL_LOYALTY="충성도"; LEVEL_SEG_X_LOY="세그×충성도"
LEVEL_SEG_X_MODEL="세그×모델"; LEVEL_MODEL_X_LOY="모델×충성도"
LEVEL_MOD_X_SEG_X_LOY="모델×세그×충성도"

# === 정규화 ===
ALL_ALIASES = {"ALL","all","All","", " ", "  ", "전체", "NONE","None","none","nan","NaN", None}
LVL_ALIASES = {
    "모델전체×세그×충성도": "모델×세그×충성도",
    "세그x모델": "세그×모델",
    "모델x충성도": "모델×충성도",
    "세그x충성도": "세그×충성도",
}

def _as_all(v) -> str:
    s = "ALL" if v is None else str(v).strip()
    return "ALL" if s in ALL_ALIASES else s

def _ensure_key_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in ["analysis_level","segment","model","loyalty"]:
        if c not in df.columns:
            df[c] = "ALL"
        df[c] = (
            df[c].astype(str).str.strip()
              .replace({
                  "": "ALL", "전체":"ALL",
                  "NONE":"ALL","None":"ALL","none":"ALL",
                  "nan":"ALL","NaN":"ALL",
                  "ALL":"ALL","All":"ALL","all":"ALL"
              })
              .fillna("ALL")
        )
    if "level" not in df.columns:
        df["level"] = df["analysis_level"] if "analysis_level" in df.columns else "전체"
    df["level"] = (
        df["level"].astype(str).str.strip()
          .replace({"ALL":"전체","All":"전체","all":"전체"})
          .replace(LVL_ALIASES)
    )
    if "analysis_level" in df.columns:
        df["analysis_level"] = df["analysis_level"].replace(LVL_ALIASES)
    return df

# ---- Store JSON 로더 & 스왑 감지 유틸 ----
def _looks_split_df_json(s: str) -> bool:
    try:
        o = json.loads(s)
        # orient="split"는 최소 columns/index/data 3셋이 있음
        return isinstance(o, dict) and {"columns","index","data"}.issubset(set(o.keys()))
    except Exception:
        return False

def _looks_overall_json(s: str) -> bool:
    try:
        o = json.loads(s)
        return isinstance(o, dict) and any(k in o for k in ("pref_mean","rec_mean","intent_mean","buy_mean"))
    except Exception:
        return False

def _safe_read_df_split(js: str | dict | None) -> pd.DataFrame:
    if js is None:
        return pd.DataFrame()
    if isinstance(js, dict):  # 이미 파싱된 경우
        # dict가 split 스키마인 경우만 처리
        if {"columns","index","data"}.issubset(set(js.keys())):
            return pd.read_json(io.StringIO(json.dumps(js)), orient="split")
        return pd.DataFrame()
    # str
    try:
        return pd.read_json(io.StringIO(js), orient="split")
    except Exception:
        return pd.DataFrame()

def _safe_read_overall(js: str | dict | None) -> dict:
    if js is None:
        return {}
    if isinstance(js, dict):
        return js
    try:
        o = json.loads(js)
        return o if isinstance(o, dict) else {}
    except Exception:
        return {}

def _maybe_swap_sankey_overall(js_sankey, js_overall):
    """
    sankey 캐시와 overall이 뒤바뀌어 들어온 경우 자동 교정.
    (js_sankey가 overall dict이고, js_overall이 split DF JSON인 케이스)
    """
    try:
        if isinstance(js_sankey, str) and _looks_overall_json(js_sankey) \
           and isinstance(js_overall, str) and _looks_split_df_json(js_overall):
            return js_overall, js_sankey, True  # (교정된 sankey, overall, swapped?)
    except Exception:
        pass
    return js_sankey, js_overall, False

_read_df_store = _safe_read_df_split
_read_overall  = _safe_read_overall

def _rebuild_hkey_using_level(df: pd.DataFrame) -> pd.DataFrame:
    df = _ensure_key_cols(df).copy()
    if "level" in df.columns and df["level"].notna().any():
        pass
    elif "analysis_level" in df.columns:
        df["level"] = df["analysis_level"]
    else:
        df["level"] = "전체"
    for c in ["level","segment","model","loyalty"]:
        if c != "level":
            df[c] = (
                df[c].astype(str).str.strip()
                  .replace({"": "ALL","전체":"ALL","NONE":"ALL","None":"ALL","none":"ALL","nan":"ALL","NaN":"ALL"})
                  .fillna("ALL")
            )
    df["level"] = df["level"].replace(LVL_ALIASES)
    df["hierarchy_key"] = df["level"] + "|" + df["segment"] + "|" + df["model"] + "|" + df["loyalty"]
    return df

def sample_col_in_df(df) -> str | None:
    for c in ["pref_sample_size","sample_size","n","N","base","베이스수","표본수"]:
        if c in df.columns: return c
    return None

# ==== 비공개 유량 스케일 ====
FLOW_GLOBAL = True
GLOBAL_K = 11.3

# ==== 공용: Shape-safe helpers (가짜 키 자동 차단 + 보정) ====

_ALLOWED_SHAPE_KEYS = {
    "editable","fillcolor","fillrule","label","layer","legend","legendgroup","legendgrouptitle",
    "legendrank","legendwidth","line","name","opacity","path","showlegend","templateitemname",
    "type","visible","x0","x1","xanchor","xref","xsizemode","y0","y1","yanchor","yref","ysizemode",
}

# 여기가 문제의 가짜 키들
_SHIFT_KEYS = ("x0shift", "x1shift", "y0shift", "y1shift")

def _line_from_kwargs(kwargs: dict):
    line = {}
    if "line_color" in kwargs: line["color"] = kwargs.pop("line_color")
    if "line_width" in kwargs: line["width"] = kwargs.pop("line_width")
    if "line_dash"  in kwargs: line["dash"]  = kwargs.pop("line_dash")
    return {k: v for k, v in line.items() if v is not None}

def _clean_shape_kwargs(kwargs: dict):
    """
    1) *_shift 키 제거
    2) line_* → line 병합
    3) 허용 키만 남기기
    """
    kwargs = dict(kwargs)  # shallow copy
    # 1) 가짜 shift 키 모두 제거
    for k in _SHIFT_KEYS:
        kwargs.pop(k, None)
    # 2) line_* → line 병합
    line = _line_from_kwargs(kwargs)
    if line:
        base_line = kwargs.get("line") or {}
        kwargs["line"] = {**base_line, **line}
    # 3) 허용 키만 통과
    return {k: v for k, v in kwargs.items() if (k in _ALLOWED_SHAPE_KEYS and v is not None)}

def add_vline_safe(fig, x, **kwargs):
    """세로 기준선(가짜 키 차단, line_* 병합)"""
    base = dict(
        type="line", xref="x", x0=float(x), x1=float(x),
        yref="paper", y0=0, y1=1,
        layer=kwargs.pop("layer", "above"),
    )
    if "opacity" in kwargs and kwargs["opacity"] is not None:
        base["opacity"] = kwargs.pop("opacity")
    base.update(_clean_shape_kwargs(kwargs))
    return fig.add_shape(**base)

def add_hline_safe(fig, y, **kwargs):
    """가로 기준선(가짜 키 차단, line_* 병합)"""
    base = dict(
        type="line", yref="y", y0=float(y), y1=float(y),
        xref="paper", x0=0, x1=1,
        layer=kwargs.pop("layer", "above"),
    )
    if "opacity" in kwargs and kwargs["opacity"] is not None:
        base["opacity"] = kwargs.pop("opacity")
    base.update(_clean_shape_kwargs(kwargs))
    return fig.add_shape(**base)

def add_vrect_safe(fig, x0, x1, **kwargs):
    """
    add_vrect 대체: x0shift/x1shift를 값에 반영 후 제거하고,
    나머지 키는 안전하게 정리해서 rect shape로 추가.
    """
    # ── shift 보정 ──
    dx0 = float(kwargs.pop("x0shift", 0) or 0)
    dx1 = float(kwargs.pop("x1shift", 0) or 0)
    x0 = float(x0) + dx0
    x1 = float(x1) + dx1

    # yref 자동 판정(명시가 있으면 존중)
    yref = kwargs.pop("yref", None)
    has_y = ("y0" in kwargs) or ("y1" in kwargs)
    if yref is None:
        yref = "y" if has_y else "paper"

    # paper 좌표 기본값
    y0_default, y1_default = (0, 1) if yref == "paper" else (None, None)

    base = dict(
        type="rect", xref="x", x0=x0, x1=x1,
        yref=yref, y0=kwargs.pop("y0", y0_default), y1=kwargs.pop("y1", y1_default),
        layer=kwargs.pop("layer", "below"),
        fillcolor=kwargs.pop("fillcolor", "rgba(0,0,0,0.06)"),
    )
    if base["yref"] == "y":
        # 데이터 축이면 None인 y0/y1 제거
        if base.get("y0") is None: base.pop("y0", None)
        if base.get("y1") is None: base.pop("y1", None)

    if "opacity" in kwargs and kwargs["opacity"] is not None:
        base["opacity"] = kwargs.pop("opacity")

    base.update(_clean_shape_kwargs(kwargs))
    return fig.add_shape(**base)

# (선택) 만약 어딘가에서 layout.shapes에 직접 dict를 넣는다면:
def sanitize_shape_dict(d: dict) -> dict:
    """외부/레거시 shape dict을 안전하게 정제.
       - x0shift/x1shift/y0shift/y1shift 값을 좌표에 반영하고 키 제거
       - line_* 키 병합
       - 허용되지 않는 키 삭제
    """
    d = dict(d or {})

    # 1) shift -> 좌표 반영
    for sh_key, coord_key in (("x0shift","x0"),("x1shift","x1"),("y0shift","y0"),("y1shift","y1")):
        if sh_key in d:
            try:
                if coord_key in d and d[coord_key] is not None:
                    d[coord_key] = float(d[coord_key]) + float(d.pop(sh_key) or 0.0)
                else:
                    d.pop(sh_key, None)
            except Exception:
                d.pop(sh_key, None)

    # 2) line_* -> line 병합
    line = {}
    if "line_color" in d: line["color"] = d.pop("line_color")
    if "line_width" in d: line["width"] = d.pop("line_width")
    if "line_dash"  in d: line["dash"]  = d.pop("line_dash")
    if line:
        base_line = d.get("line") or {}
        d["line"] = {**base_line, **{k:v for k,v in line.items() if v is not None}}

    # 3) 허용 키만 남기기
    return {k: v for k, v in d.items() if (k in _ALLOWED_SHAPE_KEYS and v is not None)}

def _scrub_layout_shapes(fig: go.Figure) -> go.Figure:
    """
    layout.shapes에 남아있는 비정상 키(x0shift 같은 잔재)를 일괄 제거.
    """
    try:
        shapes = list(fig.layout.shapes) if fig.layout.shapes is not None else []
        cleaned = []
        for sh in shapes:
            try:
                sd = sh.to_plotly_json() if hasattr(sh, "to_plotly_json") else dict(sh)
                cleaned.append(sanitize_shape_dict(sd))  # ← 기존 유틸 재사용
            except Exception:
                # 하나라도 문제면 그냥 건너뜀(도면 깨지지 않게)
                continue
        fig.update_layout(shapes=cleaned)
    except Exception:
        pass
    return fig


def sanitize_fig_shapes(fig):
    """fig.layout.shapes 전부 sanitize."""
    try:
        shapes = list(fig.layout.shapes) if fig.layout.shapes else []
    except Exception:
        shapes = []
    if not shapes:
        return fig
    new_shapes = []
    for sh in shapes:
        try:
            sd = sh.to_plotly_json() if hasattr(sh, "to_plotly_json") else dict(sh)
            new_shapes.append(sanitize_shape_dict(sd))
        except Exception:
            # 망가진 건 버림
            pass
    fig.update_layout(shapes=new_shapes)
    return fig

# ===================== 팔레트 =====================
COL_RED        = "#C32C2C"  # 빨강
COL_ORANGE     = "#D24D3E"  # 주황
COL_YELLOW     = "#DE937A"  # 노랑
COL_BEIGE      = "#D49442"  # 베이지
COL_GREEN_LITE = "#2B8E81"  # 초록(기본)
COL_GREEN_DARK = "#21786E"  # 초록 진한톤(필요시)
COL_GRAY       = "#D3D3D3"

def _hex_to_rgb_tuple(h):  # 유틸
    h = h.lstrip("#")
    return [int(h[i:i+2], 16) for i in (0,2,4)]

def royg_color_for(values: np.ndarray) -> list:
    v = np.asarray(values, dtype=float)
    if v.size == 0: return []
    if not np.isfinite(v).any():
        return [COL_GREEN_DARK] * len(v)

    lo = np.nanmin(v); hi = np.nanmax(v)
    t = np.zeros_like(v) if (not np.isfinite(lo) or not np.isfinite(hi) or hi-lo < 1e-12) else (v-lo)/(hi-lo)

    # 낮은값(좋음) → 높은값(나쁨): 초 → 베 → 노 → 주 → 빨
    cols = np.array([
        _hex_to_rgb_tuple(COL_GREEN_LITE),
        _hex_to_rgb_tuple(COL_BEIGE),
        _hex_to_rgb_tuple(COL_YELLOW),
        _hex_to_rgb_tuple(COL_ORANGE),
        _hex_to_rgb_tuple(COL_RED),
    ], dtype=float)
    stops = np.array([0.0, 0.25, 0.5, 0.75, 1.0])

    r = np.interp(t, stops, cols[:,0]); g = np.interp(t, stops, cols[:,1]); b = np.interp(t, stops, cols[:,2])
    out = []
    for rr, gg, bb in zip(r,g,b):
        if not (np.isfinite(rr) and np.isfinite(gg) and np.isfinite(bb)):
            out.append('rgb(140,140,140)')
        else:
            out.append(f'rgb({int(round(rr))},{int(round(gg))},{int(round(bb))})')
    return out


# ==== DESIGN CONSTANTS (tiers & neutrals) ====
COL_BLUE_DEEP = "#1E3A8A"   # 진파랑(하이엔드)
COL_BLUE_SKY  = "#60A5FA"   # 하늘(미드)
COL_GRAY_MED  = "#9CA3AF"   # 회색(로우/중립)
COL_BLACK     = "#111111"   # 포레스트 플롯용

# 세그/티어 → 색 매핑 (모든 키는 소문자 기준으로 저장)
_SEG_TIER_COLOR = {
    # High/Premium 계열
    "highend": COL_BLUE_DEEP, "high": COL_BLUE_DEEP, "premium": COL_BLUE_DEEP,
    "하이엔드": COL_BLUE_DEEP, "프리미엄": COL_BLUE_DEEP,
    # Mid 계열
    "midend": COL_BLUE_SKY, "mid": COL_BLUE_SKY, "midrange": COL_BLUE_SKY,
    "미드": COL_BLUE_SKY, "중간": COL_BLUE_SKY,
    # Low/Entry 계열
    "lowend": COL_GRAY_MED, "low": COL_GRAY_MED, "entry": COL_GRAY_MED,
    "로우엔드": COL_GRAY_MED, "저가": COL_GRAY_MED,
}

def _norm_key(x) -> str:
    return "" if x is None else str(x).strip().lower()

def _tier_color_for_segment(seg: str) -> str:
    """세그 이름을 느슨하게 받아 컬러로 매핑(대소문자/공백/한글 허용)."""
    return _SEG_TIER_COLOR.get(_norm_key(seg), COL_GRAY_MED)

def _model_dominant_segment(df_scope: pd.DataFrame) -> dict:
    """
    모델별 '표본수 가중' 우세 세그. segment가 ALL/전체인 행은 제외.
    반환: {model(str): segment(str)}
    """
    if df_scope is None or df_scope.empty or "model" not in df_scope.columns or "segment" not in df_scope.columns:
        return {}

    s = df_scope.copy()
    # ALL/전체 drop
    seg_norm = s["segment"].astype(str).str.strip()
    m_valid = ~seg_norm.isin(["ALL", "전체"]) & seg_norm.notna()
    s = s[m_valid]
    if s.empty:
        return {}

    w = pd.to_numeric(s.get("pref_sample_size", 1), errors="coerce").replace([np.inf, -np.inf], np.nan).fillna(1.0)
    s["__w__"] = w

    grp = s.groupby(["model", "segment"], as_index=False)["__w__"].sum()
    # 각 모델에서 가중치 최대인 세그 1개 선택
    dom = grp.sort_values(["model", "__w__"], ascending=[True, False]).drop_duplicates("model")
    return {str(r["model"]): str(r["segment"]) for _, r in dom.iterrows()}


# ===================== 앱 =====================
app = Dash(__name__)
app.title = "Bayesian Journey Dashboard"
px.defaults.template = "plotly_white"

def _safe_num(x, default=np.nan):
    try: return float(x)
    except Exception: return default

def _safe_int0(x):
    try:
        v = float(x)
        return int(v) if np.isfinite(v) else 0
    except Exception:
        return 0

def _norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty: return pd.DataFrame()
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    for c in df.columns:
        if df[c].dtype == "O":
            ser = pd.to_numeric(df[c], errors="coerce")
            if ser.notna().mean() >= 0.5: df[c] = ser
    return df

def _ci_to_sd(lo, hi):
    lo = np.asarray(lo, dtype=float); hi = np.asarray(hi, dtype=float)
    return (hi - lo)/(2*1.96)

def _grade_from_p(p):
    if not np.isfinite(p): return "N/A"
    if p >= 0.70: return "A"
    if p >= 0.55: return "B"
    if p >= 0.45: return "C"
    return "D"

def _auto_dtick(span):
    # 0~1 퍼센트 축 span 기준
    if span <= 0.30: return 0.05   # 5%
    if span >= 0.80: return 0.20   # 20%
    return 0.10                    # 10%

def apply_dense_grid(fig: go.Figure, x_prob: bool = False, y_prob: bool = False) -> go.Figure:
    # 1) 기존 높이 보존(없을 때만 360 지정)
    cur_h = getattr(fig.layout, "height", None)
    fig.update_layout(
        height=(cur_h if cur_h is not None else 360),
        showlegend=True,
        paper_bgcolor="#fff",
        plot_bgcolor="#fff",
        font=dict(color="#111"),
        margin=dict(l=10, r=10, t=30, b=10),
    )

    # 2) 기본 격자
    fig.update_xaxes(showgrid=True, gridcolor="#e5e7eb", gridwidth=0.8, zeroline=False)
    fig.update_yaxes(showgrid=True, gridcolor="#e5e7eb", gridwidth=0.8, zeroline=False)

    # 3) plotly 버전별 minor 옵션 안전 처리
    try:
        fig.update_xaxes(minor=dict(showgrid=False))
        fig.update_yaxes(minor=dict(showgrid=False))
    except Exception:
        pass

    # 4) 확률축(0~1) 포맷
    if x_prob:
        xr = (getattr(fig.layout.xaxis, "range", None) or [0, 1])
        span = (xr[1] - xr[0]) if isinstance(xr, (list, tuple)) and len(xr) == 2 else 1.0
        fig.update_xaxes(tick0=0, dtick=_auto_dtick(span), tickformat=".0%")
    if y_prob:
        yr = (getattr(fig.layout.yaxis, "range", None) or [0, 1])
        span = (yr[1] - yr[0]) if isinstance(yr, (list, tuple)) and len(yr) == 2 else 1.0
        fig.update_yaxes(tick0=0, dtick=_auto_dtick(span), tickformat=".0%")

    # 5) 인터랙션 상태 유지
    fig.update_layout(uirevision="keep")

    # 6) 레이아웃 shape 잔재(x0shift 등) 전역 스크럽
    try:
        fig = _scrub_layout_shapes(fig)  # sanitize_shape_dict를 내부에서 활용
    except Exception:
        pass

    return fig


    # ★ 여기 추가: 모든 shape 정제
    try:
        sanitize_fig_shapes(fig)
    except Exception:
        pass

    return fig


# ---- Excel 오픈(엔진 폴백 + 디버그 수집) ----
def _open_excel_with_fallback(path: str):
    errs = []
    for eng in ["openpyxl", None, "xlrd"]:
        try:
            xls = pd.ExcelFile(path, engine=eng) if eng else pd.ExcelFile(path)
            return xls, (eng or "auto")
        except Exception as e:
            errs.append(f"{(eng or 'auto')}: {type(e).__name__}::{e}")
    raise RuntimeError("Excel open failed | " + " | ".join(errs))

def _find_sheet(xls: pd.ExcelFile, candidates):
    names = xls.sheet_names
    norm = lambda s: re.sub(r"\s+", "", str(s)).lower()
    names_norm = {norm(n): n for n in names}
    for cand in candidates:
        cn = norm(cand)
        for k, orig in names_norm.items():
            if cn in k:
                return orig
    return None

def load_excel(path: str):
    if not os.path.exists(path):
        raise FileNotFoundError(f"엑셀 파일이 없습니다: {path}")
    xls, used_engine = _open_excel_with_fallback(path)
    sheets = list(xls.sheet_names)

    sh_master = _find_sheet(xls, ["VBA마스터테이블", "마스터", "master", "mastertable", "마스터테이블"])
    sh_tm     = _find_sheet(xls, ["베이지안전이확률매트릭스", "전이확률", "transition", "matrix"])
    sh_sankey = _find_sheet(xls, ["베이지안생키다이어그램", "생키", "sankey", "flow"])

    dbg = {"engine": used_engine, "sheets": sheets,
           "matched": {"master": sh_master, "tm": sh_tm, "sankey": sh_sankey}}

    if not sh_master:
        raise ValueError(f"필수 시트(마스터) 미발견 | sheets={sheets}")

    df_master = _norm_cols(pd.read_excel(xls, sh_master))
    df_tm     = _norm_cols(pd.read_excel(xls, sh_tm)) if sh_tm else pd.DataFrame()
    df_sankey = _norm_cols(pd.read_excel(xls, sh_sankey)) if sh_sankey else pd.DataFrame()

    df_master = _rebuild_hkey_using_level(df_master)
    if not df_tm.empty: df_tm = _rebuild_hkey_using_level(df_tm)
    if not df_sankey.empty: df_sankey = _rebuild_hkey_using_level(df_sankey)

    def col(name): return df_master.get(name, pd.Series(np.nan, index=df_master.index))
    overall = {
        "pref_mean":   float(np.nanmean(col("pref_success_rate"))),
        "rec_mean":    float(np.nanmean(col("rec_success_rate"))),
        "intent_mean": float(np.nanmean(col("intent_success_rate"))),
        "buy_mean":    float(np.nanmean(col("buy_success_rate"))),
        "pref_sd":     float(np.nanmean(_ci_to_sd(col("pref_ci_lower"),   col("pref_ci_upper")))),
        "rec_sd":      float(np.nanmean(_ci_to_sd(col("rec_ci_lower"),    col("rec_ci_upper")))),
        "intent_sd":   float(np.nanmean(_ci_to_sd(col("intent_ci_lower"), col("intent_ci_upper")))),
        "buy_sd":      float(np.nanmean(_ci_to_sd(col("buy_ci_lower"),    col("buy_ci_upper")))),
    }

    seg_opts = ["ALL"] + sorted([str(v) for v in df_master["segment"].dropna().unique() if str(v)!="ALL"])
    loy_opts = ["ALL"] + sorted([str(v) for v in df_master["loyalty"].dropna().unique() if str(v)!="ALL"])
    mod_opts_all = ["ALL"] + sorted([str(v) for v in df_master["model"].dropna().unique() if str(v)!="ALL"])

    return df_master, df_tm, df_sankey, overall, seg_opts, mod_opts_all, loy_opts, dbg

# ===================== 선택/집계 로직 =====================
def pick_row_for(df_master: pd.DataFrame, seg, mod, loy):
    seg = _as_all(seg); mod = _as_all(mod); loy = _as_all(loy)
    df  = _ensure_key_cols(df_master)

    sort_col = sample_col_in_df(df)
    if sort_col is None:
        sort_col = "__tmp_n__"; df[sort_col] = 1

    def add_pref_score(sub: pd.DataFrame) -> pd.DataFrame:
        # 사용자가 ALL로 둔 차원은 ALL을 선호(=덜 구체적인 행을 상단에)
        score = 0
        if seg == "ALL": score += (sub["segment"]=="ALL").astype(int)
        if mod == "ALL": score += (sub["model"]=="ALL").astype(int)
        if loy == "ALL": score += (sub["loyalty"]=="ALL").astype(int)
        sub = sub.copy(); sub["__score__"] = score
        return sub

    chosen = (seg!="ALL") + (mod!="ALL") + (loy!="ALL")
    wanted_levels = []
    if chosen == 0:
        wanted_levels = [LEVEL_OVERALL]
    elif chosen == 1:
        if seg!="ALL": wanted_levels = [LEVEL_SEGMENT, LEVEL_OVERALL]
        if mod!="ALL": wanted_levels = [LEVEL_MODEL, LEVEL_OVERALL]
        if loy!="ALL": wanted_levels = [LEVEL_LOYALTY, LEVEL_OVERALL]
    elif chosen == 2:
        if seg!="ALL" and mod!="ALL":
            wanted_levels = [LEVEL_SEG_X_MODEL, LEVEL_SEGMENT, LEVEL_MODEL, LEVEL_OVERALL]
        elif seg!="ALL" and loy!="ALL":
            wanted_levels = [LEVEL_SEG_X_LOY, LEVEL_SEGMENT, LEVEL_LOYALTY, LEVEL_OVERALL]
        elif mod!="ALL" and loy!="ALL":
            wanted_levels = [LEVEL_MODEL_X_LOY, LEVEL_MODEL, LEVEL_LOYALTY, LEVEL_OVERALL]
    else:
        wanted_levels = [
            LEVEL_MOD_X_SEG_X_LOY, LEVEL_SEG_X_LOY, LEVEL_SEG_X_MODEL, LEVEL_MODEL_X_LOY,
            LEVEL_MODEL, LEVEL_SEGMENT, LEVEL_LOYALTY, LEVEL_OVERALL
        ]

    # 1) 레벨 우선 매칭
    for lvl in wanted_levels:
        sub = df[df["level"] == lvl]
        if seg!="ALL": sub = sub[sub["segment"] == seg]
        if mod!="ALL": sub = sub[sub["model"]   == mod]
        if loy!="ALL": sub = sub[sub["loyalty"] == loy]
        if not sub.empty:
            sub = add_pref_score(sub).sort_values(["__score__", sort_col], ascending=[False, False])
            row = sub.iloc[0]
            return row.drop(labels=[c for c in ["__score__","__tmp_n__"] if c in row.index])

    # 2) 정확 조합 실패 시, 부분조합 매칭
    sub = df.copy()
    if seg!="ALL": sub = sub[sub["segment"] == seg]
    if mod!="ALL": sub = sub[sub["model"]   == mod]
    if loy!="ALL": sub = sub[sub["loyalty"] == loy]
    if not sub.empty:
        sub = add_pref_score(sub).sort_values(["__score__", sort_col], ascending=[False, False])
        row = sub.iloc[0]
        return row.drop(labels=[c for c in ["__score__","__tmp_n__"] if c in row.index])

    # 3) 단일 컬럼만 맞는 행이라도
    for col, val in [("segment", seg), ("model", mod), ("loyalty", loy)]:
        if val != "ALL":
            sub = df[df[col]==val]
            if not sub.empty:
                sub = add_pref_score(sub).sort_values(["__score__", sort_col], ascending=[False, False])
                row = sub.iloc[0]
                return row.drop(labels=[c for c in ["__score__","__tmp_n__"] if c in row.index])

    # 4) 완전 실패 시 표본수 최대
    row = df.sort_values(sort_col, ascending=False).iloc[0]
    return row.drop(labels=[c for c in ["__score__","__tmp_n__"] if c in row.index])

# ===================== 차트/표 유틸 =====================
def _pick_sample_for_stage(r, stage_prefix: str) -> int:
    for c in [f"{stage_prefix}_sample_size", "sample_size", "n", "N", "base", "베이스수", "표본수"]:
        if c in r and pd.notna(r.get(c)):
            return _safe_int0(r.get(c))
    return _safe_int0(r.get("pref_sample_size"))

def metrics_table_row(r):
    def sd_from_ci(lo, hi):
        if pd.isna(lo) or pd.isna(hi): return np.nan
        return (hi - lo)/(2*1.96)
    rows = []
    mapping = [
        ("선호",   "pref_success_rate",   "pref_ci_lower",   "pref_ci_upper",   "pref_snr",   "pref_lift_vs_galaxy"),
        ("추천", "rec_success_rate",    "rec_ci_lower",    "rec_ci_upper",    "rec_snr",    "rec_lift_vs_galaxy"),
        ("구매의향", "intent_success_rate", "intent_ci_lower", "intent_ci_upper", "intent_snr", "intent_lift_vs_galaxy"),
        ("구매",     "buy_success_rate",    "buy_ci_lower",    "buy_ci_upper",    "buy_snr",    "buy_lift_vs_galaxy"),
    ]
    for label, m, lo, hi, snr, lift in mapping:
        mval   = _safe_num(r.get(m))
        loval  = _safe_num(r.get(lo))
        hival  = _safe_num(r.get(hi))
        snrval = _safe_num(r.get(snr))
        liftval= _safe_num(r.get(lift))
        stage_prefix = m.split("_")[0]
        rows.append(dict(
            단계=label,
            베이스수=_pick_sample_for_stage(r, stage_prefix),
            성공확률=mval, 하한=loval, 상한=hival,
            실패확률=(None if pd.isna(mval) else 1-mval),
            판정=("성공" if (np.isfinite(mval) and mval>=0.5) else ("실패" if np.isfinite(mval) else "N/A")),
            평가등급=("N/A" if not np.isfinite(mval) else ("A" if mval>=0.70 else "B" if mval>=0.55 else "C" if mval>=0.45 else "D")),
            SNR=snrval, Lift=liftval, raw평균=mval,
            raw표준편차=sd_from_ci(loval, hival)
        ))
    return pd.DataFrame(rows)

def drops_from_anywhere(row, df_tm, seg, mod, loy):
    seg = _as_all(seg); mod = _as_all(mod); loy = _as_all(loy)
    d1 = _safe_num(row.get("bayesian_dropout_pref_to_rec"))
    d2 = _safe_num(row.get("bayesian_dropout_rec_to_intent"))
    d3 = _safe_num(row.get("bayesian_dropout_intent_to_buy"))
    full = _safe_num(row.get("bayesian_full_conversion"))
    if df_tm is None or df_tm.empty:
        return d1, d2, d3, full
    need = [np.isfinite(d1), np.isfinite(d2), np.isfinite(d3), np.isfinite(full)]
    if all(need): return d1, d2, d3, full
    m = pd.Series(True, index=df_tm.index)
    if "segment" in df_tm and seg!="ALL": m &= (df_tm["segment"].astype(str)==seg)
    if "model"   in df_tm and mod!="ALL": m &= (df_tm["model"].astype(str)==mod)
    if "loyalty" in df_tm and loy!="ALL": m &= (df_tm["loyalty"].astype(str)==loy)
    sub = df_tm[m].copy()
    if sub.empty: sub = df_tm.copy()
    w = pd.to_numeric(sub.get("pref_sample_size", pd.Series(1, index=sub.index)), errors="coerce").fillna(1)
    def wmean(col):
        v = pd.to_numeric(sub.get(col, pd.Series(np.nan, index=sub.index)), errors="coerce")
        if v.notna().any(): return float(np.nansum(v*w)/np.nansum(w))
        return np.nan
    d1 = d1 if np.isfinite(d1) else wmean("bayesian_dropout_pref_to_rec")
    d2 = d2 if np.isfinite(d2) else wmean("bayesian_dropout_rec_to_intent")
    d3 = d3 if np.isfinite(d3) else wmean("bayesian_dropout_intent_to_buy")
    full = full if np.isfinite(full) else wmean("bayesian_full_conversion")
    return d1, d2, d3, full

def biggest_drop_text_by_sources(row, df_tm, seg, mod, loy):
    d1, d2, d3, _ = drops_from_anywhere(row, df_tm, seg, mod, loy)
    pairs = [("선호→추천", d1), ("추천→구매의향", d2), ("구매의향→구매", d3)]
    pairs = [(n, v) for n, v in pairs if np.isfinite(v)]
    if not pairs: return "데이터 없음"
    name, val = max(pairs, key=lambda x: x[1])
    base_n = _safe_int0(row.get("pref_sample_size"))
    return f"{name}에서 {val*100:.1f}%p 손실 (샘플 {base_n:,})"

def compose_composite_row(df_scope: pd.DataFrame) -> pd.Series:
    if df_scope is None or df_scope.empty:
        return pd.Series(dtype=float)
    s = df_scope.copy()
    w = pd.to_numeric(s.get("pref_sample_size", pd.Series(1, index=s.index)), errors="coerce").fillna(1.0)
    w_sum = float(np.nansum(w)) if np.isfinite(np.nansum(w)) and np.nansum(w) > 0 else 1.0
    w_norm = w / w_sum
    def wmean(col):
        v = pd.to_numeric(s.get(col, pd.Series(np.nan, index=s.index)), errors="coerce")
        if v.notna().any(): return float(np.nansum(v * w_norm))
        return np.nan
    def combine_ci(lo_col, hi_col, mean_col):
        m = pd.to_numeric(s.get(mean_col, pd.Series(np.nan, index=s.index)), errors="coerce")
        lo = pd.to_numeric(s.get(lo_col, pd.Series(np.nan, index=s.index)), errors="coerce")
        hi = pd.to_numeric(s.get(hi_col, pd.Series(np.nan, index=s.index)), errors="coerce")
        if not (m.notna().any() and lo.notna().any() and hi.notna().any()):
            return np.nan, np.nan
        m_bar = float(np.nansum(m * w_norm))
        sd = (hi - lo) / (2 * 1.96)
        sd = pd.to_numeric(sd, errors="coerce")
        var = np.nansum(w_norm * (sd**2 + (m - m_bar)**2))
        sd_c = float(np.sqrt(var)) if np.isfinite(var) else np.nan
        if not np.isfinite(sd_c): return np.nan, np.nan
        return (m_bar - 1.96 * sd_c), (m_bar + 1.96 * sd_c)
    pref_m   = wmean("pref_success_rate")
    rec_m    = wmean("rec_success_rate")
    intent_m = wmean("intent_success_rate")
    buy_m    = wmean("buy_success_rate")
    pref_lo, pref_hi     = combine_ci("pref_ci_lower",   "pref_ci_upper",   "pref_success_rate")
    rec_lo, rec_hi       = combine_ci("rec_ci_lower",    "rec_ci_upper",    "rec_success_rate")
    intent_lo, intent_hi = combine_ci("intent_ci_lower", "intent_ci_upper", "intent_success_rate")
    buy_lo, buy_hi       = combine_ci("buy_ci_lower",    "buy_ci_upper",    "buy_success_rate")
    d1 = wmean("bayesian_dropout_pref_to_rec")
    d2 = wmean("bayesian_dropout_rec_to_intent")
    d3 = wmean("bayesian_dropout_intent_to_buy")
    full = wmean("bayesian_full_conversion")
    pref_snr = wmean("pref_snr");  rec_snr = wmean("rec_snr")
    intent_snr = wmean("intent_snr"); buy_snr = wmean("buy_snr")
    pref_lift = wmean("pref_lift_vs_galaxy"); rec_lift = wmean("rec_lift_vs_galaxy")
    intent_lift = wmean("intent_lift_vs_galaxy"); buy_lift = wmean("buy_lift_vs_galaxy")
    out = {
        "pref_sample_size": float(np.nansum(w)),
        "pref_success_rate": pref_m, "pref_ci_lower": pref_lo, "pref_ci_upper": pref_hi,
        "rec_success_rate": rec_m, "rec_ci_lower": rec_lo, "rec_ci_upper": rec_hi,
        "intent_success_rate": intent_m, "intent_ci_lower": intent_lo, "intent_ci_upper": intent_hi,
        "buy_success_rate": buy_m, "buy_ci_lower": buy_lo, "buy_ci_upper": buy_hi,
        "bayesian_dropout_pref_to_rec": d1,
        "bayesian_dropout_rec_to_intent": d2,
        "bayesian_dropout_intent_to_buy": d3,
        "bayesian_full_conversion": full,
        "pref_snr": pref_snr, "rec_snr": rec_snr, "intent_snr": intent_snr, "buy_snr": buy_snr,
        "pref_lift_vs_galaxy": pref_lift, "rec_lift_vs_galaxy": rec_lift,
        "intent_lift_vs_galaxy": intent_lift, "buy_lift_vs_galaxy": buy_lift,
    }
    return pd.Series(out)

# ===================== 차트 =====================
def _empty_fig(msg="Load data first", height=360, hide_axes=False):
    fig = go.Figure()
    fig.add_annotation(text=msg, x=0.5, y=0.5, xref="paper", yref="paper", showarrow=False)
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=30, b=10),
        paper_bgcolor="#ffffff",
        plot_bgcolor="#ffffff",
        uirevision="keep",
    )
    fig = apply_dense_grid(fig)  # 기존 스타일 유지

    if hide_axes:  # Sankey 등 카테시안 축이 불필요한 경우
        fig.update_xaxes(visible=False, showgrid=False, zeroline=False)
        fig.update_yaxes(visible=False, showgrid=False, zeroline=False)

    return fig

def hex_to_rgba(hex_color: str, a: float | None = None) -> str:
    s = hex_color.strip().lstrip("#")
    if len(s) in (3, 4):
        s = "".join(ch * 2 for ch in s)
    if len(s) == 6:
        r = int(s[0:2], 16); g = int(s[2:4], 16); b = int(s[4:6], 16)
        alpha = 1.0 if a is None else float(a)
    elif len(s) == 8:
        r = int(s[0:2], 16); g = int(s[2:4], 16); b = int(s[4:6], 16)
        hex_alpha = int(s[6:8], 16) / 255.0
        alpha = hex_alpha if a is None else float(a)
    else:
        raise ValueError("hex must be #RGB, #RRGGBB, or #RRGGBBAA")
    alpha = max(0.0, min(1.0, alpha))
    return f"rgba({r},{g},{b},{alpha:.3g})"


def _normalize_stage_label(v: str) -> str | None:
    if v is None:
        return None
    s = str(v).strip().lower()
    s = re.sub(r'[\s\-\_]+', ' ', s)           # 공백/-,_ 정리
    joined = s.replace(' ', '')

    # 전체
    if any(k in (s, joined) for k in [
        "overall","total","all","전체","전체사용자","모든사용자","allusers","all user","all-user"
    ]):
        return "전체"

    # 미선호(비선호/탈락/드랍/No preference 등)
    if any(k in (s, joined) for k in [
        "미선호","비선호","선호아님","선호 아님",
        "nopref","no preference","dislike","탈락","drop","dropped"
    ]):
        return "미선호"

    # 구매의향(의향/의도/의사/intent 계열)
    if ("의향" in s) or ("의도" in s) or ("의사" in s) \
       or ("intent" in s) or ("intention" in s) \
       or ("purchaseintent" in joined) or ("purchase-intent" in s):
        return "구매의향"

    # 구매(실제구매/구매완료/구매확정/구입/결제/매출/buy/purchase)
    if ("구매" in s) or ("구입" in s) or ("결제" in s) or ("결재" in s) or ("매출" in s) \
       or (s == "buy") or ("purchase" in s):
        return "구매"

    # 선호
    if ("선호" in s) or ("호감" in s) or ("preference" in s) or (s == "pref"):
        return "선호"

    # 추천
    if (s == "rec") or ("recommend" in s) or ("추천" in s):
        return "추천"

    return None

# ==== STAGES & ORDER (기존 것을 교체) ====
STAGES = ["전체", "미선호", "선호", "추천", "구매의향", "구매"]
ORDER  = {v:i for i,v in enumerate(STAGES)}

# 색상 하나 추가(은은한 회색 계열 권장)
COL_STAGE_DROP = "#CBD5E1"  # 미선호

def _group_forward_flows(df_sankey, seg, mod, loy):
    if df_sankey is None or df_sankey.empty:
        return pd.DataFrame(columns=["from_stage","to_stage","count","flow_phi"])
    seg = _as_all(seg); mod = _as_all(mod); loy = _as_all(loy)
    s = df_sankey.copy()
    m = pd.Series(True, index=s.index)
    if "segment" in s and seg!="ALL": m &= (s["segment"].astype(str)==seg)
    if "model"   in s and mod!="ALL": m &= (s["model"].astype(str)==mod)
    if "loyalty" in s and loy!="ALL": m &= (s["loyalty"].astype(str)==loy)
    s = s[m].copy()
    if s.empty: 
        return pd.DataFrame(columns=["from_stage","to_stage","count","flow_phi"])

    alias = {
        "all":"전체","ALL":"전체","전체":"전체",
        "pref":"선호","preference":"선호","선호도":"선호",
        "rec":"추천","recommend":"추천","추천도":"추천",
        "intent":"구매의향","intention":"구매의향","구매의도":"구매의향",
        "purchase":"구매","buy":"구매","실제구매":"구매"
    }
    s["from_stage"] = s.get("from_stage", s.get("from", s.get("source"))).astype(str).str.strip().replace(alias)
    s["to_stage"]   = s.get("to_stage",   s.get("to",   s.get("target"))).astype(str).str.strip().replace(alias)

    # 🔑 count 별칭 허용
    cnt_cands = ["bayesian_flow_count","count","value","weight","n","freq"]
    cnt_col = next((c for c in cnt_cands if c in s.columns), None)
    if cnt_col is None:
        return pd.DataFrame(columns=["from_stage","to_stage","count","flow_phi"])

    s[cnt_col] = pd.to_numeric(s[cnt_col], errors="coerce")
    s = s[np.isfinite(s[cnt_col]) & (s[cnt_col]>0)]
    s = s[s["from_stage"].isin(STAGES) & s["to_stage"].isin(STAGES)]
    s = s[s.apply(lambda r: ORDER[r["from_stage"]] < ORDER[r["to_stage"]], axis=1)]
    if s.empty:
        return pd.DataFrame(columns=["from_stage","to_stage","count","flow_phi"])

    g = (s.groupby(["from_stage","to_stage"], as_index=False)[cnt_col]
           .sum().rename(columns={cnt_col:"count"}))

    # [유입 없는 단계 보강] 전체→단계 링크 자동 추가
    pairs = set(zip(g["from_stage"], g["to_stage"]))
    def _has_incoming(stage):
        k = ORDER[stage]
        return any((prev, stage) in pairs for prev in STAGES[:k])

    add_rows = []
    for st in STAGES[1:]:
        if not _has_incoming(st):
            out_sum = float(g.loc[g["from_stage"] == st, "count"].sum())
            if out_sum > 0:
                add_rows.append({"from_stage": "전체", "to_stage": st, "count": out_sum})
    if add_rows:
        g = pd.concat([g, pd.DataFrame(add_rows)], ignore_index=True)

    # φ 스케일 적용
    k = _flow_scale(seg, mod, loy)
    g["flow_phi"] = g["count"].astype(float) * k
    return g

# ===== Sankey 내부용 테이블 빌더(간접 포함, 구매로 접기 옵션) =====

# 노드(베이지) & 링크(회색) 팔레트
COL_STAGE_OVERALL = "#B68E5C"   # 전체
COL_STAGE_PREF    = "#C6955E"   # 선호
COL_STAGE_REC     = "#D5A86D"   # 추천
COL_STAGE_INTENT  = "#BE8F4E"   # 의향
COL_STAGE_BUY     = "#A97F45"   # 구매
COL_LINK_DIRECT   = "#4B5563"   # 직접(짙은 회색)
COL_LINK_INDIRECT = "#D1D5DB"   # 간접(연한 회색)

def _sankey_build_table(df_sankey, seg="ALL", mod="ALL", loy="ALL",
                        collapse_to_buy=True, collapse_from=("선호","추천","구매의향")) -> pd.DataFrame:
    if df_sankey is None or df_sankey.empty:
        return pd.DataFrame(columns=["from_stage","to_stage","count","dist","kind","to_buy","flow_phi"])

    s = df_sankey.copy()

    # --- [NEW] 호환 가드: 열 별칭을 표준 이름으로 통일 ---
    # 1) from/to 별칭 → from_stage/to_stage
    from_col = next((c for c in ["from_stage","from","source","src"] if c in s.columns), None)
    to_col   = next((c for c in ["to_stage","to","target","dst"]     if c in s.columns), None)
    if from_col and from_col != "from_stage":
        s = s.rename(columns={from_col: "from_stage"})
    if to_col and to_col != "to_stage":
        s = s.rename(columns={to_col: "to_stage"})

    # 필수 열 없으면 빈 테이블 반환 (안전 가드)
    if "from_stage" not in s.columns or "to_stage" not in s.columns:
        return pd.DataFrame(columns=["from_stage","to_stage","count","dist","kind","to_buy","flow_phi"])

    # 2) 수치 열 별칭 → bayesian_flow_count
    alt_cnt = next((c for c in ["bayesian_flow_count","count","value","flow","weight","n","freq"]
                    if c in s.columns), None)
    if alt_cnt and alt_cnt != "bayesian_flow_count":
        s = s.rename(columns={alt_cnt: "bayesian_flow_count"})


    # 필터
    for col, val in (("segment", seg), ("model", mod), ("loyalty", loy)):
        if col in s.columns and str(val) != "ALL":
            s = s[s[col].astype(str) == str(val)]
    if s.empty:
        return pd.DataFrame(columns=["from_stage","to_stage","count","dist","kind","to_buy","flow_phi"])

    # 라벨 정규화 → 순방향만
    s["from_stage"] = s.get("from_stage", s.get("from", s.get("source"))).map(_normalize_stage_label)
    s["to_stage"]   = s.get("to_stage",   s.get("to",   s.get("target"))).map(_normalize_stage_label)
    s = s.dropna(subset=["from_stage","to_stage"])
    s = s[s["from_stage"].isin(STAGES) & s["to_stage"].isin(STAGES)]
    s = s[s.apply(lambda r: ORDER[r["from_stage"]] < ORDER[r["to_stage"]], axis=1)]
    
# 🔑 count 컬럼 별칭 허용 (원천 시트/캐시 시트 모두 커버)
    cnt_cands = ["bayesian_flow_count", "count", "value", "weight", "n", "freq"]
    cnt_col = next((c for c in cnt_cands if c in s.columns), None)
    if cnt_col is None:
        return pd.DataFrame(columns=["from_stage","to_stage","count","dist","kind","to_buy","flow_phi"])

    s[cnt_col] = pd.to_numeric(s[cnt_col], errors="coerce")
    s = s[np.isfinite(s[cnt_col]) & (s[cnt_col] > 0)]
    if s.empty:
        return pd.DataFrame(columns=["from_stage","to_stage","count","dist","kind","to_buy","flow_phi"])

    # 기본 집계
    g = (s.groupby(["from_stage","to_stage"], as_index=False)[cnt_col]
           .sum().rename(columns={cnt_col:"count"}))

    # 유입 없는 단계 보강(전체→단계)
    pairs = set(zip(g["from_stage"], g["to_stage"]))
    def _has_incoming(stage):
        k = ORDER[stage]
        return any((prev, stage) in pairs for prev in STAGES[:k])
    add_rows = []
    for st in STAGES[1:]:
        if not _has_incoming(st):
            out_sum = float(g.loc[g["from_stage"]==st, "count"].sum())
            if out_sum > 0:
                add_rows.append({"from_stage":"전체","to_stage":st,"count":out_sum})
    if add_rows:
        g = pd.concat([g, pd.DataFrame(add_rows)], ignore_index=True)

    # (옵션) 구매로 접은 간접 링크 추가: 선호/추천/구매의향 → 구매
    if collapse_to_buy:
        buy_in = float(pd.to_numeric(g.loc[g["to_stage"]=="구매","count"], errors="coerce").fillna(0).sum())
        if buy_in > 0:
            exist = set(zip(g["from_stage"], g["to_stage"]))
            extra = []
            for st in collapse_from:
                if st in ORDER and (st, "구매") not in exist and ORDER[st] < ORDER["구매"]:
                    extra.append({"from_stage": st, "to_stage": "구매", "count": buy_in})
            if extra:
                g = pd.concat([g, pd.DataFrame(extra)], ignore_index=True)

    # 메타 칼럼
    kphi = _flow_scale(seg, mod, loy)  # 비공개 스케일
    g["flow_phi"] = g["count"].astype(float) * kphi
    g["dist"]     = g["to_stage"].map(ORDER) - g["from_stage"].map(ORDER)
    g["kind"]     = np.where(g["dist"]==1, "직접", "간접")
    g["to_buy"]   = (g["to_stage"] == "구매")

    cols = ["from_stage","to_stage","count","dist","kind","to_buy","flow_phi"]
    return g[cols].sort_values(["dist","from_stage","to_stage"]).reset_index(drop=True)


# ====== Sankey 색/스테이지 ======
STAGES = ["전체","미선호","선호","추천","구매의향","구매"]
ORDER  = {v:i for i,v in enumerate(STAGES)}

COL_STAGE_OVERALL = "#B68E5C"
COL_STAGE_NONPREF = "#9CA3AF"  # ← 미선호(회색)
COL_STAGE_PREF    = "#C6955E"
COL_STAGE_REC     = "#D5A86D"
COL_STAGE_INTENT  = "#BE8F4E"
COL_STAGE_BUY     = "#A97F45"

COL_LINK_DIRECT   = "#4B5563"   # 짙은 회색 (직접)
COL_LINK_INDIRECT = "#D1D5DB"   # 연한 회색 (간접)
# ─────────────────────────────────────────────────────────
# (유틸) 간접 "→구매" 접기 보강
def add_collapsed_to_buy(tbl: pd.DataFrame, add_from=("선호","추천","구매의향")) -> pd.DataFrame:
    if tbl is None or tbl.empty:
        return tbl

    # ── 기준 단계/순서(미선호 포함 6단계)
    stages = ["전체","미선호","선호","추천","구매의향","구매"]
    order  = {v:i for i,v in enumerate(stages)}
    t = tbl.copy()

    # ── 구매 유입 총량
    buy_in = float(pd.to_numeric(t.loc[t["to_stage"]=="구매","count"], errors="coerce").fillna(0).sum())

    # ── φ 스케일(k) 추정
    kphi = 1.0
    if "flow_phi" in t.columns and "count" in t.columns:
        r = pd.to_numeric(t["flow_phi"], errors="coerce") / pd.to_numeric(t["count"], errors="coerce")
        r = r.replace([np.inf,-np.inf], np.nan).dropna()
        if not r.empty:
            kphi = float(np.median(r))

    # ── 그룹 메타(snapshot): 단일값이면 그 값, 아니면 "ALL"
    meta_cols = [c for c in ["segment","model","loyalty","level"] if c in t.columns]
    meta = {c: (t[c].dropna().iloc[0] if t[c].nunique(dropna=True)==1 else "ALL") for c in meta_cols}

    extra = []
    for s in add_from:
        if s not in order or order[s] >= order["구매"]:
            continue
        # 이미 존재하면 중복 추가 금지
        if ((t["from_stage"]==s) & (t["to_stage"]=="구매")).any():
            continue
        row = {
            "from_stage": s,
            "to_stage":   "구매",
            "count":      buy_in,
            "dist":       order["구매"] - order[s],
            "kind":       ("간접" if (order["구매"] - order[s]) > 1 else "직접"),
            "to_buy":     True,
            "flow_phi":   buy_in * kphi
        }
        # ★ 메타 동봉
        for c, v in meta.items():
            row[c] = v
        extra.append(row)

    if extra:
        t = pd.concat([t, pd.DataFrame(extra)], ignore_index=True)

    return t.sort_values(["dist","from_stage","to_stage"]).reset_index(drop=True)
# ─────────────────────────────────────────────────────────

# ⬇⬇ 핵심 수정: 라벨을 먼저 느슨한 별칭으로 치환 후, 정규화 함수에 태움
def _normalize_stage_soft(series: pd.Series) -> pd.Series:
    if series.empty:
        return series
    s = series.astype(str).str.strip()

    # 1) 강제 별칭(정확치환) — 의향/의도/의사/intent, 구매완료/실제구매, 전체사용자 등
    alias_exact = {
        # 전체
        "전체사용자": "전체", "모든 사용자": "전체", "all": "전체", "ALL": "전체",
        # 선호
        "선호도": "선호", "선호도높음": "선호", "호감도": "선호", "호감도높음": "선호",
        # 추천
        "추천도": "추천", "추천도높음": "추천",
        # 의향/의도/의사/intent (다양형)
        "구매의향": "구매의향", "구매 의향": "구매의향", "구매의향높음": "구매의향", "구매의향 높음": "구매의향",
        "구매의도": "구매의향", "구매 의도": "구매의향", "구매의도높음": "구매의향", "구매의도 높음": "구매의향",
        "구매의사": "구매의향", "의사 있음": "구매의향",
        "intent": "구매의향", "Intent": "구매의향", "Intention": "구매의향",
        "Purchase Intent": "구매의향", "PURCHASE_INTENT": "구매의향",
        # 구매
        "실제구매": "구매", "구매 확정": "구매", "구매확정": "구매", "구매 완료": "구매", "구매완료": "구매",
        "결제": "구매", "결재": "구매", "매출": "구매",
        #미선호
        "미선호": "미선호", "비선호": "미선호", "선호 아님": "미선호", "탈락": "미선호",
    }
    s = s.replace(alias_exact)

    # 2) 토큰/부분일치 기반 정규화(전역 함수가 있으면 재사용)
    def _norm_one(x: str) -> str | None:
        try:
            return _normalize_stage_label(x)  # 전역 정의 존재 시 활용
        except Exception:
            pass
        # 폴백: 부분일치
        xl = x.lower().replace(" ", "")
        if any(k in xl for k in ["all","전체"]): return "전체"
        if any(k in xl for k in ["선호","호감"]): return "선호"
        if "추천" in xl or "rec" in xl: return "추천"
        if any(k in xl for k in ["의향","의도","의사","intent"]): return "구매의향"
        if any(k in xl for k in ["구매","구입","결제","결재","완료","확정","매출","purch","buy"]): return "구매"
        if any(k in xl for k in ["미선호","비선호","선호아님","nopref","npreference","탈락","drop"]): return "미선호"
        return None

    return s.map(_norm_one)


# 파일 상단 어딘가(상수들 근처)에 추가
LVL_PRIORITY = [
    "모델×세그×충성도","세그×모델","모델×충성도","세그×충성도",
    "모델","세그먼트","충성도","전체"
]

def _sanitize_sankey_table(
    tbl: pd.DataFrame,
    seg="ALL", mod="ALL", loy="ALL",
    enforce_single_level: bool = True,
    drop_overall_if_mixed: bool = True
) -> pd.DataFrame:
    cols = ["from_stage","to_stage","count","dist","kind","to_buy","flow_phi"]
    if tbl is None or tbl.empty:
        return pd.DataFrame(columns=cols)

    t = tbl.copy()

    # (1) 선택값 필터 (있을 때만)
    for col, val in (("segment", seg), ("model", mod), ("loyalty", loy)):
        if col in t.columns and str(val) != "ALL":
            t = t[t[col].astype(str).str.strip() == str(val)]
    if t.empty:
        return pd.DataFrame(columns=cols)

    # (2) 레벨 단일화 (혼입 방지) + 과잉 드랍 완화
    original = t
    if enforce_single_level and "level" in t.columns:
        picked = None
        for lv in LVL_PRIORITY:
            cand = t[t["level"].astype(str) == lv]
            if not cand.empty:
                picked = cand; break
        if picked is not None:
            t = picked
    # 혼합이면 '전체'만 제거 (단, 전부 비면 되돌림)
    if ("level" in t.columns) and (t["level"].astype(str).nunique() > 1) and drop_overall_if_mixed:
        t2 = t[t["level"].astype(str) != "전체"]
        if not t2.empty:
            t = t2

    if t.empty:
        # 과잉 필터/드랍으로 비었으면 원본으로 되돌려 계속
        t = original.copy()

    # (3) 컬럼 별칭
    alias = {
        "from_stage": ["from_stage","from","source","src"],
        "to_stage":   ["to_stage","to","target","dst"],
        "count":      ["count","bayesian_flow_count","flow","value","weight","n","freq"],
    }
    def pick(name):
        keys = {str(c).strip().lower(): c for c in t.columns}
        for a in alias[name]:
            if a in keys: return keys[a]
        return None
    c_from = pick("from_stage"); c_to = pick("to_stage"); c_cnt = pick("count")
    if not all([c_from, c_to, c_cnt]):
        return pd.DataFrame(columns=cols)

    t = t.rename(columns={c_from:"from_stage", c_to:"to_stage", c_cnt:"count"})

    # (4) 라벨 정규화 + 순방향만
    t["from_stage"] = _normalize_stage_soft(t["from_stage"])
    t["to_stage"]   = _normalize_stage_soft(t["to_stage"])
    t = t.dropna(subset=["from_stage","to_stage"])
    t = t[t["from_stage"].isin(STAGES) & t["to_stage"].isin(STAGES)]
    t = t[t.apply(lambda r: ORDER[r["from_stage"]] < ORDER[r["to_stage"]], axis=1)]

    # (5) 수치 변환
    t["count"] = pd.to_numeric(t["count"], errors="coerce")
    t = t[np.isfinite(t["count"]) & (t["count"] > 0)]

    # (5-보강) 과도 필터로 비면 완화 모드: 단계 조건만 적용하고 수치만 보정
    if t.empty:
        t = original.rename(columns={c_from:"from_stage", c_to:"to_stage", c_cnt:"count"}).copy()
        t["from_stage"] = _normalize_stage_soft(t["from_stage"])
        t["to_stage"]   = _normalize_stage_soft(t["to_stage"])
        t = t.dropna(subset=["from_stage","to_stage"])
        t = t[t["from_stage"].isin(STAGES) & t["to_stage"].isin(STAGES)]
        t["count"] = pd.to_numeric(t["count"], errors="coerce").fillna(0)
        t = t[t["count"] > 0]
        if t.empty:
            return pd.DataFrame(columns=cols)

    # (6) 메타 보강
    t["dist"] = (t["to_stage"].map(ORDER) - t["from_stage"].map(ORDER)).astype(int)
    if "kind" not in t.columns:
        t["kind"] = np.where(t["dist"]==1, "직접", "간접")
    else:
        miss = ~t["kind"].astype(str).isin(["직접","간접"])
        t.loc[miss,"kind"] = np.where(t.loc[miss,"dist"]==1, "직접","간접")
    t["to_buy"] = (t["to_stage"]=="구매")

    # (7) φ
    kphi = _flow_scale(seg, mod, loy)
    if "flow_phi" not in t.columns:
        t["flow_phi"] = t["count"].astype(float) * kphi
    else:
        t["flow_phi"] = pd.to_numeric(t["flow_phi"], errors="coerce")
        miss = ~np.isfinite(t["flow_phi"])
        t.loc[miss, "flow_phi"] = t.loc[miss, "count"].astype(float) * kphi

    return t[cols].sort_values(["dist","from_stage","to_stage"]).reset_index(drop=True)


def _sankey_from_master_row(row: pd.Series, seg, mod, loy) -> pd.DataFrame:
    n = _safe_int0(row.get("pref_sample_size"))
    if n <= 0:
        return pd.DataFrame(columns=[
            "from_stage","to_stage","count","dist","kind","to_buy","flow_phi",
            "segment","model","loyalty"
        ])

    def P(x):
        v = _safe_num(x)
        if not np.isfinite(v): return np.nan
        return v/100.0 if v > 1.5 else v

    # (A) 확률 안전화: NaN이면 0, 0~1로 클립
    def P01(x):
        v = P(x)
        return np.nan if not np.isfinite(v) else float(min(1.0, max(0.0, v)))

    p_pref   = P(row.get("pref_success_rate"))
    p_rec    = P(row.get("rec_success_rate"))
    p_intent = P(row.get("intent_success_rate"))
    p_buy    = P(row.get("buy_success_rate"))
    d1       = P(row.get("bayesian_dropout_pref_to_rec"))
    d2       = P(row.get("bayesian_dropout_rec_to_intent"))
    d3       = P(row.get("bayesian_dropout_intent_to_buy"))

    pref   = n * (p_pref if np.isfinite(p_pref) else 0.0)
    rec    = pref * (1 - d1) if np.isfinite(pref)   and np.isfinite(d1) else n * (p_rec    if np.isfinite(p_rec)    else 0.0)
    intent = rec  * (1 - d2) if np.isfinite(rec)    and np.isfinite(d2) else n * (p_intent if np.isfinite(p_intent) else 0.0)
    buy    = intent*(1 - d3) if np.isfinite(intent) and np.isfinite(d3) else n * (p_buy    if np.isfinite(p_buy)    else 0.0)

    drop0 = max(0.0, float(n) - float(pref))

    rows = [
        {"from_stage":"전체","to_stage":"미선호", "count": drop0}, 
        {"from_stage":"전체","to_stage":"선호", "count": pref},     
        {"from_stage":"선호","to_stage":"추천",     "count":max(0.0, rec)},
        {"from_stage":"추천","to_stage":"구매의향", "count":max(0.0, intent)},
        {"from_stage":"구매의향","to_stage":"구매", "count":max(0.0, buy)},
    ]

    g = pd.DataFrame(rows).dropna()
    g["count"]  = pd.to_numeric(g["count"], errors="coerce").fillna(0)
    g           = g[g["count"] > 0]
    g["dist"]   = g["to_stage"].map(ORDER) - g["from_stage"].map(ORDER)
    g["kind"]   = np.where(g["dist"]==1, "직접", "간접")
    g["to_buy"] = (g["to_stage"]=="구매")
    kphi        = _flow_scale(seg, mod, loy)
    g["flow_phi"] = g["count"].astype(float) * kphi
    g["segment"] = seg; g["model"] = mod; g["loyalty"] = loy
    return g[[
        "from_stage","to_stage","count","dist","kind","to_buy","flow_phi",
        "segment","model","loyalty"
    ]]

LEVELS_FOR_SANKEY = [
    ("전체",               []),
    ("세그먼트",           ["segment"]),
    ("모델",               ["model"]),
    ("충성도",             ["loyalty"]),
    ("세그×모델",          ["segment","model"]),
    ("세그×충성도",        ["segment","loyalty"]),
    ("모델×충성도",        ["model","loyalty"]),
    ("모델×세그×충성도",   ["segment","model","loyalty"]),
]

def build_sankey_cache_from_master(df_master: pd.DataFrame,
                                   collapse_to_buy=True,
                                   collapse_from=("선호","추천","구매의향")) -> pd.DataFrame:
    dfm = _ensure_key_cols(df_master).copy()
    out = []
    for _lvl, keys in LEVELS_FOR_SANKEY:
        if not keys:
            seg, mod, loy = "ALL","ALL","ALL"
            row = compose_composite_row(dfm)
            if not row.empty:
                part = _sankey_from_master_row(row, seg, mod, loy)
                part["level"] = _lvl
                out.append(part)
            continue

        for vals, grp in dfm.groupby(keys, dropna=False):
            if not isinstance(vals, tuple): vals = (vals,)
            seg = vals[keys.index("segment")] if "segment" in keys else "ALL"
            mod = vals[keys.index("model")]   if "model"   in keys else "ALL"
            loy = vals[keys.index("loyalty")] if "loyalty" in keys else "ALL"
            row = compose_composite_row(grp)
            if row.empty: 
                continue
            part = _sankey_from_master_row(row, seg, mod, loy)
            part["level"] = _lvl
            out.append(part)

    if not out:
        return pd.DataFrame(columns=[
            "from_stage","to_stage","count","dist","kind","to_buy","flow_phi",
            "segment","model","loyalty","level"
        ])

    full = pd.concat(out, ignore_index=True)
    if collapse_to_buy and not full.empty:
        full = (full.groupby(["level","segment","model","loyalty"], group_keys=False)
                    .apply(lambda g: add_collapsed_to_buy(g, add_from=collapse_from))
                    .reset_index(drop=True))
    return full

def build_sankey_flow_table(
    df_or_tbl: pd.DataFrame | None,
    seg="ALL", mod="ALL", loy="ALL",
    collapse_to_buy=True,
    collapse_from=("선호","추천","구매의향")
):
    if df_or_tbl is None or df_or_tbl.empty:
        return pd.DataFrame(columns=["from_stage","to_stage","count","dist","kind","to_buy","flow_phi"])

    s = df_or_tbl.copy()
    low = {str(c).strip().lower(): c for c in s.columns}

    looks_table = (("from_stage" in low and "to_stage" in low) and
                   (("count" in low) or ("flow_phi" in low) or ("bayesian_flow_count" in low)))

    if looks_table:
        t = _sanitize_sankey_table(
            s, seg=seg, mod=mod, loy=loy,
            enforce_single_level=True, drop_overall_if_mixed=True
        )
        if collapse_to_buy:
            t = add_collapsed_to_buy(t, add_from=collapse_from)
        return t

    return _sankey_build_table(
        s, seg=seg, mod=mod, loy=loy,
        collapse_to_buy=collapse_to_buy, collapse_from=collapse_from
    )


def sankey_figure(
    df_sankey: pd.DataFrame | None,
    seg, mod, loy,
    normalize=False, base_stage="전체",
    drag=False, show_kind=True,
    table_override: pd.DataFrame | None = None,
):
    # ── 0) 레거시/실수 호환: normalize 자리에 DataFrame이 들어온 경우 보정
    #    (스모크 테스트에서 positional로 override가 들어오는 패턴 방지)
    if isinstance(normalize, pd.DataFrame) and table_override is None:
        table_override = normalize
        normalize = False  # 의미 없는 값이었으므로 안전 기본값

    # ── 1) 테이블 소스 선택
    if table_override is not None:
        # override가 raw여도 안전하게 정규화/보강
        g = _sanitize_sankey_table(table_override, seg=seg, mod=mod, loy=loy)
    else:
        g = build_sankey_flow_table(df_sankey, seg=seg, mod=mod, loy=loy, collapse_to_buy=True)

    if g is None or g.empty:
        return _empty_fig("No Sankey data")

    # ── 2) 색/인덱스 준비
    idx = {v:i for i,v in enumerate(STAGES)}

    STAGE_COLOR = {
        "전체":   COL_STAGE_OVERALL,
        "미선호": COL_STAGE_NONPREF,
        "선호":   COL_STAGE_PREF,
        "추천":   COL_STAGE_REC,
        "구매의향": COL_STAGE_INTENT,
        "구매":   COL_STAGE_BUY,
    }

    # ★ 여기 한 줄: Sankey에서 '전체'만 검정으로
    STAGE_COLOR["전체"] = "#000000"         # 또는 COL_BLACK
    node_colors = [STAGE_COLOR[s] for s in STAGES]

    # ✅ 노드 x 좌표도 6개로
    xs = [0.00, 0.18, 0.34, 0.54, 0.74, 0.94]

    # ── 3) 그림
    fig = go.Figure()
    fig.add_trace(go.Sankey(
        arrangement=("freeform" if drag else "fixed"),
        valueformat=",.1f", valuesuffix=" φ",
    node=dict(
        pad=14, thickness=18, label=STAGES,
        x=xs, y=[0.50]*len(STAGES),
        color=node_colors, line=dict(color="#9aa0a6", width=0.7),
    ),
        link=dict(
            source=[idx[a] for a in g["from_stage"]],
            target=[idx[b] for b in g["to_stage"]],
            value=g["flow_phi"].astype(float).tolist(),
            color=(
                np.where(g["kind"].astype(str)=="직접",
                         hex_to_rgba(COL_LINK_DIRECT,   0.90),
                         hex_to_rgba(COL_LINK_INDIRECT, 0.70))
                if show_kind else [hex_to_rgba(COL_LINK_DIRECT, 0.85)] * len(g)
            ).tolist(),
            customdata=np.stack([
                g["kind"].astype(str).to_numpy(),
                g["dist"].astype(int).to_numpy(),
                g["count"].astype(float).to_numpy(),
            ], axis=-1),
            hovertemplate=(
                "%{customdata[0]} | %{source.label} → %{target.label}"
                "<br>점프: %{customdata[1]}단계"
                "<br>실제유량: %{customdata[2]:,} (표시 %{value:,.1f} φ)"
                "<extra></extra>"
            ),
        ),
    ))

    if show_kind:
        fig.add_trace(go.Scatter(x=[None], y=[None], mode="markers",
            marker=dict(size=10, color=hex_to_rgba(COL_LINK_DIRECT, 0.90)),   name="직접(인접)"))
        fig.add_trace(go.Scatter(x=[None], y=[None], mode="markers",
            marker=dict(size=10, color=hex_to_rgba(COL_LINK_INDIRECT, 0.70)), name="간접(스킵)"))

    base = base_stage if base_stage in STAGES else "전체"
    tot_dir = float(g.loc[g["kind"]=="직접", "flow_phi"].sum())
    tot_ind = float(g.loc[g["kind"]=="간접", "flow_phi"].sum())
    # sankey_figure 끝부분
    fig.update_layout(
        title=f"Journey Sankey · 모든 순방향(스킵 포함) · 기준={base}",
        height=390, showlegend=True,
        paper_bgcolor="#fff", plot_bgcolor="#fff",
        font=dict(color="#111"),
        margin=dict(l=10, r=10, t=32, b=64),
    )
    fig.add_annotation(
        x=0, y=-0.20, xref="paper", yref="paper",
        showarrow=False, align="left",
        text=f"직접 {tot_dir:,.1f} φ · 간접 {tot_ind:,.1f} φ",
        font=dict(size=11, color="#444")
    )

    # ↓↓↓ 이 네 줄은 반드시 함수 안쪽(같은 들여쓰기 레벨)이어야 함
    fig = apply_dense_grid(fig)  # 공통 스타일

    # Sankey 전용: 축 감추기(카테시안 축 없음)
    fig.update_xaxes(visible=False, showgrid=False, zeroline=False, fixedrange=True)
    fig.update_yaxes(visible=False, showgrid=False, zeroline=False, fixedrange=True)

    return fig


# ==== STAGE COLORS (전체→선호→추천→의향→구매) ====
COL_STAGE_OVERALL = "#C32C2C"  # 빨
COL_STAGE_PREF    = "#D24D3E"  # 주
COL_STAGE_REC     = "#DE937A"  # 노
COL_STAGE_INTENT  = "#D49442"  # 베(골드톤)
COL_STAGE_BUY     = "#2B8E81"  # 초fig.update_xaxes
COL_STAGE_NONPREF = "#9CA3AF"  # 미선호(회색)


def matrix_funnel_figure(row, df_tm, seg, mod, loy, **kwargs):
    # --- 0) 로컬 확률 정규화 + 클리핑 유틸 ---
    def _p(x):
        x = _safe_num(x)
        if not np.isfinite(x):
            return np.nan
        # 1.5보다 크면 '퍼센트(예: 23=23%)'로 보고 100으로 나눔
        return x / 100.0 if x > 1.5 else x

    def _clip01(x):
        return np.nan if not np.isfinite(x) else float(min(1.0, max(0.0, x)))

    # 1) 드롭/최종율 확보 (확률 정규화 + [0,1] 클리핑)
    d1_raw, d2_raw, d3_raw, full_raw = drops_from_anywhere(row, df_tm, seg, mod, loy)
    d1, d2, d3 = map(_clip01, map(_p, (d1_raw, d2_raw, d3_raw)))
    full_conv  = _p(full_raw)  # 최종율은 음수/1초과가 들어올 수 있어도 아래에서 단계 보정으로 처리

    # 2) 개별 스테이지 성공률 (확률 정규화)
    pref_sr   = _p(row.get("pref_success_rate"))
    rec_sr    = _p(row.get("rec_success_rate"))
    intent_sr = _p(row.get("intent_success_rate"))
    buy_sr    = _p(row.get("buy_success_rate"))

    # 3) 누적율 계산 (드롭 우선, 결측 시 폴백)
    overall = 1.0
    pref   = pref_sr

    rec = pref * (1 - d1) if np.isfinite(pref) and np.isfinite(d1) else rec_sr
    intent = rec * (1 - d2) if np.isfinite(rec) and np.isfinite(d2) else intent_sr

    if np.isfinite(intent) and np.isfinite(d3):
        buy = intent * (1 - d3)
    elif np.isfinite(buy_sr):
        buy = buy_sr
    elif np.isfinite(full_conv):
        buy = full_conv
    else:
        buy = intent

    # 3-1) 단계 단조감소 보장 + 0~1 클리핑 (여기서 1.6 같은 값 차단)
    seq = [overall,
           _clip01(pref),
           _clip01(rec),
           _clip01(intent),
           _clip01(buy)]
    for i in range(1, len(seq)):
        if np.isfinite(seq[i]) and np.isfinite(seq[i-1]):
            if seq[i] > seq[i-1]:
                seq[i] = seq[i-1]
    overall, pref, rec, intent, buy = seq

    # 4) 라벨/값 구성 (이하는 기존 그대로)
    labels, values = ["전체"], [overall]
    if np.isfinite(pref):   labels.append("선호");     values.append(pref)
    if np.isfinite(rec):    labels.append("추천");     values.append(rec)
    if np.isfinite(intent): labels.append("구매의향"); values.append(intent)
    if np.isfinite(buy):    labels.append("구매");     values.append(buy)

    if len(labels) <= 1:
        return _empty_fig("No Funnel data")

    txtpos = ["inside" if v >= 0.07 else "outside" for v in values]

    color_map = {
        "전체":   hex_to_rgba(COL_STAGE_OVERALL, 0.85),
        "선호":   hex_to_rgba(COL_STAGE_PREF,    0.85),
        "추천":   hex_to_rgba(COL_STAGE_REC,     0.85),
        "구매의향": hex_to_rgba(COL_STAGE_INTENT,  0.85),
        "구매":   hex_to_rgba(COL_STAGE_BUY,     0.85),
    }
    colors = [color_map.get(l, hex_to_rgba(COL_GRAY, 0.85)) for l in labels]

    fig = go.Figure(go.Funnel(
        y=labels,
        x=values,
        name="누적율",
        customdata=values,
        textinfo="none",
        texttemplate="%{customdata:.1%}",
        textposition=txtpos,
        hovertemplate="%{label}: %{customdata:.1%}<extra></extra>",
        marker=dict(color=colors, line=dict(width=0.6, color="rgba(0,0,0,0.25)")),
        connector=dict(line=dict(color="rgba(0,0,0,0.25)", width=0.6)),
    ))

    fig.update_layout(
        title="Funnel (누적율)",
        height=360,
        margin=dict(l=10, r=10, t=30, b=10),
        paper_bgcolor="#ffffff",
        plot_bgcolor="#ffffff",
    )
    fig.update_xaxes(dtick=_auto_dtick(1.0), tickformat=".0%")
    return apply_dense_grid(fig, x_prob=True)

def survival_curve_figure(row, df_tm, seg, mod, loy):
    d1, d2, d3, _ = drops_from_anywhere(row, df_tm, seg, mod, loy)
    vals = [1.0]
    if np.isfinite(d1): vals.append(vals[-1]*(1-d1))
    if np.isfinite(d2): vals.append(vals[-1]*(1-d2))
    if np.isfinite(d3): vals.append(vals[-1]*(1-d3))
    if len(vals) == 1: return _empty_fig("No Survival data")
    stages = ["Start","선호","추천","구매의향","구매"][:len(vals)]
    xs = list(range(len(vals)))
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=xs, y=vals, mode="lines+markers",
        line=dict(width=3, color=COL_GRAY), marker=dict(color=COL_GREEN_LITE),
        hovertemplate="단계=%{text}<br>생존=%{y:.1%}<extra></extra>", text=stages, name="생존확률"
    ))
    drops = [d1,d2,d3]
    for i, dv in enumerate(drops, start=1):
        if i < len(vals) and np.isfinite(dv):
            fig.add_annotation(x=i-0.5, y=(vals[i-1]+vals[i])/2,
                               text=f"실패 {dv:.1%}", showarrow=False,
                               font=dict(size=11, color=COL_ORANGE))
    fig.update_layout(height=320, title="스테이지 생존 커브",
                      xaxis=dict(tickmode="array", tickvals=xs, ticktext=stages),
                      yaxis=dict(range=[0,1], tickformat=".1%"))
    return apply_dense_grid(fig, y_prob=True)

def waterfall_figure(row, df_tm, seg, mod, loy):
    d1, d2, d3, full = drops_from_anywhere(row, df_tm, seg, mod, loy)

    def _as_prob(p):
        p = _safe_num(p)
        if not np.isfinite(p): return np.nan
        return p/100.0 if p > 1.5 else p

    d1, d2, d3 = map(_as_prob, [d1, d2, d3])
    buy_sr  = _as_prob(row.get("buy_success_rate"))
    intent  = _as_prob(row.get("intent_success_rate"))
    full_in = _as_prob(full)

    # 최종 구매율 보정
    full = full_in
    if not np.isfinite(full):
        if np.isfinite(buy_sr): full = buy_sr
        elif np.isfinite(intent) and np.isfinite(d3): full = intent * (1.0 - d3)
        elif all(np.isfinite([d1, d2, d3])): full = (1.0 - d1) * (1.0 - d2) * (1.0 - d3)

    # 절대 드롭
    if all(np.isfinite([d1, d2, d3])):
        drop1 = 1.0 * d1
        drop2 = (1.0 - d1) * d2
        drop3 = (1.0 - d1) * (1.0 - d2) * d3
    else:
        drop1 = d1 if np.isfinite(d1) else 0.0
        drop2 = d2 if np.isfinite(d2) else 0.0
        drop3 = d3 if np.isfinite(d3) else 0.0

    # 최종율 미지정이면 드롭 합으로 보정
    final_rate = float(full) if np.isfinite(full) else max(0.0, 1.0 - drop1 - drop2 - drop3)
    if not any(np.isfinite(v) for v in [drop1, drop2, drop3]) and not np.isfinite(final_rate):
        return _empty_fig("No Waterfall data")

    def _fmt_drop(v): 
        return "" if not np.isfinite(v) else (f"-{v:.1%}" if v >= 1e-6 else "-0.0%")

    # ★ 여기부터: '전체 100%' 막대 제거 버전
    measures  = ["relative", "relative", "relative", "total"]
    x         = ["선호→추천<br>Drop", "추천→구매의향<br>Drop", "구매의향→구매<br>Drop", "구매율"]
    y         = [-drop1, -drop2, -drop3, final_rate]
    texts     = [_fmt_drop(drop1), _fmt_drop(drop2), _fmt_drop(drop3), f"{final_rate:.1%}"]
    positions = ["inside", "inside", "inside", "outside"]

    fig = go.Figure(go.Waterfall(
        measure=measures, x=x, y=y,
        name="drop-off",
        text=texts, textposition=positions,
        insidetextfont=dict(color="white"),
        outsidetextfont=dict(color="#111"),
        decreasing={"marker":{"color": COL_GRAY_MED}},
        increasing={"marker":{"color": COL_GRAY_MED}},
        totals={"marker":{"color": COL_BLUE_DEEP}},
        connector={"line":{"color":"rgba(0,0,0,0.25)", "width":0.6}},
        cliponaxis=False, constraintext="both"
    ))

    fig.update_layout(
        height=320,
        title="드롭오프 워터폴",
        yaxis_tickformat=".1%",
        xaxis=dict(tickangle=0, automargin=True),
        margin=dict(l=8, r=8, t=30, b=14),  # 좌우 여백 살짝 더 줄임
        uniformtext_minsize=9, uniformtext_mode="hide",
    )

    # 공통 스타일 먼저
    fig = apply_dense_grid(fig, y_prob=True)

    # ── 워터폴 가독성 튜닝(Apply 후 다시 덮어쓰기)
    fig.update_layout(
        showlegend=False,   # 범례 숨겨 상단 공간 확보
        bargap=0.15,        # 바 사이 간격 축소 → 막대가 두툼하게
        margin=dict(l=8, r=8, t=30, b=14),
    )
    fig.update_xaxes(automargin=True)

    return fig


def stacked_funnel_figure(row):
    stages = [("선호", "pref_success_rate"), ("추천", "rec_success_rate"),
              ("구매의향", "intent_success_rate"), ("구매", "buy_success_rate")]
    succ = []; fail = []; labs=[]
    for lab, col in stages:
        p = _safe_num(row.get(col))
        if np.isfinite(p):
            succ.append(p); fail.append(1-p); labs.append(lab)
    if not succ: return _empty_fig("No Funnel data")
    fig = go.Figure()
    fig.add_bar(x=labs, y=succ, name="성공", text=[f"{v:.1%}" for v in succ], textposition="inside",
                marker_color=COL_GREEN_LITE)
    fig.add_bar(x=labs, y=fail, name="실패", text=[f"{v:.1%}" for v in fail], textposition="inside",
                marker_color=COL_RED)
    fig.update_layout(barmode="stack", yaxis=dict(range=[0,1], tickformat=".1%"),
                      height=320, title="100% 스택 퍼널 (성공/실패)")
    return apply_dense_grid(fig, y_prob=True)

def forest_figure(df_scope: pd.DataFrame):
    if df_scope is None or df_scope.empty:
        return _empty_fig("No Forest data")

    if not {"model", "segment"}.issubset(set(df_scope.columns)):
        return _empty_fig("Need 'model' and 'segment'")

    s = df_scope.copy()

    # ----- 1) 사용할 단계(성공률) 선택: buy → intent → rec → pref → success_rate → rate
    stage_order = [
        ("buy",    "buy_success_rate"),
        ("intent", "intent_success_rate"),
        ("rec",    "rec_success_rate"),
        ("pref",   "pref_success_rate"),
        ("",       "success_rate"),
        ("",       "rate"),
    ]
    stage = ""
    rate_col = None
    for st, col in stage_order:
        if col in s.columns:
            stage, rate_col = st, col
            break
    if rate_col is None:
        return _empty_fig("No rate column")

    # ----- 2) 표본(n) 컬럼 찾기(단계별 우선, 없으면 일반 표본명으로 폴백)
    def _find_n_col(stage_name: str) -> str | None:
        cands = []
        if stage_name:
            cands += [f"{stage_name}_sample_size", f"{stage_name}_n", f"{stage_name}_total"]
        cands += ["sample_size", "n", "N", "total", "count", "nobs", "베이스수", "표본수", "pref_sample_size"]
        for c in cands:
            if c in s.columns:
                return c
        return None

    n_col = _find_n_col(stage)
    if n_col is None:
        return _empty_fig("No sample size column")

    # ----- 3) 숫자화 + 비율 정규화
    s[rate_col] = pd.to_numeric(s[rate_col], errors="coerce")
    s[n_col]    = pd.to_numeric(s[n_col],    errors="coerce")
    s = s.dropna(subset=[rate_col, n_col])
    if s.empty:
        return _empty_fig("No Forest values")

    r = np.where(s[rate_col] > 1.5, s[rate_col] / 100.0, s[rate_col])       # % → 비율
    r = np.clip(r, 0.0, 1.0)
    n = np.clip(s[n_col].to_numpy().astype(float), 0.0, np.inf)
    k = np.clip(np.round(r * n), 0.0, n)                                     # 성공 수 추정

    # ----- 4) 모델 단위로 집계(중복 y축 제거)
    agg = (pd.DataFrame({
                "model":   s["model"].astype(str),
                "segment": s["segment"].astype(str),
                "k": k, "n": n
           })
           .groupby("model", as_index=False)
           .agg(k=("k","sum"), n=("n","sum"), seg=("segment", lambda x: x.iloc[0])))

    if agg.empty or not np.isfinite(agg["n"]).any():
        return _empty_fig("No Forest values")

    # ----- 5) Jeffreys 95% CI
    alpha = 0.05
    try:
        from scipy.stats import beta as _beta
        agg["p"]  = (agg["k"] + 0.5) / (agg["n"] + 1.0)
        agg["lo"] = _beta.ppf(alpha/2,     agg["k"] + 0.5, agg["n"] - agg["k"] + 0.5)
        agg["hi"] = _beta.ppf(1 - alpha/2, agg["k"] + 0.5, agg["n"] - agg["k"] + 0.5)
    except Exception:
        try:
            from statsmodels.stats.proportion import proportion_confint
            agg["p"]  = (agg["k"] + 0.5) / (agg["n"] + 1.0)
            lo, hi = proportion_confint(agg["k"], agg["n"], alpha=alpha, method="beta")
            agg["lo"], agg["hi"] = lo, hi
        except Exception:
            # Wilson 폴백
            z = 1.959963984540054
            p = agg["k"] / agg["n"]
            denom  = 1 + z*z/agg["n"]
            center = (p + z*z/(2*agg["n"])) / denom
            half   = z*np.sqrt((p*(1-p) + z*z/(4*agg["n"])) / agg["n"]) / denom
            agg["p"]  = p
            agg["lo"] = np.maximum(0.0, center - half)
            agg["hi"] = np.minimum(1.0, center + half)

    use = agg.sort_values("p").reset_index(drop=True)

    # ----- 6) 색(모델의 우세 세그먼트) 지정
    dom_seg = _model_dominant_segment(df_scope)
    mapped_seg = use["model"].map(dom_seg).fillna(use["seg"])
    colors = mapped_seg.apply(_tier_color_for_segment).tolist()

    err_plus  = (use["hi"] - use["p"]).to_numpy()
    err_minus = (use["p"]  - use["lo"]).to_numpy()

    # ----- 7) 플롯
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=use["p"].astype(float),
        y=use["model"].astype(str),
        mode="markers",
        name="모델",   # ← trace 이름 지정 (trace 0 제거)
        hovertemplate="%{y}: %{x:.1%}<extra></extra>",
        marker=dict(size=10, color=colors, line=dict(color=COL_BLACK, width=1.6)),
    ))
    fig.update_traces(error_x=dict(
        type="data", symmetric=False,
        array=err_plus, arrayminus=err_minus,
        color=COL_BLACK, thickness=1.2, width=3
    ))
    add_vline_safe(fig, 0.5, line_dash="dot", line_color=COL_BLACK, opacity=0.4)
    fig.update_layout(
        height=320,
        title="포레스트 플롯 (모델 비교) — 95% CI",
        xaxis=dict(range=[0, 1], dtick=0.1, tickformat=".0%", title="성공률"),
        margin=dict(l=10, r=10, t=30, b=10),
        showlegend=False,   # ← 단일 트레이스면 범례 숨김 (원하면 True로)
    )
    return apply_dense_grid(fig, x_prob=True)


def compare_distribution_figure(df_master, seg, mod, loy, stage_label):
    if df_master is None or df_master.empty:
        return _empty_fig("No Ranking data")

    seg = _as_all(seg); mod = _as_all(mod); loy = _as_all(loy)

    stage2lift = {
        "선호": "pref_lift_vs_galaxy",
        "추천": "rec_lift_vs_galaxy",
        "구매의향": "intent_lift_vs_galaxy",
        "구매": "buy_lift_vs_galaxy",
    }
    lift_col = stage2lift.get(stage_label, "buy_lift_vs_galaxy")
    if lift_col not in df_master.columns:
        return _empty_fig("No lift column")

    # 1) 비교 축 고르기
    candidates = []
    if mod == "ALL": candidates.append("model")
    if seg == "ALL": candidates.append("segment")
    if loy == "ALL": candidates.append("loyalty")

    key = None
    for k in candidates:
        if k in df_master.columns and df_master[k].astype(str).nunique(dropna=True) > 1:
            key = k
            break
    if key is None:
        # fallback: 유니크 가장 많은 축
        avail = [c for c in ["model","segment","loyalty"] if c in df_master.columns]
        if not avail:
            return _empty_fig("No grouping key")
        key = max(avail, key=lambda c: df_master[c].astype(str).nunique(dropna=True))

    # 2) 전체/선택 집계
    overall = (df_master.groupby(key, as_index=False)
                        .agg({lift_col: "mean"})
                        .rename(columns={lift_col: "전체"}))

    scope = df_master.copy()
    if seg != "ALL": scope = scope[scope["segment"].astype(str) == seg]
    if mod != "ALL": scope = scope[scope["model"].astype(str)   == mod]
    if loy != "ALL": scope = scope[scope["loyalty"].astype(str) == loy]

    if scope.empty:
        return _empty_fig("No values")

    selected = (scope.groupby(key, as_index=False)
                      .agg({lift_col: "mean"})
                      .rename(columns={lift_col: "선택"}))

    merged = pd.merge(overall, selected, on=key, how="outer")
    if merged.empty:
        return _empty_fig("No values")

    # 3) 정리: 키는 문자열로, 결측 수치만 0.0으로
    merged[key] = merged[key].astype(str)
    for col in ["전체", "선택"]:
        if col in merged.columns:
            merged[col] = pd.to_numeric(merged[col], errors="coerce")
    merged[["전체","선택"]] = merged[["전체","선택"]].fillna(0.0)

    # 정렬 순서(선택 오름차순이 기본, 전부 0이면 전체 기준)
    if (merged["선택"] != 0).any():
        order = merged.sort_values("선택", ascending=True)[key].tolist()
    else:
        order = merged.sort_values("전체", ascending=True)[key].tolist()

    base = merged.set_index(key).loc[order]

    # 4) 색상
    vals_sel = base["선택"].to_numpy()
    if key == "model":
        dom_seg = _model_dominant_segment(df_master)
        bar_colors = [_tier_color_for_segment(dom_seg.get(k, "LowEnd")) for k in order]
    else:
        bar_colors = royg_color_for(vals_sel)

    # 5) 그림
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=base["전체"], y=order, orientation="h", name="전체",
        marker_color="rgba(150,150,150,0.35)"
    ))
    fig.add_trace(go.Bar(
        x=vals_sel, y=order, orientation="h", name="선택",
        marker=dict(color=bar_colors, line=dict(color=COL_GRAY, width=0.5)),
        text=[f"{v:+.1f}" for v in vals_sel], textposition="outside"
    ))
    add_vline_safe(fig, 0, line_dash="dot", line_color=COL_GRAY)
    fig.update_layout(
        barmode="group",
        title=f"{stage_label} 리프트 — 전체 vs 선택 ({key})",
        xaxis_title="Lift vs Galaxy",
        height=320,
        margin=dict(l=10, r=10, t=30, b=10),
        paper_bgcolor="#ffffff", plot_bgcolor="#ffffff"
    )
    return apply_dense_grid(fig)

def bubble_figure(
    df_scope: pd.DataFrame,
    lift_col: str,
    snr_col: str,
    label_top_n: int = 4,
    label_inside: bool = False,
    textfont_size: int = 11
) -> go.Figure:
    # --- 가드 ---
    if df_scope is None or df_scope.empty:
        return _empty_fig("No Bubble data")
    if lift_col not in df_scope.columns or snr_col not in df_scope.columns:
        return _empty_fig("No Bubble data")

    s = df_scope.copy()
    s[lift_col] = pd.to_numeric(s[lift_col], errors="coerce")
    s[snr_col]  = pd.to_numeric(s[snr_col],  errors="coerce")
    s["pref_sample_size"] = pd.to_numeric(
        s.get("pref_sample_size", pd.Series(1, index=s.index)),
        errors="coerce"
    ).fillna(1.0)

    key = "model" if ("model" in s.columns and s["model"].notna().any()) else (
          "segment" if "segment" in s.columns else None)
    if key is None:
        return _empty_fig("No Bubble key")

    need_cols = [key, lift_col, snr_col, "pref_sample_size"]
    # 있으면 segment를 추가(색상용)
    if "segment" in s.columns and "segment" not in need_cols:
        need_cols.append("segment")

    use = s[need_cols].dropna(subset=[lift_col, snr_col])
    if use.empty:
        return _empty_fig("No Bubble values")

    # ---- 집계 (segment 유무에 따라 분기) ----
    if "segment" in use.columns:
        grp = (use.groupby(key, as_index=False)
                 .agg(x=(lift_col, "mean"),
                      y=(snr_col,  "mean"),
                      n=("pref_sample_size", "sum"),
                      seg=("segment", "first")))
    else:
        grp = (use.groupby(key, as_index=False)
                 .agg(x=(lift_col, "mean"),
                      y=(snr_col,  "mean"),
                      n=("pref_sample_size", "sum")))
        # 색상 함수가 참조하는 seg 컬럼 보강(없으면 NaN)
        grp["seg"] = np.nan

    # ---- 색상 ----
    dom_seg = _model_dominant_segment(df_scope)
    def _color_for(row):
        if key == "model":
            base_seg = dom_seg.get(str(row[key]), row["seg"])
        else:
            base_seg = row["seg"] if pd.notna(row["seg"]) else row[key]
        return _tier_color_for_segment(base_seg)
    grp["color"] = grp.apply(_color_for, axis=1)

    # ---- 버블 크기(√스케일) ----
    n = grp["n"].astype(float).to_numpy()
    if np.isfinite(n).any():
        r = np.sqrt(np.maximum(n, 0))
        r0, r1 = float(np.nanmin(r)), float(np.nanmax(r))
        size = 24.0 if abs(r1 - r0) < 1e-9 else 12 + (r - r0)/(r1 - r0) * 48
    else:
        size = np.full(len(grp), 24.0)

    # ---- 라벨 ----
    labels_all = grp[key].astype(str).tolist()
    if label_top_n is None or label_top_n == -1:
        text = labels_all
    elif label_top_n <= 0:
        text = [""] * len(labels_all)
    else:
        top_idx = np.argsort(-grp["n"].to_numpy())[:label_top_n]
        show = set(top_idx.tolist())
        text = [labels_all[i] if i in show else "" for i in range(len(labels_all))]
    hovertext = grp[key].astype(str)

    # ===== 승/패 분할 경계 & 음영 =====
    x_vals = grp["x"].astype(float).to_numpy()
    y_vals = grp["y"].astype(float).to_numpy()
    x_thr = 0.0 if (np.nanmin(x_vals) < 0 < np.nanmax(x_vals)) else float(np.nanmedian(x_vals))
    y_thr = 2.0 if (np.nanmin(y_vals) <= 2.0 <= np.nanmax(y_vals)) else float(np.nanmedian(y_vals))

    x_min, x_max = float(np.nanmin(x_vals)), float(np.nanmax(x_vals))
    y_min, y_max = float(np.nanmin(y_vals)), float(np.nanmax(y_vals))
    x_pad = (x_max - x_min) * 0.03 if np.isfinite(x_max - x_min) else 0.0
    y_pad = (y_max - y_min) * 0.03 if np.isfinite(y_max - y_min) else 0.0
    x0, x1 = x_min - x_pad, x_max + x_pad
    y0, y1 = y_min - y_pad, y_max + y_pad

    winner_fill = hex_to_rgba("#FDE68A", 0.16)  # 승자(연노랑)
    loser_fill  = hex_to_rgba("#9CA3AF", 0.14)  # 패자(연회색)

    fig = go.Figure()

    # 음영 영역(안전 사각형 헬퍼 사용)
    add_vrect_safe(fig, x0, x_thr, y0=y_thr, y1=y1, fillcolor=loser_fill, layer="below")
    add_vrect_safe(fig, x_thr, x1, y0=y_thr, y1=y1, fillcolor=winner_fill, layer="below")

    # 경계선(안전 버전)
    add_vline_safe(fig, x_thr, line_dash="dot", line_color="#888", opacity=0.6)
    add_hline_safe(fig, y_thr, line_dash="dot", line_color="#888", opacity=0.6)

    # ---- 버블 (✅ _trace → add_trace) ----
    fig.add_trace(go.Scatter(
        x=grp["x"], y=grp["y"],
        mode="markers+text",
        text=text,
        hovertext=hovertext,
        textposition=("middle center" if label_inside else "top center"),
        textfont=dict(size=textfont_size),
        cliponaxis=False,
        marker=dict(size=size, color=grp["color"], line=dict(color="#111", width=0.7)),
        customdata=grp["n"].astype(float),
        hovertemplate=(
            f"{key}=%{{hovertext}}<br>"
            "Lift=%{x:.1f}<br>"
            "SNR=%{y:.1f}<br>"
            "표본=%{customdata:,}<extra></extra>"
        ),
        name="모델/세그"
    ))

    # 레이아웃
    fig.update_layout(
        title="Lift vs SNR (버블=표본수)",
        xaxis_title=None,
        yaxis_title="SNR",
        height=320,
        showlegend=False,
        paper_bgcolor="#fff", plot_bgcolor="#fff",
        margin=dict(l=10, r=10, t=26, b=48)
    )
    fig.update_xaxes(title_standoff=18, automargin=True)
    fig.update_yaxes(title_standoff=8,  automargin=True)

    # 각주
    foot_y = -0.20
    fig.add_annotation(xref="paper", yref="paper", x=0.00, y=foot_y,
        text="<b>■</b>", showarrow=False, font=dict(size=11, color="#FDE68A"))
    fig.add_annotation(xref="paper", yref="paper", x=0.035, y=foot_y,
        text="승자 영역 (Lift↑, SNR↑)", showarrow=False, font=dict(size=10, color="#555"), xanchor="left")
    fig.add_annotation(xref="paper", yref="paper", x=0.32, y=foot_y,
        text="<b>■</b>", showarrow=False, font=dict(size=11, color="#9CA3AF"))
    fig.add_annotation(xref="paper", yref="paper", x=0.355, y=foot_y,
        text="패자 영역 (Lift↓, SNR↑)", showarrow=False, font=dict(size=10, color="#555"), xanchor="left")
    fig.add_annotation(xref="paper", yref="paper", x=0.67, y=foot_y,
        text="○ 원 크기 = 표본수(√스케일)", showarrow=False, font=dict(size=10, color="#666"), xanchor="left")

    # 공통 스타일 후 재고정
    fig = apply_dense_grid(fig)
    fig.update_layout(height=320, margin=dict(l=10, r=10, t=26, b=48), showlegend=False, xaxis_title=None)
    fig.update_xaxes(title_standoff=18, automargin=True)
    fig.update_yaxes(title_standoff=8,  automargin=True)

    # x축 제목을 annotation으로
    fig.add_annotation(
        text="Lift vs Galaxy",
        xref="x domain", yref="paper",
        x=1, y=-0.10,
        xanchor="right", yanchor="top",
        showarrow=False,
        font=dict(size=12)
    )

    return fig

def ppc_purchase_overlay_figure(row: pd.Series, m: int | None = None, draws: int = 6000) -> go.Figure:
    """관측 구매율과 Posterior(베타) & Posterior Predictive(베타-이항) 오버레이."""
    # 관측치
    n = _pick_sample_for_stage(row, "buy")
    if n <= 0:
        n = _safe_int0(row.get("pref_sample_size"))
    p_obs = _safe_num(row.get("buy_success_rate"))
    if not np.isfinite(p_obs):
        return _empty_fig("No PPC data")
    p_obs = float(np.clip(p_obs/100.0 if p_obs > 1.5 else p_obs, 0.0, 1.0))
    k_obs = int(np.clip(round(p_obs * max(n, 1)), 0, max(n, 1)))
    if m is None:
        m = n

    # Posterior (Jeffreys prior: Beta(0.5,0.5))
    a, b = k_obs + 0.5, (n - k_obs) + 0.5
    p = np.random.beta(a, b, size=draws)

    # Posterior predictive (새 표본 m개 관측 시 비율)
    m = max(int(m), 1)
    k_pred = np.random.binomial(m, p)
    rate_pred = k_pred / m

    # 95% HDI
    lo, hi = np.quantile(p, [0.025, 0.975])

    fig = go.Figure()
    fig.add_histogram(
        x=p, nbinsx=60, histnorm="probability density",
        name="Posterior p", marker_color=hex_to_rgba("#9CA3AF", 0.45), opacity=0.55
    )
    fig.add_histogram(
        x=rate_pred, nbinsx=60, histnorm="probability density",
        name=f"PPC n={m:,}", marker_color=hex_to_rgba(COL_STAGE_BUY, 0.55), opacity=0.55
    )

    # 관측치/구간 표시
    add_vline_safe(fig, p_obs, line_color="#111", line_width=2, opacity=0.9)
    fig.add_vrect(x0=lo, x1=hi, fillcolor=hex_to_rgba("#60A5FA", 0.18), line_width=0)

    # ← 핵심: 범례를 아래로(도면 밖) 보내고 아주 작게
    fig.update_layout(
        barmode="overlay",
        title="PPC(구매율) — Posterior & Posterior Predictive",
        height=320,
        margin=dict(l=10, r=10, t=30, b=64),   # 바닥 여백 확보
        showlegend=True,
        legend=dict(
            orientation="h",
            y=-0.22, yanchor="top",   # 플롯 아래쪽, 도면 밖
            x=0.0,   xanchor="left",
            font=dict(size=9),
            itemsizing="constant",
            itemwidth=30
        )
    )
    fig.update_xaxes(range=[0, 1], tickformat=".0%", title="구매율")
    fig.update_yaxes(title="밀도")
    return apply_dense_grid(fig, x_prob=True)

percent1 = FormatTemplate.percentage(1)
num1 = Format(precision=1, scheme=Scheme.fixed)

CARD_STYLE = {
    "background":"white","border":"1px solid #eaeaea","borderRadius":"14px",
    "padding":"14px","boxShadow":"0 2px 10px rgba(0,0,0,0.04)"
}

# (추가) KPI 전용 카드 — 하늘색 배경
KPI_CARD_STYLE = {
    **CARD_STYLE,
    "background": "#EAF2FF",
    "border": "1px solid #d6e4ff"
}

ROW2_CARD_H  = 360
ROW2_GRAPH_H = 320

# ───────────────── app.layout 교체 ─────────────────
app.layout = html.Div(
    [
        dcc.Store(id="store-master"),
        dcc.Store(id="store-tm"),
        dcc.Store(id="store-sankey"),
        dcc.Store(id="store-overall"),
        dcc.Store(id="store-mod-opts"),

        # Sankey 드래그 토글 + 인터랙션 로그
        html.Div(
            [
                dcc.Checklist(
                    id="sankey-drag",
                    options=[{"label": " Sankey 드래그 허용", "value": "drag"}],
                    value=[],
                    inputStyle={"marginRight": "6px"},
                    style={"fontSize": "12px", "color": "#555"},
                ),
                html.Div(id="interact-msg", style={"marginTop": "6px","fontSize": "12px","color": "#444"}),
            ],
            style={"display":"flex","justifyContent":"space-between","alignItems":"center","padding":"0 16px 8px"},
        ),

        # 상단 바
        html.Div(
            [
                html.Div("Bayesian Journey Dashboard", style={"fontWeight":"700","fontSize":"18px"}),
                html.Div(
                    [
                        dcc.Input(id="excel-path", value=DEFAULT_PATH, placeholder="Excel 경로",
                                  style={"width":"520px","marginRight":"8px"}),
                        html.Button("Load", id="load-btn", n_clicks=0, className="btn", style={"marginRight":"8px"}),
                    ],
                    style={"display":"flex","alignItems":"center"},
                ),
            ],
            style={"display":"flex","justifyContent":"space-between","alignItems":"center",
                   "padding":"12px 16px","borderBottom":"1px solid #eee","position":"sticky",
                   "top":"0","background":"#fafafa","zIndex":10},
        ),

        html.Div(id="status-msg", style={"padding":"8px 16px","color":"#555","fontSize":"12px"}),

        # 필터
        html.Div(
            [
                html.Div([html.Label("Segment", style={"fontWeight":"600"}),
                          dcc.Dropdown(id="dd-seg", options=[], value="ALL", clearable=True)],
                         style={"flex":"1","minWidth":"220px","marginRight":"8px"}),
                html.Div([html.Label("Model",   style={"fontWeight":"600"}),
                          dcc.Dropdown(id="dd-mod", options=[], value="ALL", clearable=True)],
                         style={"flex":"1","minWidth":"220px","marginRight":"8px"}),
                html.Div([html.Label("Loyalty", style={"fontWeight":"600"}),
                          dcc.Dropdown(id="dd-loy", options=[], value="ALL", clearable=True)],
                         style={"flex":"1","minWidth":"220px"}),
            ],
            style={"display":"flex","gap":"8px","padding":"12px 16px"},
        ),

        # KPI
        html.Div(
            [
                html.Div([html.Div("표본 수", style={"color":"#888","fontSize":"12px"}),
                          html.H3(id="kpi-sample", style={"margin":"4px 0 0"})], style=KPI_CARD_STYLE),
                html.Div([html.Div("최종 구매율 (Δ 포함)", style={"color":"#888","fontSize":"12px"}),
                          html.H3(id="ins-final", style={"margin":"4px 0 0"})], style=KPI_CARD_STYLE),
                html.Div([html.Div("최대 드롭", style={"color":"#888","fontSize":"12px"}),
                          html.H3(id="ins-drop", style={"margin":"4px 0 0","fontSize":"18px"})], style=KPI_CARD_STYLE),
                html.Div([html.Div("불확실성 (95% HDI 폭)", style={"color":"#888","fontSize":"12px"}),
                          html.H3(id="ins-uncert", style={"margin":"4px 0 0"})], style=KPI_CARD_STYLE),
            ],
            style={"display":"grid","gridTemplateColumns":"repeat(4,1fr)","gap":"12px","padding":"0 16px 12px"},
        ),

        # 숨김 KPI(호환)
        html.Div([html.H3(id="kpi-buy-success"), html.H3(id="kpi-buy-fail")], style={"display":"none"}),

        # Row 1: Sankey + 전이 퍼널 + (워터폴/PPC 탭)
        html.Div(
            [
                html.Div(dcc.Graph(id="fig-sankey", config=GRAPH_CONFIG, style={"height":"400px"}),
                         style={**CARD_STYLE, "height":"440px"}),

                html.Div(dcc.Graph(id="fig-matrix", config=GRAPH_CONFIG, style={"height":"400px"}),
                         style={**CARD_STYLE, "height":"440px"}),

                html.Div(
                    [
                        dcc.Tabs(
                            id="tab-right", value="waterfall",
                            children=[
                                dcc.Tab(label="워터폴", value="waterfall"),
                                dcc.Tab(label="PPC(구매율)", value="ppc"),
                            ],
                            style={"marginBottom":"6px"},
                        ),
                        dcc.Graph(id="fig-right", config=GRAPH_CONFIG, style={"height":"400px"}),
                    ],
                    style={**CARD_STYLE, "height":"440px"},
                ),
            ],
            style={
                "display":"grid",
                "gridTemplateColumns":"1.3fr 1fr 1.5fr",   # ← 오른쪽 카드 넓힘
                "gap":"18px",
                "padding":"30px 30px 30px", "marginBottom":"36px"
            },
        ),

        # Row 2: 스테이지 리프트 + 포레스트 + 버블
        html.Div(
            [
                html.Div(
                    [
                        html.Div(
                            [
                                html.Span("Stage", style={"fontSize":"12px","color":"#666","marginRight":"8px"}),
                                dcc.Dropdown(
                                    id="dd-stage-rank",
                                    options=[{"label": v, "value": v} for v in ["선호","추천","구매의향","구매"]],
                                    value="구매", clearable=False,
                                    style={"width":"140px","fontSize":"12px"},
                                ),
                            ],
                            style={"display":"flex","justifyContent":"flex-end","alignItems":"center","marginBottom":"6px"},
                        ),
                        dcc.Graph(id="fig-stage-rank", config=GRAPH_CONFIG, style={"height":"380px"}),
                    ],
                    style={**CARD_STYLE, "height":"420px","overflow":"hidden"},
                ),

                html.Div(dcc.Graph(id="fig-forest", config=GRAPH_CONFIG, style={"height":"380px"}),
                         style={**CARD_STYLE, "height":"420px","overflow":"hidden"}),

                html.Div(dcc.Graph(id="fig-bubble", config=GRAPH_CONFIG, style={"height":"380px"}),
                         style={**CARD_STYLE, "height":"420px","overflow":"hidden"}),
            ],
            style={"display":"grid","gridTemplateColumns":"1fr 1fr 1fr","gap":"18px","padding":"30px 30px 30px", "marginTop":"36px"},
        ),


        # 숨김 그래프
        html.Div(
            [
                dcc.Graph(id="fig-survival", config=GRAPH_CONFIG),
                dcc.Graph(id="fig-funnel",   config=GRAPH_CONFIG),
            ],
            style={"display":"none"},
        ),

        # 상세 테이블
        html.Div(
            [
                html.H4("상세 메트릭", style={"margin":"0 0 8px 0"}),
                dash_table.DataTable(
                    id="metrics-table",
                    columns=[
                        {"name": "단계",        "id": "단계"},
                        {"name": "베이스수",    "id": "베이스수",    "type": "numeric",
                         "format": Format(precision=0, scheme=Scheme.fixed)},
                        {"name": "성공확률",    "id": "성공확률",    "type": "numeric", "format": percent1},
                        {"name": "실패확률",    "id": "실패확률",    "type": "numeric", "format": percent1},
                        {"name": "하한",        "id": "하한",        "type": "numeric", "format": percent1},
                        {"name": "상한",        "id": "상한",        "type": "numeric", "format": percent1},
                        {"name": "판정",        "id": "판정"},
                        {"name": "평가등급",    "id": "평가등급"},
                        {"name": "SNR",         "id": "SNR",         "type": "numeric", "format": num1},
                        {"name": "Lift",        "id": "Lift",        "type": "numeric", "format": num1},
                        {"name": "raw평균",     "id": "raw평균",     "type": "numeric", "format": percent1},
                        {"name": "raw표준편차", "id": "raw표준편차", "type": "numeric", "format": percent1},
                    ],
                    data=[],
                    page_size=10,
                    style_table={"overflowX":"auto"},
                    style_cell={
                        "fontFamily":"Noto Sans KR, Arial, sans-serif",
                        "fontSize":"12px",
                        "padding":"6px",
                    },
                    style_header={"fontWeight":"bold"},
                    style_data_conditional=[
                        {"if": {"column_id": "베이스수"},     "textAlign": "right"},
                        {"if": {"column_id": "성공확률"},     "textAlign": "right"},
                        {"if": {"column_id": "실패확률"},     "textAlign": "right"},
                        {"if": {"column_id": "하한"},         "textAlign": "right"},
                        {"if": {"column_id": "상한"},         "textAlign": "right"},
                        {"if": {"column_id": "SNR"},          "textAlign": "right"},
                        {"if": {"column_id": "Lift"},         "textAlign": "right"},
                        {"if": {"column_id": "raw평균"},      "textAlign": "right"},
                        {"if": {"column_id": "raw표준편차"},  "textAlign": "right"},
                        {"if": {"row_index": "odd"}, "backgroundColor": "#fafafa"},
                    ],
                ),
            ],
            style={**CARD_STYLE, "margin":"18px 16px 24px"},
        ),
    ],
    style={"background":"#f6f7fb","minHeight":"100vh"},
)
# ───────────────── app.layout 교체 끝 ─────────────────

# ===================== 콜백: Load =====================
@app.callback(
    Output("store-master","data"),
    Output("store-tm","data"),
    Output("store-sankey","data"),
    Output("store-overall","data"),
    Output("dd-seg","options"),
    Output("dd-seg","value"),
    Output("store-mod-opts","data"),
    Output("dd-loy","options"),
    Output("dd-loy","value"),
    Output("status-msg","children"),
    Input("load-btn","n_clicks"),
    State("excel-path","value"),
    prevent_initial_call=True
)
def on_load(n, path):
    try:
        exists = os.path.exists(path)
        size   = (os.path.getsize(path) if exists else 0)

        # 1) 엑셀 로드
        df_master, df_tm, df_sankey, overall, seg_opts, mod_opts_all, loy_opts, dbg = load_excel(path)

        # 2) 마스터로부터 모든 조합 Sankey 캐시 합성
        df_sankey_syn = build_sankey_cache_from_master(df_master, collapse_to_buy=True)

        # 3) 상태 메시지(캐시 행수 포함)
        status = (f"✅ 로드 완료 | path={path} (exists={exists}, size={size:,} bytes) | "
                  f"engine={dbg.get('engine')} | sheets={dbg.get('sheets')} | matched={dbg.get('matched')} | "
                  f"sankey_cache={len(df_sankey_syn):,} rows")

        # 4) 리턴: 세 번째(store-sankey)에 캐시를 넣는다
        return (
            df_master.to_json(date_format="iso", orient="split"),
            df_tm.to_json(date_format="iso", orient="split"),
            df_sankey_syn.to_json(date_format="iso", orient="split"),  # ⬅ 여기!
            json.dumps(overall),
            [{"label":v, "value":v} for v in seg_opts], "ALL",
            json.dumps(mod_opts_all),
            [{"label":v, "value":v} for v in loy_opts], "ALL",
            status
        )
    except Exception as e:
        err = f"❌ LOAD ERROR: {type(e).__name__}: {e}"
        print("LOAD ERROR TRACE:\n", traceback.format_exc())
        return None, None, None, None, [], None, None, [], None, err


# 세그먼트 변경 시 모델 옵션 업데이트
@app.callback(
    Output("dd-mod","options"),
    Output("dd-mod","value"),
    Input("dd-seg","value"),
    State("store-master","data"),
    State("store-mod-opts","data"),
)
def on_seg_change(seg, js_master, js_allmods):
    if not js_master or not js_allmods:
        return [], None
    df_master = pd.read_json(js_master, orient="split")
    seg_val = _as_all(seg)
    if seg_val!="ALL":
        mods = ["ALL"] + sorted([str(v) for v in df_master[df_master["segment"].astype(str)==seg_val]["model"].dropna().astype(str).unique().tolist() if str(v)!="ALL"])
    else:
        mods = json.loads(js_allmods)
    return [{"label":v,"value":v} for v in mods], "ALL"

@app.callback(
    Output("interact-msg","children"),
    Input("fig-sankey","clickData"),
    Input("fig-matrix","relayoutData"),
    Input("fig-right","relayoutData"),  
    Input("fig-stage-rank","selectedData"),
    Input("fig-forest","selectedData"),
    Input("fig-bubble","selectedData"),
    prevent_initial_call=True
)
def on_interact(sankey_click, matrix_relayout, wf_relayout, rank_sel, forest_sel, bubble_sel):
    ctx = dash.callback_context
    if not ctx.triggered:
        return dash.no_update

    tid = ctx.triggered[0]["prop_id"]  # e.g. "fig-bubble.selectedData"
    comp, prop = tid.split(".")
    payload = ctx.triggered[0]["value"]

    if prop == "clickData" and payload:
        pt = (payload.get("points") or [{}])[0]
        label = pt.get("label") or f"{pt.get('sourceLabel','?')}→{pt.get('targetLabel','?')}"
        return f"🖱 {comp}: {label} 클릭"
    if prop == "selectedData" and payload:
        n = len(payload.get("points", []))
        return f"🔎 {comp}: {n}개 선택"
    if prop == "relayoutData" and payload:
        keys = ", ".join(list(payload.keys())[:3])
        return f"🧭 {comp}: 뷰 변경({keys}...)"

    return dash.no_update

# update_all 위쪽(같은 파일)에 추가
def _slice_sankey_cache_by_choice(df, seg, mod, loy):
    if df is None or df.empty:
        return pd.DataFrame()
    sub = df.copy()
    if "segment" in sub.columns and seg != "ALL":
        sub = sub[(sub["segment"].astype(str) == seg) | sub["segment"].isna() | (sub["segment"].astype(str) == "ALL")]
    if "model" in sub.columns and mod != "ALL":
        sub = sub[(sub["model"].astype(str) == mod) | sub["model"].isna() | (sub["model"].astype(str) == "ALL")]
    if "loyalty" in sub.columns and loy != "ALL":
        sub = sub[(sub["loyalty"].astype(str) == loy) | sub["loyalty"].isna() | (sub["loyalty"].astype(str) == "ALL")]
    if "level" in sub.columns:
        for lv in LVL_PRIORITY:
            cand = sub[sub["level"].astype(str) == lv]
            if not cand.empty:
                return cand.copy()
    return sub


    # 레벨 우선순위(가장 세분화된 것부터)로 하나만 남기기
    if "level" in sub.columns:
        for lv in LVL_PRIORITY:
            cand = sub[sub["level"].astype(str) == lv]
            if not cand.empty:
                return cand.copy()
    return sub

def _read_df_store(js):
    if not js:
        return pd.DataFrame()
    # 이미 dict/object로 들어오면 시도
    if isinstance(js, dict):
        if {"columns","data"}.issubset(js.keys()):
            return pd.DataFrame(js["data"], columns=js["columns"])
        try:
            return pd.DataFrame(js)
        except Exception:
            return pd.DataFrame()
    # 문자열이면 우선 split → 실패 시 일반 json 해석
    if isinstance(js, str):
        try:
            return pd.read_json(io.StringIO(js), orient="split")
        except Exception:
            try:
                obj = json.loads(js)
                if isinstance(obj, dict) and {"columns","data"}.issubset(obj.keys()):
                    return pd.DataFrame(obj["data"], columns=obj["columns"])
                elif isinstance(obj, list):
                    return pd.DataFrame(obj)
                elif isinstance(obj, dict):
                    # overall 같은 dict가 오면 DF로 만들지 않고 빈 DF 반환
                    return pd.DataFrame()
            except Exception:
                return pd.DataFrame()
    return pd.DataFrame()

def _read_overall(js_overall):
    if not js_overall:
        return {}
    if isinstance(js_overall, dict):
        return js_overall
    try:
        return json.loads(js_overall)
    except Exception:
        return {}

# ===================== 콜백: 대시보드 계산 =====================
@app.callback(
    Output("kpi-sample","children"),
    Output("kpi-buy-success","children"),
    Output("kpi-buy-fail","children"),
    Output("ins-final","children"),
    Output("ins-drop","children"),
    Output("ins-uncert","children"),
    Output("metrics-table","data"),
    Output("fig-sankey","figure"),
    Output("fig-matrix","figure"),
    #Output("fig-simfan","figure"),
    Output("fig-bubble","figure"),
    Output("fig-stage-rank","figure"),
    Output("fig-survival","figure"),
    Output("fig-right","figure"),  
    # Output("fig-waterfall","figure"),
    Output("fig-funnel","figure"),
    Output("fig-forest","figure"),
    Input("dd-seg","value"),
    Input("dd-mod","value"),
    Input("dd-loy","value"),
    Input("sankey-drag","value"),
    Input("dd-stage-rank","value"),
    Input("tab-right","value"),            # ← 추가
    Input("store-master","data"),
    Input("store-tm","data"),
    Input("store-sankey","data"),
    Input("store-overall","data"),
)

def update_all(seg, mod, loy, drag_val, stage_label, tab_right,
               js_master, js_tm, js_sankey, js_overall=None):
    # 기본값 보정
    seg = _as_all(seg); mod = _as_all(mod); loy = _as_all(loy)
    if not isinstance(stage_label, str) or not stage_label:
        stage_label = "구매"
    empty = _empty_fig("Load data first")

    # 가드: 마스터 없으면 15개 템플릿 리턴
    if not js_master:
        return (
            "–", "–", "–",   # kpi-sample, kpi-buy-success, kpi-buy-fail
            "–", "–", "–",   # ins-final, ins-drop, ins-uncert
            [],              # metrics-table.data
            empty, empty,    # fig-sankey, fig-matrix
            empty, empty,    # fig-bubble, fig-stage-rank
            empty, empty,    # fig-survival, fig-right
            empty,           # fig-funnel
            empty            # fig-forest
        )

    js_sankey, js_overall, _ = _maybe_swap_sankey_overall(js_sankey, js_overall)

    try:
        # 0) sankey/overall 뒤바뀜 자동 교정
        js_sankey, js_overall, _ = _maybe_swap_sankey_overall(js_sankey, js_overall)

        # 1) 스토어 읽기(안전)
        df_master = _read_df_store(js_master)
        df_tm     = _read_df_store(js_tm)
        df_sankey = _read_df_store(js_sankey)
        overall   = _read_overall(js_overall)

        # 2) 선택/스코프
        row_pick = pick_row_for(df_master, seg, mod, loy)
        scope = df_master.copy()
        if seg!="ALL": scope = scope[scope["segment"].astype(str)==seg]
        if mod!="ALL": scope = scope[scope["model"].astype(str)==mod]
        if loy!="ALL": scope = scope[scope["loyalty"].astype(str)==loy]

        # 집계행으로 결측 보강
        row_agg = compose_composite_row(scope)
        rowd = {k: row_pick[k] for k in row_pick.index}

        def _safe_num_or_nan(x):
            try:
                fx = float(x)
                return fx if np.isfinite(fx) else np.nan
            except Exception:
                return np.nan

        def coalesce_into(dst_dict, src_series, cols):
            for c in cols:
                va = _safe_num_or_nan(dst_dict.get(c))
                if np.isnan(va):
                    dst_dict[c] = (src_series.get(c) if isinstance(src_series, pd.Series) else np.nan)

        core_cols = [
            "pref_sample_size",
            "pref_success_rate","pref_ci_lower","pref_ci_upper",
            "rec_success_rate","rec_ci_lower","rec_ci_upper",
            "intent_success_rate","intent_ci_lower","intent_ci_upper",
            "buy_success_rate","buy_ci_lower","buy_ci_upper",
            "bayesian_dropout_pref_to_rec","bayesian_dropout_rec_to_intent","bayesian_dropout_intent_to_buy",
            "bayesian_full_conversion",
            "pref_snr","rec_snr","intent_snr","buy_snr",
            "pref_lift_vs_galaxy","rec_lift_vs_galaxy","intent_lift_vs_galaxy","buy_lift_vs_galaxy",
        ]
        coalesce_into(rowd, row_agg, core_cols)
        row = pd.Series(rowd)

        # 3) KPI/테이블
        tbl = metrics_table_row(row)

        def _face(val, good, soso, reverse=False):
            if not np.isfinite(val): return "❔"
            v = (1 - val) if reverse else val
            return "🟢" if v >= good else ("🟡" if v >= soso else "🔴")

        GOOD_P, SOSO_P = 0.55, 0.45
        GOOD_DROP, SOSO_DROP = 0.20, 0.35
        GOOD_W, SOSO_W = 0.08, 0.12

        sample = _safe_int0(row.get("pref_sample_size"))
        kpi_sample_text = f"📊 {sample:,}"

        buy_p = _safe_num(row.get("buy_success_rate"))
        buy_s = (f"{buy_p:.1%}" if np.isfinite(buy_p) else "N/A")
        buy_f = (f"{(1-buy_p):.1%}" if np.isfinite(buy_p) else "N/A")

        overall_buy = _safe_num(overall.get("buy_mean"))
        delta = (buy_p - overall_buy) if (np.isfinite(buy_p) and np.isfinite(overall_buy)) else np.nan
        face_final = _face(buy_p, GOOD_P, SOSO_P, reverse=False)
        ins_final = (f"{face_final} 성공 {buy_s} / 실패 {buy_f} (vs 전체 {delta:+.1%}p)"
                     if np.isfinite(delta) else f"{face_final} 성공 {buy_s} / 실패 {buy_f}")

        d1, d2, d3, _ = drops_from_anywhere(row, df_tm, seg, mod, loy)
        drops = [v for v in [d1, d2, d3] if np.isfinite(v)]
        dmax = max(drops) if drops else np.nan
        face_drop = _face(dmax, GOOD_DROP, SOSO_DROP, reverse=True)
        ins_drop = f"{face_drop} " + biggest_drop_text_by_sources(row, df_tm, seg, mod, loy)

        def _widest_hdi(r):
            pick = []
            for stage, lo_col, hi_col in [("선호","pref_ci_lower","pref_ci_upper"),
                                          ("추천","rec_ci_lower","rec_ci_upper"),
                                          ("구매의향","intent_ci_lower","intent_ci_upper"),
                                          ("구매","buy_ci_lower","buy_ci_upper")]:
                lo = _safe_num(r.get(lo_col)); hi = _safe_num(r.get(hi_col))
                if np.isfinite(lo) and np.isfinite(hi):
                    pick.append((stage, max(0.0, hi - lo)))
            return max(pick, key=lambda x: x[1]) if pick else (None, np.nan)

        stage_w, width_w = _widest_hdi(row)
        face_unc = _face(width_w, GOOD_W, SOSO_W, reverse=True)
        ins_uncert = "데이터 없음" if stage_w is None else f"{face_unc} {stage_w} 단계 {width_w*100:.1f}%p"

        # 4) Sankey (캐시 정규화 → 보강)
        g_for_sankey = build_sankey_flow_table(df_sankey, seg=seg, mod=mod, loy=loy, collapse_to_buy=True)
        if g_for_sankey is None or g_for_sankey.empty:
            # 완전 비면 현재 row로 즉석 합성
            g_for_sankey = _sankey_from_master_row(row, seg, mod, loy)
            g_for_sankey = add_collapsed_to_buy(g_for_sankey, add_from=("선호","추천","구매의향"))

        fig_sankey = sankey_figure(
            df_sankey=None,
            seg=seg, mod=mod, loy=loy,
            drag=("drag" in (drag_val or [])),
            table_override=g_for_sankey
        )

        # 5) 나머지 그래프
        fig_matrix     = matrix_funnel_figure(row, df_tm, seg, mod, loy)
        lift_col       = "buy_lift_vs_galaxy" if "buy_lift_vs_galaxy" in scope.columns else "pref_lift_vs_galaxy"
        snr_col        = "buy_snr"            if "buy_snr"            in scope.columns else "pref_snr"
        fig_bubble     = bubble_figure(scope, lift_col, snr_col)
        fig_stage_rank = compare_distribution_figure(df_master, seg, mod, loy, stage_label)
        fig_survival   = survival_curve_figure(row, df_tm, seg, mod, loy)
        fig_funnel     = stacked_funnel_figure(row)
        fig_forest     = forest_figure(scope)

        fig_right = (ppc_purchase_overlay_figure(row)
                     if (tab_right or "waterfall") == "ppc"
                     else waterfall_figure(row, df_tm, seg, mod, loy))

        # 6) 최종 15개 리턴(콜백 Output 순서대로)
        return (
            kpi_sample_text, buy_s, buy_f,     # kpi-sample, kpi-buy-success, kpi-buy-fail
            ins_final, ins_drop, ins_uncert,   # 인사이트 3개
            tbl.to_dict("records"),            # metrics-table.data
            fig_sankey, fig_matrix,            # sankey, matrix
            fig_bubble, fig_stage_rank,        # bubble, stage-rank
            fig_survival, fig_right,           # survival, right-panel(waterfall/ppc)
            fig_funnel,                        # funnel
            fig_forest                         # forest
        )

    except Exception:
        print("UPDATE ERROR:\n", traceback.format_exc())
        return (
            "–","–","–","–","–","–",
            [],
            empty, empty, empty, empty, empty, empty, empty, empty
        )

# ===================== 실행 =====================
if __name__ == "__main__":
    base_port = int(os.getenv("PORT", "8059"))
    for i in range(5):
        try:
            app.run_server(host="0.0.0.0", port=base_port + i, debug=False, use_reloader=False)
            break
        except (OSError, SystemExit) as e:
            if "Address already in use" in str(e) or getattr(e, "code", None) == 1:
                continue
            raise

