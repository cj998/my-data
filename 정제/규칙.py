"""정제 규칙 가~사.

각 규칙은 '하는 일'만 하고 '하지 않는 일'은 하지 않는다.
- 규칙 함수는 표(DataFrame, 모든 값이 글자)를 받아 새 표를 돌려준다. 받은 표는 고치지 않는다.
- 규칙끼리 일을 나눠 갖지 않는다. 예) 나(빈칸 표기)는 칸 전체가 정확히 '-' 일 때만 바꾸고,
  ' - ' 처럼 공백이 붙은 것은 가(공백 떼기)가 먼저 돌아야 잡힌다. 그래서 목록 순서대로 실행한다.
"""
from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass
from typing import Callable

import pandas as pd

from .읽기 import is_number_mask, numbers


def _map_unique(df: pd.DataFrame, fn: Callable[[str], str]) -> pd.DataFrame:
    """칸마다 서로 다른 값에만 fn 을 한 번씩 돌려 바꾼다 (같은 값이 많을 때 빠름)."""
    out = df.copy()
    for c in out.columns:
        u = out[c].unique()
        m = {v: fn(v) for v in u}
        out[c] = out[c].map(m).astype(object)
    return out


# ----------------------------------------------------------------- 가. 공백 떼기
_INVIS = r"\s\x00-\x1f\x7f-\x9f­​-‏ - ⁠-⁤﻿"
EDGE_INVIS = re.compile(f"^[{_INVIS}]+|[{_INVIS}]+$")


def trim(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    """값 앞뒤의 공백·탭·줄바꿈·제로폭 공백 등 보이지 않는 글자를 지운다. 가운데는 그대로."""
    return _map_unique(df, lambda v: EDGE_INVIS.sub("", v))


# ----------------------------------------------------------------- 나. 빈칸 표기 맞추기
BLANK_MARKS = {"-", "n/a", "없음", "null"}  # 영문은 대소문자 무시 (N/A, NULL 등)


def blank_marks(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    """칸 전체가 - · N/A · 없음 · null 이면 빈칸("")으로. 0, 미응답 등은 그대로."""
    return _map_unique(df, lambda v: "" if v.lower() in BLANK_MARKS else v)


# ----------------------------------------------------------------- 다. 날짜 모양 맞추기
_DATE_SEP = re.compile(r"([0-9]{4})\s*([-./])\s*([0-9]{1,2})\s*\2\s*([0-9]{1,2})\s*\.?")
_DATE_KO = re.compile(r"([0-9]{4})\s*년\s*([0-9]{1,2})\s*월\s*([0-9]{1,2})\s*일")
_DATE_YLAST = re.compile(r"([0-9]{1,2})\s*/\s*([0-9]{1,2})\s*/\s*([0-9]{4})")
_DATE_8 = re.compile(r"([0-9]{4})([0-9]{2})([0-9]{2})")


def _iso(y: int, mo: int, d: int) -> str | None:
    try:
        return dt.date(y, mo, d).isoformat()
    except ValueError:
        return None


def _year_last(a: int, b: int, y: int) -> str | None:
    """숫자/숫자/4자리 에서 어느 쪽이 달인지 소거법으로 정한다.

    한쪽이 12보다 크면 그쪽은 달이 될 수 없으니 일, 다른 쪽이 달 (12/18/2026 → 18은 일, 12는 달).
    그렇게 정한 년·월·일이 실제로 있는 날짜일 때만 바꾼다.
    둘 다 12 이하이고 서로 다르면(6/7/2026) 달·일을 가를 근거가 없으니 None (= 그대로 둔다).
    """
    if a > 12 and b > 12:
        return None
    if a > 12:
        return _iso(y, b, a)
    if b > 12:
        return _iso(y, a, b)
    if a == b:
        return _iso(y, a, b)
    return None


def to_iso_date(v: str) -> str | None:
    """날짜로 확정할 수 있는 값만 YYYY-MM-DD 로. 아니면 None (= 그대로 둔다).

    바꾸는 것: 2025/03/14, 2025.3.14, 2025. 3. 14., 2025-3-14, 2026년6월19일, 2025년 3월 14일,
             12/18/2026·18/12/2026(4자리=연도, 12보다 큰 쪽=일, 나머지=달),
             20260615(앞 4자리=연도, 뒤 4자리=MMDD)
    두는 것: 6/7/2026(달·일 구분 불가), 2025-02-30·20261345(없는 날짜), 시간이 붙은 값
    """
    for pat in (_DATE_SEP, _DATE_KO, _DATE_8):
        m = pat.fullmatch(v)
        if m:
            g = [x for x in m.groups() if x.isdigit()]
            return _iso(int(g[0]), int(g[1]), int(g[2]))
    m = _DATE_YLAST.fullmatch(v)
    if m:
        return _year_last(int(m[1]), int(m[2]), int(m[3]))
    return None


def dates(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    return _map_unique(df, lambda v: to_iso_date(v) or v)


# ----------------------------------------------------------------- 라. 숫자 다듬기
_NUM_UNIT = re.compile(r"([+-]?)([0-9]{1,3}(?:,[0-9]{3})+|[0-9]+)(\.[0-9]+)?(\s*(?:원|점|건|개|분|%))?")


def clean_number(v: str) -> str | None:
    """천 단위 쉼표와 뒤에 붙은 원·점·건·개·분·% 만 뗀다. 환산하지 않는다 (12% → 12).

    쉼표도 단위도 없는 값, ₩처럼 앞에 붙은 기호, 그 밖의 단위는 None (= 그대로 둔다).
    """
    m = _NUM_UNIT.fullmatch(v)
    if not m or ("," not in m[2] and not m[4]):
        return None
    return m[1] + m[2].replace(",", "") + (m[3] or "")


def tidy_numbers(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    return _map_unique(df, lambda v: clean_number(v) or v)


# ----------------------------------------------------------------- 마. 표기 하나로
def norm_key(v: str) -> str:
    """대소문자·띄어쓰기만 다른 표기를 같은 묶음으로 보는 열쇠."""
    return re.sub(r"\s+", "", v.lower())


def unify_candidates(df: pd.DataFrame) -> dict[str, list[dict]]:
    """칸마다, 대소문자·띄어쓰기만 다른 표기가 2가지 이상 섞인 묶음.

    결과: {칸 이름: [{"key": 열쇠, "variants": [(표기, 개수), ...많은 순]}, ...많은 순]}
    뜻만 같은 다른 말(유선/전화)은 열쇠가 달라 절대 같은 묶음이 되지 않는다.
    """
    out: dict[str, list[dict]] = {}
    for c in df.columns:
        vc = df[c][df[c] != ""].value_counts()
        groups: dict[str, list[tuple[str, int]]] = {}
        for v, n in vc.items():
            k = norm_key(v)
            if k:
                groups.setdefault(k, []).append((v, int(n)))
        found = [{"key": k, "variants": vs} for k, vs in groups.items() if len(vs) > 1]
        if found:
            found.sort(key=lambda g: -sum(n for _, n in g["variants"]))
            out[c] = found
    return out


def unify(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    """사용자가 고른 묶음만 고른 표기로 맞춘다.

    cfg["picks"] = {칸 이름: {열쇠: 남길 표기}}
    남길 표기가 그 묶음의 표기가 아니면(열쇠가 다르면) 무시한다 → 다른 말로 바꾸는 일은 생기지 않는다.
    """
    out = df.copy()
    for col, picks in (cfg.get("picks") or {}).items():
        if col not in out.columns:
            continue
        safe = {k: canon for k, canon in picks.items() if norm_key(canon) == k}
        if not safe:
            continue
        s = out[col]
        u = s.unique()
        m = {v: safe.get(norm_key(v), v) if v != "" else v for v in u}
        out[col] = s.map(m).astype(object)
    return out


# ----------------------------------------------------------------- 바. 겹친 줄 지우기
def dedup(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    """모든 칸이 완전히 같은 줄만 첫 줄을 남기고 지운다. 일부 칸만 같은 줄은 남긴다."""
    return df[~df.duplicated(keep="first")].copy()


# ----------------------------------------------------------------- 사. 튀는 값 표시
NUMERIC_SHARE = 0.8   # 빈칸 아닌 값의 80% 이상이 숫자면 '숫자 칸'
MIN_FOR_IQR = 8       # 숫자가 이보다 적으면 이상치는 판단하지 않음 (음수만 표시)


def iqr_bounds(vals: pd.Series, k: float) -> tuple[float, float] | None:
    """정상 범위 = [Q1 − k×IQR, Q3 + k×IQR]. 값이 적거나 IQR=0 이면 None."""
    if len(vals) < MIN_FOR_IQR:
        return None
    q1, q3 = vals.quantile(0.25), vals.quantile(0.75)
    iqr = q3 - q1
    return (q1 - k * iqr, q3 + k * iqr) if iqr > 0 else None


def flag_outliers(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    """숫자 칸에서 음수·동떨어진 값을 찾아 맨 끝 '_표시' 칸에 적는다. 원래 줄·값은 그대로.

    cfg["k"]: IQR 배수 (기본 3). 실행 뒤 cfg["info"] 에 칸별 결과를 적어 둔다.
    """
    k = float(cfg.get("k", 3))
    notes: dict = {i: [] for i in df.index}
    info = []
    for c in df.columns:
        if str(c).startswith("_표시"):
            continue
        s = df[c]
        filled = s[s != ""]
        mask = is_number_mask(filled)
        if not mask.any() or mask.sum() < len(filled) * NUMERIC_SHARE:
            continue
        vals = numbers(filled)
        b = iqr_bounds(vals, k)
        neg = out = 0
        for idx, x in vals.items():
            tags = []
            if x < 0:
                tags.append("음수")
                neg += 1
            if b and (x < b[0] or x > b[1]):
                tags.append("이상치")
                out += 1
            if tags:
                notes[idx].append(f"{c}: {'·'.join(tags)}({s[idx]})")
        info.append({"숫자 칸": c, "숫자 수": len(vals),
                     "정상 범위": f"{b[0]:,.2f} ~ {b[1]:,.2f}" if b else "판단 안 함",
                     "음수": neg, "이상치": out})
    name, n = "_표시", 2
    while name in df.columns:
        name, n = f"_표시_{n}", n + 1
    out_df = df.copy()
    out_df[name] = pd.Series({i: "; ".join(v) for i, v in notes.items()}, dtype=object).reindex(df.index)
    cfg["info"] = info
    return out_df


# ----------------------------------------------------------------- 목록 · 실행
@dataclass(frozen=True)
class Rule:
    id: str
    label: str
    desc: str
    apply: Callable[[pd.DataFrame, dict], pd.DataFrame]


RULES: list[Rule] = [
    Rule("trim", "가. 공백 떼기",
         "값 앞뒤의 공백·탭·줄바꿈·제로폭 공백 같은 보이지 않는 글자를 지웁니다. 가운데 공백은 그대로 둡니다.", trim),
    Rule("blank", "나. 빈칸 표기 맞추기",
         "칸 전체가 - · N/A · 없음 · null 인 경우만 빈칸으로 바꿉니다. 0, 미응답 같은 값은 그대로 둡니다.", blank_marks),
    Rule("date", "다. 날짜 모양 맞추기",
         "2025/03/14 · 2025.3.14 · 2025년 3월 14일 · 20250314를 2025-03-14로 바꿉니다. "
         "12/18/2026처럼 연도가 뒤에 오면 12보다 큰 쪽을 일, 나머지를 달로 봅니다. "
         "6/7/2026처럼 달·일을 가를 수 없는 값, 시간이 붙은 값, 없는 날짜는 그대로 둡니다.", dates),
    Rule("num", "라. 숫자 다듬기",
         "89,000원 → 89000, 12% → 12처럼 천 단위 쉼표와 뒤에 붙은 원·점·건·개·분·%만 뗍니다. "
         "단위 환산은 하지 않습니다. (₩ 같은 앞 기호, 그 밖의 단위는 그대로)", tidy_numbers),
    Rule("unify", "마. 표기 하나로",
         "대소문자·띄어쓰기만 다른 표기 중 직접 고른 묶음만 하나로 맞춥니다. 뜻만 같은 다른 말(유선/전화 등)은 후보에 나오지 않습니다.", unify),
    Rule("dedup", "바. 겹친 줄 지우기",
         "모든 칸의 값이 완전히 똑같은 줄만 첫 줄 하나를 남기고 지웁니다. 일부 칸만 같은 줄은 남깁니다.", dedup),
    Rule("flag", "사. 튀는 값 표시",
         "숫자 칸(값의 80% 이상이 숫자)에서 음수와 동떨어진 값을 찾아 맨 끝 _표시 칸에 적습니다. 원래 줄과 값은 바꾸지 않습니다.", flag_outliers),
]


def diff(a: pd.DataFrame, b: pd.DataFrame) -> dict:
    """두 표의 차이: 바뀐 칸 수(같은 줄·같은 칸 이름끼리), 없어진/생긴 줄·칸 수."""
    idx = a.index.intersection(b.index)
    cols = [c for c in a.columns if c in b.columns]
    changed = int((a.loc[idx, cols].to_numpy() != b.loc[idx, cols].to_numpy()).sum()) if len(idx) and cols else 0
    return {
        "바뀐 칸": changed,
        "지운 줄": len(a.index.difference(b.index)),
        "생긴 줄": len(b.index.difference(a.index)),
        "지운 칸": len([c for c in a.columns if c not in b.columns]),
        "생긴 칸": len([c for c in b.columns if c not in a.columns]),
    }


def run(df: pd.DataFrame, selected: list[str], cfgs: dict[str, dict] | None = None):
    """고른 규칙을 목록 순서대로 실행. (정제된 표, 규칙별 차이 목록)을 돌려준다."""
    if cfgs is None:
        cfgs = {}
    log = []
    cur = df
    for r in RULES:
        if r.id not in selected:
            continue
        new = r.apply(cur, cfgs.setdefault(r.id, {}))
        log.append({"규칙": r.label, **diff(cur, new)})
        cur = new
    return cur, log
