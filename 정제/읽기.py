"""파일을 읽고 세는 계산.

원칙
- 값은 올린 그대로 '글자'로 읽는다. 읽는 동안 숫자·날짜로 바꾸지 않는다.
  (CSV는 파일 속 글자 그대로. 엑셀은 셀에 저장된 값을 글자로 옮긴다. 아래 _cell_text 참고)
- 파일은 bytes 로만 받는다. 디스크에 쓰지 않는다.
"""
from __future__ import annotations

import codecs
import csv
import datetime as dt
import io
import re

import pandas as pd

# 숫자로 읽히는 글자: 12 / -3.5 / .5 / 1e3  (앞뒤 공백·쉼표·단위가 붙으면 숫자로 보지 않음)
NUMBER_RE = re.compile(r"[+-]?(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][+-]?[0-9]+)?")

csv.field_size_limit(2**31 - 1)


# ----------------------------------------------------------------- 파일 읽기
def is_excel(filename: str) -> bool:
    return filename.lower().endswith(".xlsx")


def sheet_names(data: bytes, filename: str) -> list[str]:
    """엑셀이면 시트 이름 목록, CSV면 빈 목록."""
    if not is_excel(filename):
        return []
    from openpyxl import load_workbook

    wb = load_workbook(io.BytesIO(data), read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def detect_encoding(data: bytes) -> str:
    """CSV 글자 인코딩: UTF-8로 읽히면 UTF-8, 아니면 CP949(한글 윈도우)."""
    if data.startswith(codecs.BOM_UTF8):
        return "utf-8-sig"
    try:
        data.decode("utf-8")
        return "utf-8"
    except UnicodeDecodeError:
        return "cp949"


def read_rows(data: bytes, filename: str, sheet: str | None = None) -> list[list[str]]:
    """파일의 모든 줄을 글자 목록으로 돌려준다. 머리글 줄도 포함된 '원본 그대로'의 줄들.

    맨 끝의 완전히 빈 줄만 잘라 낸다 (엑셀이 서식 때문에 빈 줄을 범위에 넣는 경우가 많음).
    """
    rows = _read_excel(data, sheet) if is_excel(filename) else _read_csv(data)
    while rows and all(v == "" for v in rows[-1]):
        rows.pop()
    return rows


def _sniff_delimiter(text: str) -> str:
    lines = text[:20000].splitlines()[:10] or [""]
    best, best_score = ",", -1.0
    for d in (",", "\t", ";", "|"):
        score = sum(line.count(d) for line in lines) / len(lines)
        if score > best_score:
            best, best_score = d, score
    return best


def _read_csv(data: bytes) -> list[list[str]]:
    enc = detect_encoding(data)
    text = data.decode(enc, errors="replace" if enc == "cp949" else "strict")
    reader = csv.reader(io.StringIO(text, newline=""), delimiter=_sniff_delimiter(text))
    return [list(r) for r in reader]


def _cell_text(v) -> str:
    """엑셀 셀 값 → 글자. 엑셀 파일에 저장된 값을 그대로 글자로 적는다.

    - 빈 셀 → ""
    - 정수로 떨어지는 숫자 89000.0 → "89000", 그 밖의 숫자 → 파이썬 표기 그대로 ("0.125")
    - 참/거짓 → "TRUE" / "FALSE" (엑셀 표기)
    - 엑셀이 날짜로 저장한 셀 → "2026-07-01" (시간이 있으면 "2026-07-01 13:05:00")
      ※ 엑셀 날짜는 파일 속에 '날짜 값'으로 들어 있어 표시 모양(7/1/26 등)은 파일에 없다.
    """
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() and abs(v) < 1e16 else repr(v)
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d") if v.time() == dt.time(0) else v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, dt.time):
        return v.strftime("%H:%M:%S")
    return str(v)


def _read_excel(data: bytes, sheet: str | None) -> list[list[str]]:
    from openpyxl import load_workbook

    # data_only=True: 수식 칸은 수식이 아니라 마지막으로 계산된 값을 읽는다
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    try:
        ws = wb[sheet] if sheet else wb.worksheets[0]
        ws.reset_dimensions()  # 파일에 적힌 범위가 틀린 경우가 있어 실제 셀을 끝까지 읽는다
        return [[_cell_text(v) for v in row] for row in ws.iter_rows(values_only=True)]
    finally:
        wb.close()


# ----------------------------------------------------------------- 표 만들기
def to_table(rows: list[list[str]], header_row: int = 1) -> pd.DataFrame:
    """header_row(1부터 세는 줄 번호)를 머리글로, 그 아래 줄을 데이터로 하는 표.

    - 모든 값은 글자(object)로 둔다. 모자란 칸은 "" 로 채운다.
    - 머리글 위의 줄은 버린다.
    - 행 이름(index) = 원본 파일의 줄 번호(_row). 정제 뒤에도 원본 줄을 찾아갈 수 있다.
    - 칸 이름: 앞뒤 공백을 떼고, 비어 있으면 '열3', 겹치면 '이름_2' 로 만든다.
    """
    if not rows or header_row < 1 or header_row > len(rows):
        return pd.DataFrame(index=pd.RangeIndex(0, name="_row"), dtype=object)
    h = header_row - 1
    width = max(len(r) for r in rows[h:])
    raw_names = list(rows[h]) + [""] * (width - len(rows[h]))
    names, seen = [], {}
    for i, n in enumerate(raw_names, start=1):
        n = str(n).strip() or f"열{i}"
        if n in seen:
            seen[n] += 1
            n = f"{n}_{seen[n]}"
        else:
            seen[n] = 1
        names.append(n)
    body = [list(r) + [""] * (width - len(r)) for r in rows[h + 1:]]
    index = pd.RangeIndex(h + 2, h + 2 + len(body), name="_row")
    if not body:
        return pd.DataFrame(columns=names, index=index, dtype=object)
    return pd.DataFrame(body, columns=names, index=index, dtype=object)


# ----------------------------------------------------------------- 세기
def is_number_mask(s: pd.Series) -> pd.Series:
    """각 값이 '숫자로 읽히는 글자'인지 (True/False)."""
    return s.astype(str).str.fullmatch(NUMBER_RE.pattern).fillna(False).astype(bool)


def numbers(s: pd.Series) -> pd.Series:
    """숫자로 읽히는 값만 숫자로 바꿔서 (세기용, 표 자체는 바꾸지 않음)."""
    return pd.to_numeric(s[is_number_mask(s)], errors="coerce").dropna().astype(float)


def column_summary(df: pd.DataFrame) -> pd.DataFrame:
    """칸마다 한 줄: 칸 이름 · 빈칸 수 · 값의 종류 수 · 숫자로 읽히는 값의 비율 · 예시 값 셋.

    - 빈칸 = 글자가 하나도 없는 칸("") — 공백만 있는 칸은 빈칸이 아니다 (원본 그대로 셈)
    - 값의 종류 수 = 빈칸이 아닌 값 중 서로 다른 값의 수
    - 숫자로 읽히는 값의 비율 = 빈칸이 아닌 값 중 숫자로 읽히는 값의 비율 (0~1)
    - 예시 값 셋 = 위에서부터 서로 다른 값 3개
    """
    out = []
    for name in df.columns:
        s = df[name]
        filled = s[s != ""]
        out.append({
            "칸 이름": name,
            "빈칸 수": int((s == "").sum()),
            "값의 종류 수": int(filled.nunique()),
            "숫자로 읽히는 값의 비율": float(is_number_mask(filled).mean()) if len(filled) else 0.0,
            "예시 값 셋": filled.drop_duplicates().head(3).tolist(),
        })
    return pd.DataFrame(out, columns=["칸 이름", "빈칸 수", "값의 종류 수", "숫자로 읽히는 값의 비율", "예시 값 셋"])


def column_stats(s: pd.Series) -> dict:
    """한 칸의 숫자 요약 (원본↔정제 비교용)."""
    filled = s[s != ""]
    nums = numbers(filled)
    has = len(nums) > 0
    return {
        "줄 수": len(s),
        "값 있는 칸": len(filled),
        "빈칸": int((s == "").sum()),
        "값의 종류 수": int(filled.nunique()),
        "숫자로 읽히는 값": len(nums),
        "숫자 비율": (len(nums) / len(filled)) if len(filled) else 0.0,
        "평균": float(nums.mean()) if has else None,
        "중앙값": float(nums.median()) if has else None,
        "최소": float(nums.min()) if has else None,
        "최대": float(nums.max()) if has else None,
        "합계": float(nums.sum()) if has else None,
        "표준편차": float(nums.std(ddof=0)) if has else None,
        "많이 나온 값": filled.value_counts().head(5),
    }
