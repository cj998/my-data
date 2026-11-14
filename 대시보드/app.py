"""정제 대시보드 — 화면만 그린다. 읽기·세기·정제 계산은 정제/ 폴더의 함수를 불러 쓴다.

실행:  streamlit run 대시보드/app.py   (my-data 폴더에서)

올린 파일은 이 화면(세션)의 메모리에만 있다. 폴더에 저장하지 않는다.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))  # 정제/ 를 불러오기 위해
from 정제 import 규칙, 읽기  # noqa: E402

st.set_page_config(page_title="정제 대시보드", layout="wide")
st.title("정제 대시보드 — 올리고, 정리하고, 숫자를 본다")
ss = st.session_state


def fmt(v) -> str:
    if v is None:
        return "—"
    if isinstance(v, float):
        return f"{v:,.2f}".rstrip("0").rstrip(".") if abs(v) < 1e15 else f"{v:.4g}"
    if isinstance(v, int):
        return f"{v:,}"
    return str(v)


def show(v: str) -> str:
    """값을 «»로 감싸 앞뒤 공백이 보이게."""
    return "(빈칸)" if v == "" else f"«{v}»"


# ================================================================ 파일 올리기
up = st.file_uploader("엑셀(.xlsx) 또는 CSV 파일을 올리세요", type=["xlsx", "csv"])
if up is None:
    for k in [k for k in ss if k.startswith(("raw::", "unify"))]:
        del ss[k]
    ss.pop("file_key", None)
    st.info("파일을 올리면 ① 원본 그대로 → ② 정제 규칙 → ③ 칸 비교 → ④ 내려받기 순서로 보여 줍니다.")
    st.stop()

data = up.getvalue()  # 메모리 속 bytes. 디스크에 쓰지 않는다
file_key = f"{up.name}:{len(data)}:{getattr(up, 'file_id', '')}"
if ss.get("file_key") != file_key:  # 새 파일이면 이전 파일의 기억을 지운다
    for k in [k for k in ss if k.startswith(("raw::", "unify", "hdr", "sheet", "cmp", "rule_", "flag_k"))]:
        del ss[k]
    ss["file_key"] = file_key

try:
    sheets = 읽기.sheet_names(data, up.name)
except Exception as e:  # noqa: BLE001
    st.error(f"엑셀 파일을 열 수 없습니다: {e}")
    st.stop()

c1, c2, c3 = st.columns([2, 1, 3])
sheet = None
if len(sheets) > 1:
    sheet = c1.selectbox("시트", sheets, key="sheet")
elif sheets:
    sheet = sheets[0]

raw_key = f"raw::{file_key}::{sheet}"
if raw_key not in ss:  # 같은 파일·시트를 화면이 다시 그려질 때마다 다시 읽지 않도록 세션 메모리에만 둔다
    for k in [k for k in ss if k.startswith("raw::")]:
        del ss[k]
    with st.spinner("읽는 중…"):
        ss[raw_key] = 읽기.read_rows(data, up.name, sheet)
rows: list[list[str]] = ss[raw_key]
if not rows:
    st.warning("읽을 수 있는 줄이 없습니다.")
    st.stop()

header_row = c2.number_input("머리글 줄 (몇 번째 줄)", min_value=1, max_value=len(rows), value=1, step=1,
                             key=f"hdr::{sheet}")
if up.name.lower().endswith(".csv"):
    c3.caption(f"CSV 인코딩: {읽기.detect_encoding(data)}")

df = 읽기.to_table(rows, int(header_row))

with st.expander("파일 앞부분 보기 — 머리글이 몇 번째 줄인지 확인용"):
    n_show = min(len(rows), max(15, int(header_row) + 5))
    w = max(len(r) for r in rows[:n_show])
    head = pd.DataFrame([list(r) + [""] * (w - len(r)) for r in rows[:n_show]],
                        index=pd.RangeIndex(1, n_show + 1, name="줄"), dtype=object)
    head.columns = [f"{i + 1}" for i in range(w)]
    st.dataframe(head.style.apply(
        lambda r: ["background-color:#dbe7fb;color:#1b1f24;font-weight:600" if r.name == header_row else ""] * len(r),
        axis=1), width="stretch")

# ================================================================ ① 원본 그대로
st.header("① 원본 그대로")
m1, m2, _ = st.columns([1, 1, 4])
m1.metric("줄 수", f"{len(df):,}", help="머리글 줄 아래의 데이터 줄 수")
m2.metric("칸 수", f"{len(df.columns):,}")

summary = 읽기.column_summary(df)
st.dataframe(
    summary, hide_index=True, width="stretch",
    column_config={
        "숫자로 읽히는 값의 비율": st.column_config.ProgressColumn(
            "숫자로 읽히는 값의 비율", min_value=0.0, max_value=1.0, format="percent",
            help="빈칸이 아닌 값 중 그대로 숫자로 읽히는 값 (89,000원 · ' 64000 ' 은 숫자가 아님)"),
        "예시 값 셋": st.column_config.ListColumn("예시 값 셋"),
        "빈칸 수": st.column_config.NumberColumn(help="글자가 하나도 없는 칸 (공백만 있는 칸은 빈칸 아님)"),
        "값의 종류 수": st.column_config.NumberColumn(help="빈칸이 아닌 서로 다른 값의 수"),
    },
)

# ================================================================ ② 정제 규칙
st.header("② 정제 규칙")
st.caption("체크한 규칙이 위에서 아래 순서로 적용됩니다.")


def unify_ui(cur: pd.DataFrame) -> dict:
    """마. 표기 하나로 — 합칠 묶음과 남길 표기를 사용자가 직접 고른다."""
    store: dict = ss.setdefault("unify_store", {})  # {칸: {열쇠: 남길 표기}} — 다른 칸을 보는 동안에도 기억
    cands = 규칙.unify_candidates(cur)
    if not cands:
        st.caption("대소문자·띄어쓰기만 다른 표기가 없습니다.")
        return {"picks": store}
    col = st.selectbox("칸", list(cands), key="unify_col",
                       format_func=lambda c: f"{c} — 후보 {len(cands[c])}묶음")
    groups = {g["key"]: g for g in cands[col]}
    gk = f"unify_groups::{col}"
    if gk not in ss:
        ss[gk] = [k for k in store.get(col, {}) if k in groups]
    chosen = st.multiselect(
        "합칠 묶음 (고른 것만 합칩니다)", list(groups), key=gk,
        format_func=lambda k: " · ".join(f"{show(v)} {n}개" for v, n in groups[k]["variants"]))
    picks = {}
    for k in chosen:
        variants = [v for v, _ in groups[k]["variants"]]
        ck = f"unify_canon::{col}::{k}"
        if ck not in ss:
            prev = store.get(col, {}).get(k)
            ss[ck] = prev if prev in variants else variants[0]
        picks[k] = st.selectbox(f"{' / '.join(show(v) for v in variants)} → 이렇게 맞춤", variants, key=ck,
                                format_func=show)
    store[col] = picks
    n = sum(len(m) for m in store.values())
    st.caption(f"모든 칸에서 고른 묶음: {n}개")
    return {"picks": {c: m for c, m in store.items() if m}}


cur = df
log = []
for rule in 규칙.RULES:
    with st.container(border=True):
        on = st.checkbox(rule.label, key=f"rule_{rule.id}")
        st.caption(rule.desc)
        if not on:
            continue
        cfg: dict = {}
        if rule.id == "unify":
            cfg = unify_ui(cur)
        elif rule.id == "flag":
            cfg["k"] = st.radio(
                "동떨어진 값 기준", [3.0, 1.5], horizontal=True, key="flag_k",
                format_func=lambda k: "아주 동떨어진 값만 (Q1−3×IQR 미만, Q3+3×IQR 초과)" if k == 3
                else "보통 (Q1−1.5×IQR 미만, Q3+1.5×IQR 초과)",
                help="IQR = 가운데 50% 값의 폭. 숫자가 8개 미만이거나 값이 거의 같으면 이상치는 판단하지 않고 음수만 표시합니다.")
        new = rule.apply(cur, cfg)
        d = 규칙.diff(cur, new)
        log.append({"규칙": rule.label, **d})
        parts = [f"칸 {d['바뀐 칸']:,}개 바뀜"] + [f"{k} {v:,}개" for k, v in d.items() if k != "바뀐 칸" and v]
        st.markdown("**결과:** " + " · ".join(parts))
        if rule.id == "flag":
            if cfg.get("info"):
                st.dataframe(pd.DataFrame(cfg["info"]), hide_index=True)
            else:
                st.caption("숫자 칸이 없습니다. (라. 숫자 다듬기를 함께 켜면 «89,000원» 같은 칸도 숫자 칸으로 잡힙니다)")
        cur = new
clean = cur

total = 규칙.diff(df, clean)
t1, t2, t3, _ = st.columns([1, 1, 1, 3])
t1.metric("정제 후 줄 수", f"{len(clean):,}", delta=f"{len(clean) - len(df):,}" if len(clean) != len(df) else None)
t2.metric("정제 후 칸 수", f"{len(clean.columns):,}",
          delta=f"{len(clean.columns) - len(df.columns):+,}" if len(clean.columns) != len(df.columns) else None)
t3.metric("바뀐 칸", f"{total['바뀐 칸']:,}")

# ================================================================ ③ 칸 비교
st.header("③ 원본 ↔ 정제 후 비교")
all_cols = list(df.columns) + [c for c in clean.columns if c not in df.columns]
col = st.selectbox("비교할 칸", all_cols, key="cmp_col")
a = 읽기.column_stats(df[col]) if col in df.columns else None
b = 읽기.column_stats(clean[col]) if col in clean.columns else None

if a is None:
    st.info(f"«{col}» 칸은 정제 과정에서 새로 생긴 칸입니다.")
elif b is None:
    st.warning(f"«{col}» 칸은 정제 과정에서 없어졌습니다.")
if a and b:
    k1, k2, k3, k4 = st.columns(4)

    def dlt(x, y):
        return None if x is None or y is None or abs(y - x) < 1e-12 else round(y - x, 4)

    k1.metric("평균", fmt(b["평균"]), delta=dlt(a["평균"], b["평균"]), delta_color="off",
              help=f"원본 평균 {fmt(a['평균'])} (숫자로 읽히는 값만으로 계산)")
    k2.metric("숫자로 읽히는 값", f"{b['숫자로 읽히는 값']:,}", delta=dlt(a["숫자로 읽히는 값"], b["숫자로 읽히는 값"]))
    k3.metric("빈칸", f"{b['빈칸']:,}", delta=dlt(a["빈칸"], b["빈칸"]), delta_color="off")
    k4.metric("값의 종류 수", f"{b['값의 종류 수']:,}", delta=dlt(a["값의 종류 수"], b["값의 종류 수"]), delta_color="off")
    if a["평균"] is not None or b["평균"] is not None:
        st.markdown(f"#### {col} 평균 {fmt(a['평균'])} → {fmt(b['평균'])}")

keys = ["줄 수", "값 있는 칸", "빈칸", "값의 종류 수", "숫자로 읽히는 값", "숫자 비율",
        "평균", "중앙값", "최소", "최대", "합계", "표준편차"]


def cell(st_, k):
    if st_ is None:
        return "—"
    v = st_[k]
    return f"{v:.1%}" if k == "숫자 비율" else fmt(v)


def change(k):
    if not a or not b or a[k] is None or b[k] is None:
        return "—"
    d = b[k] - a[k]
    if abs(d) < 1e-12:
        return "0"
    return f"{d:+.1%}" if k == "숫자 비율" else ("+" if d > 0 else "") + fmt(d)


g1, g2 = st.columns([3, 2])
g1.dataframe(pd.DataFrame({"지표": keys, "원본": [cell(a, k) for k in keys],
                           "정제 후": [cell(b, k) for k in keys], "변화": [change(k) for k in keys]}),
             hide_index=True, width="stretch")
with g2:
    for title, s_ in (("원본 · 많이 나온 값", a), ("정제 후 · 많이 나온 값", b)):
        if s_ is not None:
            st.caption(title)
            vc = s_["많이 나온 값"]
            st.dataframe(pd.DataFrame({"값": [show(v) for v in vc.index], "개수": vc.to_numpy()}),
                         hide_index=True, width="stretch")

if a and b:
    common = df.index.intersection(clean.index)
    x, y = df.loc[common, col], clean.loc[common, col]
    ch = x != y
    removed = len(df.index.difference(clean.index))
    st.markdown(f"**바뀐 값 {int(ch.sum()):,}개**" + (f" · 정제로 지운 줄 {removed:,}개" if removed else ""))
    if ch.any():
        st.dataframe(pd.DataFrame({"원본 줄": common[ch.to_numpy()], "원본 값": x[ch].map(show).to_numpy(),
                                   "정제 후 값": y[ch].map(show).to_numpy()}).head(500),
                     hide_index=True, width="stretch", height=280)

# ---------------------------------------------------------------- 미리보기
with st.expander("데이터 미리보기 (앞 100줄) — 노란 칸 = 정제로 바뀐 값"):
    tab_o, tab_c = st.tabs(["원본", "정제 후"])
    tab_o.dataframe(df.head(100), width="stretch")
    view = clean.head(100)
    before = df.reindex(view.index)
    mask = pd.DataFrame(False, index=view.index, columns=view.columns)
    for c in view.columns:
        if c in before.columns:
            mask[c] = before[c].notna() & (before[c] != view[c])
    hl = pd.DataFrame(np.where(mask.to_numpy(), "background-color:#fff1b8;color:#1b1f24", ""),
                      index=view.index, columns=view.columns)
    tab_c.dataframe(view.style.apply(lambda _: hl, axis=None), width="stretch")

# ================================================================ ④ 내려받기
st.header("④ 정제된 데이터 내려받기")
base = up.name.rsplit(".", 1)[0] + "_정제"
d1, d2 = st.columns([2, 3])
name = d1.text_input("파일 이름", base).strip() or base
with_row = d2.checkbox("원본 줄 번호(_row) 칸 넣기", value=False)
out = clean.reset_index() if with_row else clean
csv_bytes = out.to_csv(index=False).encode("utf-8-sig")  # 엑셀에서 한글이 깨지지 않도록 BOM
json_bytes = json.dumps(
    [{k: (None if v == "" else v) for k, v in rec.items()} for rec in out.to_dict(orient="records")],
    ensure_ascii=False, indent=2, default=str).encode("utf-8")
b1, b2, _ = st.columns([1, 1, 4])
b1.download_button("CSV 내려받기", csv_bytes, f"{name}.csv", "text/csv", type="primary")
b2.download_button("JSON 내려받기", json_bytes, f"{name}.json", "application/json", type="primary")
st.caption("모든 값은 글자로 내려갑니다 (예: 89000). JSON의 빈칸은 null 입니다.")
