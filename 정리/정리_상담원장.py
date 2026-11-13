"""상담원장_2025.xlsx → 정리본/상담원장_정리.csv · 정리본/지운줄.csv · 정리본/변경요약.md

결정 기록: d-005  (my-wiki/decisions/d-005-상담원장-정리의-다섯-결정.md)
  d-005-1  «종결»을 «완료»로, «대기»를 «보류»로 합친다 (CS기획팀 답)
  d-005-2  처리시간 음수는 지우지 않는다. 값은 그대로 두고 맨 끝 _표시 칸에 이유를 적는다
  d-005-3  만족도 6~10점은 지우지 않고 _표시 칸에 이유를 적는다. 평균 만족도는 표시한 줄을 빼고 낸다
  d-005-4  칸 이름은 원본 그대로 둔다 (24칸). 더하는 칸은 맨 끝 _표시 하나뿐이다
  d-005-5  고객번호 앞뒤 공백을 없앤다

위 다섯 가지 밖의 처리(표기 통일, 테스트·중복·이전 상태 줄 지움, 상담시간 초→분 등)는
d-005 이전에 정한 그대로이며 이 파일에서 새로 더한 것이 없다. 그 처리의 결정 기록은 없다.

원본은 읽기만 한다. 쓰는 것은 아래 세 파일뿐이다.
  실행 (cleaner 폴더에서):  python -I 정리/정리_상담원장.py ../messy/상담원장_2025.xlsx 정리본
"""
import sys
from collections import Counter

import numpy as np
import pandas as pd

src, outdir = sys.argv[1], sys.argv[2]

raw = pd.read_excel(src, sheet_name="원장", dtype=object)      # 읽기만 한다
raw.insert(0, "_row", range(2, len(raw) + 2))                   # 엑셀 행 번호 (머리글 = 1행)
ORIG_COLS = list(raw.columns[1:])
df = raw.copy()
removed = {}                                                    # 엑셀 행 번호 → 지운 이유
R_TEST = "테스트 접수 (상담원ID AG999)"
R_DUP = "완전 중복 (모든 칸이 같은 줄이 이미 있음)"
R_PRIOR = "같은 접수번호의 이전 상태 줄 (최종수정일이 더 늦은 줄이 있음)"

# 1. 앞뒤 공백 지움 — d-005-5: 고객번호 포함 (글자 칸 전체에 같은 처리를 한다)
for c in df.columns[1:]:
    m = df[c].map(lambda v: isinstance(v, str) and v != v.strip())
    if m.any():
        df.loc[m, c] = df.loc[m, c].str.strip()
assert (df["고객번호"] == df["고객번호"].str.strip()).all()

# 2. 접수일 표기 → YYYY-MM-DD
df["접수일"] = df["접수일"].str.replace(r"[/.]", "-", regex=True)
pd.to_datetime(df["접수일"], format="%Y-%m-%d")                 # 없는 날짜면 여기서 멈춘다

# 3. 같은 뜻 다른 표기 (d-005 이전 정리 그대로)
MAPS = {
    "접수채널": {"유선": "전화", "전화상담": "전화", "TEL": "전화",
               "chatbot": "챗봇", "Chatbot": "챗봇", "챗 봇": "챗봇",
               "app": "앱", "APP": "앱", "모바일앱": "앱",
               "e-mail": "이메일", "E-mail": "이메일", "메일": "이메일"},
    "상담원팀": {"요금 팀": "요금팀", "품질 팀": "품질팀", "가입 팀": "가입팀", "해지방어 팀": "해지방어팀"},
    "지역": {"경기도": "경기", "서울시": "서울", "서울특별시": "서울"},
    "처리상태": {"처리완료": "완료", "처리중": "진행중"},
    "재문의": {"y": "Y", "예": "Y", "TRUE": "Y", "n": "N", "아니오": "N", "FALSE": "N"},
}
for c, mp in MAPS.items():
    df[c] = df[c].replace(mp)

# 4. 월정액 '33,000원' → 33000
df["월정액"] = pd.to_numeric(df["월정액"].astype(str).str.replace(r"[,원]", "", regex=True),
                           errors="coerce").astype("Int64")

# 5. 테스트 접수 지움 (작성기준: 상담원ID AG999 는 시스템 점검용 테스트 접수)
m = df["상담원ID"] == "AG999"
for r in df.loc[m, "_row"]:
    removed[r] = R_TEST
df = df[~m]

# 6. 완전 중복 지움 (표기를 맞춘 뒤 모든 칸이 같은 줄. 먼저 나온 줄을 남긴다)
cols = [c for c in df.columns if c != "_row"]
m = df[cols].astype(str).duplicated()
for r in df.loc[m, "_row"]:
    removed[r] = R_DUP
df = df[~m]

# 7. 같은 접수번호는 최종수정일이 가장 늦은 줄만 남긴다
df = df.sort_values(["접수번호", "최종수정일"])
m = df["접수번호"].duplicated(keep="last")
for r in df.loc[m, "_row"]:
    removed[r] = R_PRIOR
df = df[~m]

# 8. 상담시간: 2025-07-01 이후 전화 줄은 초 단위로 적혀 있어 분으로 바꾼다 (작성기준)
tel = (df["접수채널"] == "전화") & (df["접수일"] >= "2025-07-01")
df["상담시간"] = pd.to_numeric(df["상담시간"]).astype(float)
df.loc[tel, "상담시간"] = (df.loc[tel, "상담시간"] / 60).round(2)

# 9. 만족도: 0·'-'·'미응답'은 응답 안 함(빈칸), '5점'은 5 (작성기준 · d-005 이전 정리 그대로)
s = df["만족도"].astype(str)
m0 = df["만족도"].isin([0, "0", "-", "미응답"])
sat = pd.to_numeric(s.str.replace("점", ""), errors="coerce")
sat[m0] = pd.NA
df["만족도"] = sat.astype("Int64")
# d-005-3: 6~10점은 지우지 않는다 (예전에는 여기서 빈칸으로 바꿨다)
hi = (df["만족도"] > 5).fillna(False).astype(bool)

# 10. d-005-1: 종결 → 완료, 대기 → 보류
n_jong = int((df["처리상태"] == "종결").sum())
n_dae = int((df["처리상태"] == "대기").sum())
kinds_before = int(df["처리상태"].nunique())
df["처리상태"] = df["처리상태"].replace({"종결": "완료", "대기": "보류"})

# 11. d-005-2 · d-005-3: 지우지 않는 값은 맨 끝 _표시 칸에 이유를 적는다 (한 줄에 둘이면 « · »로 잇는다)
pt = pd.to_numeric(df["처리시간"], errors="coerce")
neg = (pt < 0).fillna(False).astype(bool)
marks = {i: [] for i in df.index}
for i in df.index[neg]:
    marks[i].append(f"처리시간 음수({pt[i]:g}) — 틀린 값이지만 진짜 값을 몰라 그대로 둠")
for i in df.index[hi]:
    marks[i].append(f"만족도 {int(df.at[i, '만족도'])}점 — 1~5점 척도 밖. NPS를 잘못 넣은 것으로 보이나 확인 못 해 그대로 둠. 평균 만족도에서 뺌")
df["_표시"] = pd.Series({i: " · ".join(v) for i, v in marks.items()})

# 12. 저장 — 칸 이름은 원본 그대로 24칸 + 맨 끝 _표시 (d-005-4)
final = df.sort_values("_row")
hi_f = hi.loc[final.index]
neg_f = neg.loc[final.index]
out = final.drop(columns="_row")
assert list(out.columns) == ORIG_COLS + ["_표시"]
out.to_csv(f"{outdir}/상담원장_정리.csv", index=False, encoding="utf-8-sig", float_format="%g")

rem = raw[raw["_row"].isin(list(removed))].copy()               # 지운 줄은 원본 값 그대로
rem["지운이유"] = rem["_row"].map(removed)
rem = rem.rename(columns={"_row": "원본행"})
rem.to_csv(f"{outdir}/지운줄.csv", index=False, encoding="utf-8-sig")

# ── 변경요약.md ───────────────────────────────────────────────
fm = lambda x: f"{int(x):,}"
n0, n1 = len(raw), len(final)
cnt = Counter(removed.values())
assert n0 - len(removed) == n1

def tostr(x):
    x = x.astype(object)
    return x.where(x.notna(), "").astype(str)

rawk = raw.set_index("_row").loc[final["_row"].values].reset_index(drop=True)
fin = final.reset_index(drop=True)
changed = {}
for c in ORIG_COLS:
    a, b = rawk[c], fin[c]
    if c in ("상담시간", "처리시간"):
        x, y = pd.to_numeric(a, errors="coerce"), pd.to_numeric(b, errors="coerce")
        diff = ~((x == y) | (x.isna() & y.isna()))
    else:
        diff = tostr(a) != tostr(b)
    changed[c] = int(diff.sum())

WHAT = {
    "접수일": "`2025/01/01` · `2025.01.01` 꼴을 `2025-01-01` 로",
    "고객번호": "앞 공백 제거 (d-005-5)",
    "접수채널": "17가지 이름을 전화·챗봇·앱·이메일로 통일, 앞 공백 제거",
    "상담원팀": "띄어쓰기 통일 (요금 팀→요금팀 …)",
    "지역": "경기도→경기, 서울시·서울특별시→서울",
    "처리상태": "처리완료→완료, 처리중→진행중, 뒤 공백 제거, 종결→완료·대기→보류 (d-005-1)",
    "재문의": "y·예·TRUE→Y, n·아니오·FALSE→N",
    "월정액": "`33,000원` → 33000",
    "상담시간": "2025-07-01 이후 전화 줄을 초에서 분으로 (아래)",
    "만족도": "0·`-`·`미응답`→빈칸, `5점`→5",
}
for c, k in changed.items():
    assert k == 0 or c in WHAT, c

tel_f = (fin["접수채널"] == "전화") & (fin["접수일"] >= "2025-07-01")
before = pd.to_numeric(rawk["상담시간"])
after = fin["상담시간"].astype(float)
ex = fin.loc[tel_f, ["접수번호"]].head(3).copy()
ex["전"], ex["후"] = before[tel_f].head(3).values, after[tel_f].head(3).values
pad_cust_all = int((raw["고객번호"] != raw["고객번호"].str.strip()).sum())
pad_cust_kept = int((tostr(rawk["고객번호"]) != tostr(fin["고객번호"])).sum())
cust_after = int(fin["고객번호"].nunique())
cust_before = int(rawk["고객번호"].nunique())
st_pairs = Counter((a.strip(), b) for a, b in zip(rawk["처리상태"], fin["처리상태"]) if a.strip() != b)
sat_use = fin.loc[~hi_f.values, "만족도"].dropna().astype(float)
re_ = fin["재문의"].dropna()
metrics = {
    "상담 건수": n1,
    "평균 상담시간(분)": round(float(fin["상담시간"].mean()), 2),
    "재문의율(%)": round(float((re_ == "Y").mean() * 100), 2),
    "평균 만족도": round(float(sat_use.mean()), 2),
    "완료율(%)": round(float((fin["처리상태"] == "완료").mean() * 100), 2),
}

L = []
A = L.append
A("# 상담원장 변경 요약")
A("")
A("- 원본: `../messy/상담원장_2025.xlsx` 시트 `원장` (읽기만 했다)")
A("- 결과: `정리본/상담원장_정리.csv` · 지운 줄: `정리본/지운줄.csv`")
A("- 스크립트: `정리/정리_상담원장.py`")
A("- 결정 기록: d-005 (`my-wiki/decisions/d-005-상담원장-정리의-다섯-결정.md`)")
A("- 아래 «정리본 줄 기준» 숫자는 정리본 %s줄 가운데 원본 값과 달라진 칸의 수다." % fm(n1))
A("")
A("## 1. 줄 수")
A("")
A("| 단계 | 줄 수 | 계산 |")
A("|---|---|---|")
a = n0
A(f"| 원본 | {fm(n0)} | 엑셀 2~{fm(n0 + 1)}행 |")
for r in (R_TEST, R_DUP, R_PRIOR):
    A(f"| {r} | −{fm(cnt[r])} | {fm(a)} − {fm(cnt[r])} = {fm(a - cnt[r])} |")
    a -= cnt[r]
A(f"| 정리본 | {fm(n1)} | 지운 줄 합계 {fm(len(removed))} |")
A("")
A("## 2. 결정 기록 d-005 대로 바꾼 것")
A("")
A("| 결정 | 한 일 | 몇 줄 |")
A("|---|---|---|")
A(f"| d-005-1 | «종결»→«완료», «대기»→«보류». 처리상태 값 종류 {kinds_before}종 → {int(fin['처리상태'].nunique())}종 | 종결 {fm(n_jong)} + 대기 {fm(n_dae)} = **{fm(n_jong + n_dae)}**줄 (정리본 줄 기준) |")
A(f"| d-005-2 | 처리시간 음수는 지우지 않고 값 그대로. `_표시` 칸에 이유를 적음 | 값이 바뀐 줄 0 · 표시한 줄 **{fm(neg_f.sum())}** ({', '.join(fin.loc[neg_f.values, '접수번호'])}) |")
A(f"| d-005-3 | 만족도 6~10점은 지우지 않고 원래 값 그대로. `_표시` 칸에 이유를 적음. 평균 만족도에서는 뺌 | 표시한 줄 **{fm(hi_f.sum())}** · 평균에서 뺀 줄 {fm(hi_f.sum())} |")
A(f"| d-005-4 | 칸 이름은 원본 {len(ORIG_COLS)}칸 그대로. 더한 칸은 맨 끝 `_표시` 하나 | 이름이 바뀐 칸 0 · `_표시`에 내용이 있는 줄 {fm((out['_표시'] != '').sum())} (겹치는 줄 {fm((neg_f & hi_f).sum())}) |")
A(f"| d-005-5 | 고객번호 앞뒤 공백 제거. 고객 수 {fm(cust_before)}명 → {fm(cust_after)}명 | 정리본 줄 기준 **{fm(pad_cust_kept)}**줄 (원본 전체 {fm(n0)}줄 중 {fm(pad_cust_all)}줄) |")
A("")
A("`_표시` 문구는 결정 기록에 정해져 있지 않아, 결정에 적힌 이유를 그대로 옮겨 적었다.")
A("")
A("## 3. 상담시간 칸 — 7월 이후 전화 줄은 초에서 분으로 바뀌었다")
A("")
A(f"- 대상: 접수일이 2025-07-01 이후이고 접수채널이 «전화»(유선·전화상담·TEL 포함)인 **{fm(tel_f.sum())}줄** (정리본 줄 기준). 작성기준: «2025-07-01 전화 시스템 교체 뒤로 전화 채널은 초 단위로 기록됨»")
A("- 한 일: 60으로 나누고 소수 둘째 자리까지 반올림. 다른 줄의 상담시간은 바꾸지 않았다. 정리본의 상담시간은 모두 분 단위다.")
A(f"- 이 줄들의 평균: 바꾸기 전 {before[tel_f].mean():,.1f}(초) → 바꾼 뒤 {after[tel_f].mean():,.2f}(분)")
A(f"- 60으로 나눈 뒤 소수가 된 줄: {fm((after[tel_f] != after[tel_f].round(0)).sum())}")
A("- 예:")
for _, r in ex.iterrows():
    A(f"  - {r['접수번호']}: {r['전']:g} → {r['후']:g}")
A("")
A("## 4. 칸별로 원본 값과 달라진 줄 (정리본 줄 기준)")
A("")
A("| 칸 | 달라진 줄 | 무엇을 |")
A("|---|---|---|")
for c in ORIG_COLS:
    if changed[c]:
        A(f"| {c} | {fm(changed[c])} | {WHAT[c]} |")
A("")
A("원본과 같은 칸: " + " · ".join(c for c in ORIG_COLS if not changed[c]))
A("")
A("처리상태가 달라진 줄의 «원래 값 → 바뀐 값» (앞뒤 공백만 다른 줄은 뺐다)")
A("")
A("| 원래 | 바뀐 | 줄 |")
A("|---|---|---|")
for (a_, b_), k in sorted(st_pairs.items(), key=lambda kv: (-kv[1], kv[0])):
    A(f"| {a_} | {b_} | {fm(k)} |")
A("")
A("## 5. 결과 숫자")
A("")
A("| 숫자 | 값 | 센 기준 |")
A("|---|---|---|")
A(f"| 상담 건수 | {fm(metrics['상담 건수'])} | 정리본 전체 줄 |")
A(f"| 평균 상담시간(분) | {metrics['평균 상담시간(분)']:.2f} | 정리본 전체 줄 |")
A(f"| 재문의율(%) | {metrics['재문의율(%)']:.2f} | 재문의가 적힌 {fm(len(re_))}줄 가운데 Y (빈칸은 분모에서 뺌) |")
A(f"| 평균 만족도 | {metrics['평균 만족도']:.2f} | 응답 {fm(len(sat_use))}줄 (표시한 {fm(hi_f.sum())}줄 뺌) |")
A(f"| 완료율(%) | {metrics['완료율(%)']:.2f} | 완료 {fm((fin['처리상태'] == '완료').sum())} ÷ {fm(n1)} (종결 포함) |")
A(f"| 고객 수 | {fm(cust_after)} | 고객번호 종류 (공백 제거 뒤) |")
open(f"{outdir}/변경요약.md", "w", encoding="utf-8", newline="\n").write("\n".join(L) + "\n")
