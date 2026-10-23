#!/usr/bin/env python3
"""Multi-scale façade fire risk workbook: report/다척도_위험성평가_계산서.xlsx

Specimen (cone, TGA-DSC, FT-IR) -> mid-scale cavity test -> simulation -> building risk, each scale stated in the
mid-scale equivalent (t600 per storey, p_B). TGA and the cavity tests are filled in; the cone, FT-IR and simulation
sheets are templates whose indices are cell formulas (one grey example row each, made-up values, to show the format).
The TGA -> t600 / p_B transfer, the Bayesian combination with the cavity tests and the consistency checks are formulas;
the Monte Carlo building risk comes from multiscale.py as values.

    python ../multiscale.py && python build_ms_xlsx.py
    python <xlsx skill>/scripts/recalc.py 다척도_위험성평가_계산서.xlsx
"""
from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

HERE = Path(__file__).resolve().parent
BASE = HERE.parent
ROOT = BASE.parent
OUT = HERE / "다척도_위험성평가_계산서.xlsx"
D = json.loads((HERE / "ms_data.json").read_text(encoding="utf-8"))

ALL = ["EPS_C", "EPS_SC", "PIR_C", "PIR_SC", "PF_out", "PF_in"]
MATS = D["materials"]
CHAR = D["char"]
LABEL = D["label"]
GRADE = D["grade"]

FONT = "Arial"
F_BASE = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=14, bold=True)
F_SUB = Font(name=FONT, size=11, bold=True)
F_NOTE = Font(name=FONT, size=9, color="4B5563")
F_INPUT = Font(name=FONT, size=10, color="0000FF")      # hardcoded inputs / measured values
F_LINK = Font(name=FONT, size=10, color="008000")       # links to another sheet
F_EX = Font(name=FONT, size=10, color="6B7280", italic=True)
FILL_HEAD = PatternFill("solid", fgColor="23303D")
FILL_TODO = PatternFill("solid", fgColor="FFF2CC")       # cells to fill in
FILL_KEY = PatternFill("solid", fgColor="E2F3E2")        # key results
FILL_GREY = PatternFill("solid", fgColor="F3F4F6")       # section bands, example rows
FILL_MISS = PatternFill("solid", fgColor="FDE2E2")
FILL_HAVE = PatternFill("solid", fgColor="E2F3E2")
FILL_PART = PatternFill("solid", fgColor="FEF2D3")
THIN = Side(style="thin", color="D1D5DB")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="center")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
SCI = "0.00E+00"


def cell(ws, ref, value, font=F_BASE, fill=None, fmt=None, align=WRAP, border=True):
    c = ws[ref]
    c.value = value
    c.font = font
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    c.alignment = align
    if border:
        c.border = BOX
    return c


def rc(ws, row, col, value, **kw):
    return cell(ws, f"{get_column_letter(col)}{row}", value, **kw)


def header(ws, row, names, col0=1, widths=None, height=None):
    for k, n in enumerate(names):
        rc(ws, row, col0 + k, n, font=F_HEAD, fill=FILL_HEAD, align=CENTER)
        if widths:
            ws.column_dimensions[get_column_letter(col0 + k)].width = widths[k]
    if height:
        ws.row_dimensions[row].height = height


def title(ws, text, sub=None):
    cell(ws, "A1", text, F_TITLE, border=False, align=Alignment(vertical="center"))
    if sub:
        cell(ws, "A2", sub, F_NOTE, border=False, align=Alignment(vertical="center"))


def band(ws, row, text, ncol):
    rc(ws, row, 1, text, font=F_SUB, fill=FILL_GREY)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncol)


def note(ws, row, text, col=1):
    rc(ws, row, col, text, font=F_NOTE, border=False, align=Alignment(vertical="center"))


def define(wb, name, sheet, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=f"'{sheet}'!{ref}")


def A(col, row):
    return f"${get_column_letter(col)}${row}"


def L(c):
    return get_column_letter(c)


# --------------------------------------------------------------------------------------------- 읽는 법
def readme(wb):
    ws = wb.active
    ws.title = "읽는 법"
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 116
    title(ws, "외벽 화재위험성 다척도 평가 계산서 — 소형시료(콘·TGA–DSC·FT-IR) · 중규모 · 시뮬레이션 → 건물 위험도")
    rows = [
        ("목적", "소형시료 시험, 중규모 성능시험, 시뮬레이션이 각각 화재위험성평가 결과를 내고, 같은 단위로 결합되도록 만든 계산서입니다. "
                 "TGA–DSC는 확산(r_max,air)·지속(f400,air)·급발(용융형) 인자로 위험도에 정량 반영됩니다."),
        ("공통 단위", "모든 척도의 결과를 '중규모 등가값'으로 바꿉니다: 층당 확산시간 t600(3 m 중공층 시험의 3000 mm 600 K 도달시각)과 "
                      "버너 소화 후 자립 연소 확률 p_B. 이 두 값이 건물 위험도 모델(facade_fra, NFPA 551 정량법 R = Σ F·C)의 재료 입력입니다."),
        ("척도", "L1 소형시료: 콘칼로리미터(착화·발열·성장, 국내 준불연 판정), TGA–DSC(분해속도·탄화층 산화·용융형), FT-IR(고분자·난연제 식별, 유해가스) / "
                 "L2 중규모: 3 m 중공층 시험 / L3 시뮬레이션: FDS(PyroSim)로 중규모 재현 → 건물 외벽 시나리오, 시험→실물 비 s 확정 / "
                 "L4 건물: 사건수목 × 층간 확산 모델, 2중 몬테카를로"),
        ("게이트", "가연성 등급(콘 준불연 기준 미달)이면 TGA 전이함수를 쓰지 않습니다(가연성 PIR(PUR): TGA 예측 812 s, 시험 30 s). "
                   "용융형(f400,air < 0.5, EPS)이면 확산 전이함수를 쓰지 않습니다. FT-IR은 고분자·난연제를 확인해 게이트를 검증합니다."),
        ("칸 색", "노란 바탕 = 입력할 칸 · 파란 글자 = 입력·측정값 · 검은 글자 = 계산식 · 초록 글자 = 다른 시트에서 가져온 값 · 연두 바탕 = 주요 결과 · "
                  "회색 기울임 = 형식 예시(가상값, 계산식 확인용 — 실측 아님)"),
        ("시트", "1 체계(척도 × 위험차원) · 2 콘(양식) · 3 TGA(입력됨) · 4 FT-IR(양식) · 5 중규모(입력됨) · 6 전이함수(상수·통계) · "
                 "7 통합(베이즈 결합·정합성) · 8 시뮬레이션(양식) · 9 시뮬레이션 정보가치 · 10 위험도 결과 · 11 EFFECT 비교 · 12 시험계획 · 13 검증"),
        ("현재 자료", "있음: TGA–DSC 6종(공기·질소 각 2회), 중공층 시험 5종 11회(+ EPS 가연성 4분 기록 2회). "
                      "없음: 콘칼로리미터, FT-IR, FDS 결과(양식과 계산식만 제공). 빈도·분기확률 등 건물 입력은 facade_fra 예시값입니다."),
        ("결과 요약", "TGA 예측은 준불연 탄화형 3종에서 시험과 정합(z −1.5 – 0.9), 가연성 PIR(PUR)에서 크게 불일치(z ≈ −18.5) → 게이트가 필요. "
                      "재료 평균 t600의 90 % 구간 폭은 등급만 5.3배 → TGA 1.7배 → 중규모 1.24배 → 통합 1.22배로 줄어듭니다. "
                      "건물 위험도 구간은 시험→실물 비 s가 지배 → L3(시뮬레이션)가 다음 단계입니다."),
        ("재현", "python multiscale_fra/multiscale.py → python multiscale_fra/report/build_ms_xlsx.py → recalc. 엑셀 수식과 파이썬 값의 일치는 13_검증."),
    ]
    for i, (k, v) in enumerate(rows, start=3):
        cell(ws, f"A{i}", k, F_BOLD)
        cell(ws, f"B{i}", v)
        ws.row_dimensions[i].height = 46


# --------------------------------------------------------------------------------------------- 1 체계
MATRIX = [
    ("D1 착화성", "t_ig(50 kW/m²), TRP·CHF(다열류속) [미측정]", "T_onset,air·T_5% [측정] — 중규모 착화는 시스템 상수(준불연 188–288 s)라 예측력 약함",
     "난연제(P=O) 유무 [미측정]", "t_ig(시스템 착화) [측정]", "외벽 착화시각 [미수행]", "t600 = t_ig + t_sp에 포함"),
    ("D2 확산성", "pHRR/t_pHRR, b = k_f·Q̇″ − 1 − t_b/t_ig, M = Q̇″^(1/3)/TRP [미측정]", "r_max,air → t̂_sp = C_sys/r_max (준불연 탄화형) [측정]",
     "이소시아누레이트 지수(PIR) [미측정]", "t_sp, t600, FRI, P [측정]", "층별 도달시각 → s [미수행]", "τ_F = s·t600"),
    ("D3 발열성", "THR_600(준불연 ≤ 8 MJ/m²), pHRR, EHC [미측정]", "절대량 불가(초기 질량 없음), DSC 미보정 [불가]", "—",
     "H600(3000 mm 초과 열량), 최고온도 [측정]", "외벽 HRR, 창 열류속 [미수행]", "(현 모델 미반영: 확장 항목)"),
    ("D4 지속성", "연소 종료·작열 시간 [미측정]", "f400,air(탄화층 산화 몫) → p̂_B [측정]", "—", "플래그 B(소화 후 ssr ≥ 1, 3/6 높이) [측정]",
     "화원 소멸 후 연소 [미수행]", "p_B (자립이면 T_stop = ∞)"),
    ("D5 급발성", "용융·적하 관찰 [미측정]", "용융형(f400,air < 0.5) [측정]", "—", "플래그 A(τ300→600 ≤ 30 s) [측정]", "용융 흐름은 모델 한계 [불가]",
     "(현 모델 미반영: 표시 항목)"),
    ("D6 유해성", "CO 수율, TSP [미측정]", "—", "TGA–FTIR: HCN·CO·NCO·HCl [미측정]", "가스 측정 없음 [미측정]", "FED(위층 세대) [미수행]",
     "(확장: 노출 인원 → 유해 노출)"),
    ("G 조성·등급", "국내 준불연 판정(THR_600 ≤ 8 등) [미측정 · 인증서 등급 사용 중]", "탄화형/용융형 [측정]", "고분자 식별, 난연제 확인 [미측정]",
     "가연성 PIR(PUR) 30 s 전면 착화 [측정]", "—", "게이트: 전이함수 적용 여부"),
]


def matrix(wb):
    ws = wb.create_sheet("1_체계")
    title(ws, "① 다척도 평가 체계 — 6개 위험차원 × 척도별 지표와 자료 상태",
          "[측정] 보유 · [미측정] 시험 필요 · [미수행] 해석 필요 · [불가] 현 방법으로 얻을 수 없음. 마지막 열은 건물 위험도 모델에 들어가는 자리입니다.")
    cols = ["위험 차원", "L1 콘칼로리미터", "L1 TGA–DSC", "L1 FT-IR", "L2 중규모(3 m 중공층)", "L3 시뮬레이션(FDS)", "L4 건물 모델 입력"]
    header(ws, 4, cols, widths=[13, 30, 30, 24, 26, 22, 22], height=30)
    for i, row in enumerate(MATRIX, start=5):
        for j, v in enumerate(row, start=1):
            fill = None
            if j > 1 and j < 7:
                fill = FILL_HAVE if "[측정]" in v else (FILL_MISS if "[미측정]" in v or "[미수행]" in v else (FILL_GREY if v in ("—",) or "[불가]" in v else None))
            rc(ws, i, j, v, font=F_BOLD if j == 1 else F_BASE, fill=fill)
        ws.row_dimensions[i].height = 58
    note(ws, 13, "칸 색: 연두 = 자료 있음, 분홍 = 시험·해석 필요, 회색 = 해당 없음·불가. 지표는 모두 '중규모 등가값'(t600, p_B)으로 바꿔 비교합니다.")


# --------------------------------------------------------------------------------------------- 2 콘
def cone(wb):
    ws = wb.create_sheet("2_L1_콘")
    title(ws, "② L1 콘칼로리미터 (ISO 5660-1) — 입력 양식과 지표 계산",
          "50 kW/m² 10분 결과로 국내 준불연 판정(게이트)을, 25·35·50·65 kW/m² 착화시간으로 TRP·CHF를 구합니다. 회색 기울임 행은 형식 예시(가상값)입니다.")
    r = 4
    band(ws, r, "변수", 8)
    params = [("A_exp", "노출 면적 (ISO 5660-1, 받침틀 사용)", 0.0088, "m²", "0.0000"),
              ("k_f", "벽면 화염 길이 계수 (Quintiere)", 0.01, "m²/kW", "0.000"),
              ("thr_lim", "준불연: 10분 총방출열량 상한", 8, "MJ/m²", "0"),
              ("dur_lim", "준불연: 200 kW/m² 초과 연속시간 상한(미만)", 10, "s", "0"),
              ("t600_ref_cone", "콘 상대식 기준 t600 (PIR(PUR) 준불연 기하평균)", 1079, "s", "0")]
    for k, (name, item, val, unit, fmt) in enumerate(params, start=r + 1):
        rc(ws, k, 1, name, font=F_BOLD)
        rc(ws, k, 2, item)
        rc(ws, k, 3, val, font=F_INPUT, fill=FILL_TODO, fmt=fmt, align=CENTER)
        rc(ws, k, 4, unit, align=CENTER)
        define(wb, name, "2_L1_콘", A(3, k))
    r = r + len(params) + 2
    band(ws, r, "A. 시편별 기록 (반복 3회 권장, 면재 포함 · 심재 각각)", 22)
    cols = ["재료", "시편", "열류속 q″ (kW/m²)", "두께 (mm)", "면재", "t_ig (s)", "pHRR (kW/m²)", "t_pHRR (s)", "THR_600 (MJ/m²)",
            "THR_total (MJ/m²)", "질량감소 Δm (g)", "200 kW/m² 초과 최대 연속시간 (s)", "관통 균열·구멍·용융 (Y/N)", "TSP (m²)", "CO 수율 (kg/kg)",
            "FIGRA_콘 = pHRR/t_pHRR", "t_ig/pHRR (Petrella)", "EHC (MJ/kg)", "t_b = THR/pHRR (s)", "b (Quintiere)", "준불연 판정 (50 kW/m²)", "1/√t_ig"]
    header(ws, r + 1, cols, widths=[16, 8, 10, 8, 7, 8, 10, 9, 10, 10, 10, 13, 11, 8, 9, 12, 11, 9, 10, 9, 12, 9], height=44)
    first = r + 2
    ex = ["예시(가상값)", "E-1", 50, 50, "유", 42, 180, 65, 6.5, 14.0, 9.5, 0, "N", 1.2, 0.05]
    nrows = 25
    for i in range(nrows):
        rr = first + i
        is_ex = i == 0
        vals = ex if is_ex else [None] * 15
        for j, v in enumerate(vals, start=1):
            rc(ws, rr, j, v, font=F_EX if is_ex else F_INPUT, fill=FILL_GREY if is_ex else FILL_TODO,
               align=CENTER if j > 1 else WRAP, fmt="0.00" if j in (9, 10, 15) else None)
        f = lambda c: f"{L(c)}{rr}"
        forms = [
            f'=IF(OR({f(7)}="",{f(8)}=""),"",{f(7)}/{f(8)})',
            f'=IF(OR({f(6)}="",{f(7)}=""),"",{f(6)}/{f(7)})',
            f'=IF(OR({f(10)}="",{f(11)}=""),"",{f(10)}*A_exp/({f(11)}/1000))',
            f'=IF(OR({f(10)}="",{f(7)}=""),"",{f(10)}*1000/{f(7)})',
            f'=IF(OR({f(6)}="",{f(7)}="",{f(10)}=""),"",k_f*{f(7)}-1-{f(19)}/{f(6)})',
            f'=IF(OR({f(3)}<>50,{f(9)}="",{f(12)}="",{f(13)}=""),"",IF(AND({f(9)}<=thr_lim,{f(12)}<dur_lim,{f(13)}="N"),"충족","미충족"))',
            f'=IF({f(6)}="","",1/SQRT({f(6)}))',
        ]
        fmts = ["0.00", "0.000", "0.0", "0", "0.00", None, "0.0000"]
        for j, (fm, nf) in enumerate(zip(forms, fmts), start=16):
            rc(ws, rr, j, fm, font=F_EX if is_ex else F_BASE, fmt=nf, align=CENTER, fill=FILL_GREY if is_ex else None)
    last = first + nrows - 1
    rng = lambda c: f"${L(c)}${first + 1}:${L(c)}${last}"       # excludes the example row
    r = last + 2
    note(ws, r, "FIGRA_콘·t_ig/pHRR: Schartel & Hull (2007), Petrella (1994) · b > 0이면 상향 확산 가속(Quintiere 2006, k_f ≈ 0.01 m²/kW) · "
                "준불연: 국토교통부 고시 「건축자재등 품질인정 및 관리기준」(10분 THR ≤ 8 MJ/m², 200 kW/m² 초과 10초 이상 연속 금지, 관통 균열·구멍·용융 없음; 가스유해성 별도)")
    # B: multi-flux TRP / CHF
    r += 2
    band(ws, r, "B. 다열류속 착화시간 → TRP·CHF (1/√t_ig = (q″ − CHF)/TRP, 열적으로 두꺼운 시편)", 8)
    header(ws, r + 1, ["재료", "q″ (kW/m²)", "평균 t_ig (s)", "1/√t_ig", "TRP (kW·s½/m²)", "CHF (kW/m²)"])
    rr = r + 2
    trp_ref = {}
    for m in ["예시(가상값)"] + ALL:
        is_ex = m.startswith("예시")
        tvals = [330, 120, 50, 30] if is_ex else [None] * 4
        rows_m = []
        for q, t in zip((25, 35, 50, 65), tvals):
            rc(ws, rr, 1, m if is_ex else LABEL[m], font=F_EX if is_ex else F_BOLD, fill=FILL_GREY if is_ex else None)
            rc(ws, rr, 2, q, font=F_INPUT, align=CENTER, fill=FILL_GREY if is_ex else None)
            rc(ws, rr, 3, t, font=F_EX if is_ex else F_INPUT, fill=FILL_GREY if is_ex else FILL_TODO, align=CENTER)
            rc(ws, rr, 4, f'=IF(C{rr}="","",1/SQRT(C{rr}))', fmt="0.0000", align=CENTER, fill=FILL_GREY if is_ex else None)
            rows_m.append(rr)
            rr += 1
        y, x = f"D{rows_m[0]}:D{rows_m[-1]}", f"B{rows_m[0]}:B{rows_m[-1]}"
        rc(ws, rows_m[0], 5, f'=IFERROR(1/SLOPE({y},{x}),"")', fmt="0", align=CENTER, fill=FILL_KEY)
        rc(ws, rows_m[0], 6, f'=IFERROR(-INTERCEPT({y},{x})/SLOPE({y},{x}),"")', fmt="0.0", align=CENTER, fill=FILL_KEY)
        if not is_ex:
            trp_ref[m] = (rows_m[0])
    # C: per-material summary and mid-scale equivalent
    r = rr + 1
    band(ws, r, "C. 재료별 요약 → 게이트와 중규모 등가 예측 (50 kW/m², 예시 행 제외)", 11)
    header(ws, r + 1, ["재료", "시편 수(50)", "t_ig (s)", "pHRR (kW/m²)", "THR_600 (MJ/m²)", "b", "준불연(전 시편)", "TRP", "M = Q̇″^(1/3)/TRP",
                       "t̂600_콘 (s)", "게이트"], height=32)
    rows_c = {}
    for k, m in enumerate(ALL):
        rr = r + 2 + k
        rows_c[m] = rr
    for m, rr in rows_c.items():
        lab = LABEL[m]
        crit = f'{rng(1)},$A{rr},{rng(3)},50'
        rc(ws, rr, 1, lab, font=F_BOLD)
        rc(ws, rr, 2, f'=COUNTIFS({crit},{rng(6)},">0")', align=CENTER)
        rc(ws, rr, 3, f'=IF(B{rr}=0,"",AVERAGEIFS({rng(6)},{crit},{rng(6)},">0"))', fmt="0", align=CENTER)
        rc(ws, rr, 4, f'=IF(B{rr}=0,"",AVERAGEIFS({rng(7)},{crit},{rng(7)},">0"))', fmt="0", align=CENTER)
        rc(ws, rr, 5, f'=IF(B{rr}=0,"",AVERAGEIFS({rng(9)},{crit},{rng(9)},">0"))', fmt="0.0", align=CENTER)
        rc(ws, rr, 6, f'=IF(B{rr}=0,"",AVERAGEIFS({rng(20)},{crit},{rng(6)},">0",{rng(7)},">0",{rng(10)},">0"))', fmt="0.00", align=CENTER)
        rc(ws, rr, 7, f'=IF(COUNTIFS({crit})=0,"",IF(COUNTIFS({crit},{rng(21)},"충족")=COUNTIFS({crit}),"충족","미충족"))', align=CENTER, fill=FILL_KEY)
        rc(ws, rr, 8, f"=E{trp_ref[m]}", font=F_LINK, fmt="0", align=CENTER)
        rc(ws, rr, 9, f'=IFERROR(D{rr}^(1/3)/H{rr},"")', fmt="0.0000", align=CENTER)
        rc(ws, rr, 10, f'=IFERROR(t600_ref_cone*(I{rows_c["PIR_SC"]}/I{rr})^2,"")', fmt="0", align=CENTER, fill=FILL_KEY)
        rc(ws, rr, 11, f'=IF(G{rr}="","(콘 미측정: 인증서 등급 사용)",IF(G{rr}="충족","준불연 → TGA 전이함수 적용 가능","가연성 → TGA 미적용, 중규모 필수"))')
    note(ws, r + 9, "t̂600_콘 = 1079 s × (M_PIR(PUR) 준불연 / M)²: cavity_fire_index/RISK_MODEL.md의 상대식(잠정). 콘 자료가 생기면 중규모 t600과 비교해 보정합니다.")
    ws.freeze_panes = "B5"


# --------------------------------------------------------------------------------------------- 3 TGA
def tga(wb):
    ws = wb.create_sheet("3_L1_TGA")
    title(ws, "③ L1 TGA–DSC (5 °C/min, 공기·질소 각 2회 평균) → 재료 인자와 중규모 등가 예측",
          "TGA 값은 tga_dsc/results/tga_summary.csv. 등급은 인증서 등급(콘 자료가 생기면 2_L1_콘의 판정으로 교체). 전이함수 상수는 6_전이함수.")
    cols = ["재료", "재료명", "등급", "T_onset,air (°C)", "T_5%,air (°C)", "r_max,air (%/°C)", "± 반폭", "f400,air", "± 반폭",
            "r_max,N₂", "f400,N₂", "공기 자료 판정", "재료형", "게이트", "t̂600 전체 보정 (s)", "t̂600 사용값 (s)", "FRI^ = 900/t̂600",
            "P̂", "m (램프)", "a₀ = n₀m", "b₀ = n₀(1−m)", "p̂_B 사전평균", "급발 예상"]
    header(ws, 4, cols, widths=[9, 16, 6, 10, 9, 10, 8, 8, 8, 9, 8, 10, 9, 22, 11, 11, 10, 6, 8, 8, 9, 9, 16], height=44)
    air = {p["material"]: p for p in D["profile"]}
    tga_sum = pd.read_csv(ROOT / "tga_dsc" / "results" / "tga_summary.csv")
    n2 = tga_sum[tga_sum.atmosphere == "N2"].set_index("material")
    ox = tga_sum[tga_sum.atmosphere.str.startswith("oxidative")].set_index("material")
    vko = {"reliable": "신뢰", "conditional": "조건부", "endpoint": "종점 의존", "remeasure": "재측정"}
    rows = {}
    for i, m in enumerate(ALL, start=5):
        rows[m] = i
        rc(ws, i, 1, m, font=F_BOLD)
        rc(ws, i, 2, LABEL[m])
        rc(ws, i, 3, GRADE[m], font=F_INPUT, fill=FILL_TODO, align=CENTER)
        for j, (val, fmt) in enumerate([(ox.loc[m, "T_onset_C_mean"], "0"), (ox.loc[m, "T_a05_C_mean"], "0"),
                                         (ox.loc[m, "r_max_pct_C_mean"], "0.000"), (ox.loc[m, "r_max_pct_C_halfrange"], "0.000"),
                                         (ox.loc[m, "f_above400_mean"], "0.000"), (ox.loc[m, "f_above400_halfrange"], "0.000"),
                                         (n2.loc[m, "r_max_pct_C_mean"], "0.000"), (n2.loc[m, "f_above400_mean"], "0.000")], start=4):
            rc(ws, i, j, float(val), font=F_INPUT, fmt=fmt, align=CENTER)
        rc(ws, i, 12, vko.get(air[m]["tga_verdict"], air[m]["tga_verdict"]), font=F_INPUT, align=CENTER)
        rc(ws, i, 13, f'=IF(H{i}<f_melt,"용융형","탄화형")', align=CENTER)
        rc(ws, i, 14, f'=IF(C{i}="c","가연성 → TGA 미적용",IF(H{i}<f_melt,"용융형 → 확산식 미적용","적용"))', fill=FILL_KEY)
        rc(ws, i, 15, f"=t_ig_sys+C_sys/F{i}", fmt="0", align=CENTER)
        rc(ws, i, 16, f'=IFERROR(INDEX(loo_t600,MATCH(A{i},loo_key,0)),O{i})', fmt="0", align=CENTER, fill=FILL_KEY)
        rc(ws, i, 17, f"=t_ref/P{i}", fmt="0.00", align=CENTER)
        rc(ws, i, 18, f'=IF(N{i}<>"적용","—",IF(Q{i}>=2,4,IF(Q{i}>=1,3,IF(Q{i}>=0.5,2,1))))', align=CENTER)
        rc(ws, i, 19, f"=MIN(0.95,MAX(0.05,(H{i}-f_lo)/(f_hi-f_lo)))", fmt="0.000", align=CENTER)
        rc(ws, i, 20, f"=n0*S{i}", fmt="0.000", align=CENTER)
        rc(ws, i, 21, f"=n0*(1-S{i})", fmt="0.000", align=CENTER)
        rc(ws, i, 22, f"=T{i}/(T{i}+U{i})", fmt="0.000", align=CENTER, fill=FILL_KEY)
        rc(ws, i, 23, f'=IF(M{i}="용융형","A 예상(용융·순간연소)",IF(C{i}="c","A 예상(가연성 전면착화)","—"))')
    define(wb, "tga_key", "3_L1_TGA", f"$A$5:$A${4 + len(ALL)}")
    note(ws, 12, "t̂600 사용값: 준불연 탄화형 3종은 해당 제품을 빼고 보정한 상수(6_전이함수의 LOO), 나머지는 전체 보정 상수. P̂ 경계: FRI 0.5·1·2 (7.5·15·30분).")
    note(ws, 13, "p̂_B 사전분포: f400,air가 0.60 → 0.70으로 갈 때 0.05 → 0.95로 오르는 램프(경계 0.65 = 통합 모델의 B̂ 기준), 무게 n₀ = 2회 시험분. 잠정값입니다.")
    note(ws, 14, "TGA로 못 보는 것: 기상 난연 효과(가연/준불연 차이), 발열 절대량(초기 질량 미기록), DSC 열량(미보정) → 콘·FT-IR·보정 DSC가 보완합니다.")
    return rows


# --------------------------------------------------------------------------------------------- 4 FT-IR
BANDS = [
    ("폴리스티렌 (EPS)", "3060 · 3026", "방향족 C–H 신축", "식별"),
    ("폴리스티렌 (EPS)", "1601 · 1493 · 1452", "방향족 고리", "식별"),
    ("폴리스티렌 (EPS)", "756 · 698", "일치환 벤젠 C–H 면외굽힘", "식별(강함)"),
    ("PIR/PUR (MDI계)", "3300", "N–H 신축 (우레탄)", "PUR 성분"),
    ("PIR/PUR (MDI계)", "2270", "잔류 –N=C=O", "미반응 이소시아네이트"),
    ("PIR/PUR (MDI계)", "1710 – 1730", "C=O (우레탄·이소시아누레이트)", "식별"),
    ("PIR/PUR (MDI계)", "1595", "방향족 C=C (내부 기준 대역)", "지수의 분모"),
    ("PIR/PUR (MDI계)", "1510 – 1530", "아마이드 II (N–H 굽힘 + C–N)", "우레탄 성분"),
    ("PIR/PUR (MDI계)", "1410", "이소시아누레이트 고리 C–N", "이소시아누레이트 지수 A₁₄₁₀/A₁₅₉₅"),
    ("페놀 (PF)", "3300 (넓음)", "O–H 신축", "식별"),
    ("페놀 (PF)", "1600 · 1500", "방향족 고리", "기준 대역"),
    ("페놀 (PF)", "1475 · 1440", "메틸렌 가교 CH₂", "경화 정도"),
    ("페놀 (PF)", "1230 · 1010", "페놀 C–O · 메틸올 C–O", "식별"),
    ("인계 난연제 (TCPP 등)", "1270 – 1290", "P=O 신축", "난연 지수 A_P=O/A_기준"),
    ("인계 난연제 (TCPP 등)", "1030 – 1050 · 980", "P–O–C", "난연제 확인"),
    ("인계 난연제 (TCPP 등)", "760 – 680", "C–Cl", "염소 함유 확인"),
    ("수산화알루미늄 (ATH)", "3620 · 3523 · 3450 · 3380", "O–H 신축", "무기 충전재 지수"),
    ("수산화마그네슘", "3698 (예리)", "O–H 신축", "무기 충전재"),
    ("브롬계 난연제", "(적외선 감도 낮음)", "C–Br", "XRF·SEM/EDS로 Br 확인"),
]
EGA = [
    ("CO₂", "2360 · 2340 · 669"), ("CO", "2180 · 2110 (P·R 가지)"), ("HCN", "713 (Q 가지)"), ("NH₃", "966 · 931"),
    ("–NCO (이소시아네이트)", "2270 – 2250"), ("HCl", "2700 – 3100 (회전선)"), ("H₂O", "3500 – 4000 · 1300 – 1900"),
    ("스티렌", "3080 – 3030 · 1630 · 990 · 910 · 775 · 695"), ("페놀", "3650 (자유 O–H) · 1600 · 1500"),
]


def ftir(wb):
    ws = wb.create_sheet("4_L1_FT-IR")
    title(ws, "④ L1 FT-IR — 고분자·난연제 식별(게이트 검증)과 유해가스(TGA–FTIR) 입력 양식",
          "대표 흡수대는 문헌값입니다. 실험실 기준 스펙트럼으로 확인한 뒤 쓰십시오. 회색 기울임 행은 형식 예시(가상값)입니다.")
    r = 4
    band(ws, r, "A. ATR 식별 대역 (고체: 심재·면재)", 4)
    header(ws, r + 1, ["대상", "대역 (cm⁻¹)", "귀속", "용도"], widths=[22, 24, 30, 30])
    for k, row in enumerate(BANDS, start=r + 2):
        for j, v in enumerate(row, start=1):
            rc(ws, k, j, v, font=F_BOLD if j == 1 else F_BASE)
    r = r + 2 + len(BANDS) + 1
    band(ws, r, "B. ATR 측정 기록 (흡광도: 기준선 보정 후 피크 높이) → 지수", 17)
    cols = ["재료", "부위", "위치", "선언 고분자", "판정 고분자", "A₁₄₁₀", "A₁₅₉₅", "A₂₂₇₀", "A_P=O (1270–1290)", "A_기준", "A_OH (3620)",
            "이소시아누레이트 지수", "잔류 NCO 지수", "난연(P) 지수", "무기충전 지수", "식별 일치"]
    header(ws, r + 1, cols, height=32)
    for j, w in enumerate([16, 8, 8, 12, 12, 8, 8, 8, 10, 8, 9, 12, 10, 10, 10, 9], start=1):
        ws.column_dimensions[L(j)].width = max(ws.column_dimensions[L(j)].width or 0, w)
    first = r + 2
    ex = ["예시(가상값)", "심재", "1", "PIR", "PIR", 0.42, 0.60, 0.02, 0.15, 0.60, None]
    n = 19
    for i in range(n):
        rr = first + i
        is_ex = i == 0
        vals = ex if is_ex else [None] * 11
        for j, v in enumerate(vals, start=1):
            rc(ws, rr, j, v, font=F_EX if is_ex else F_INPUT, fill=FILL_GREY if is_ex else FILL_TODO, align=CENTER if j > 1 else WRAP)
        forms = [f'=IF(OR(F{rr}="",G{rr}=""),"",F{rr}/G{rr})', f'=IF(OR(H{rr}="",G{rr}=""),"",H{rr}/G{rr})',
                 f'=IF(OR(I{rr}="",J{rr}=""),"",I{rr}/J{rr})', f'=IF(OR(K{rr}="",J{rr}=""),"",K{rr}/J{rr})',
                 f'=IF(OR(D{rr}="",E{rr}=""),"",IF(D{rr}=E{rr},"일치","불일치"))']
        for j, fm in enumerate(forms, start=12):
            rc(ws, rr, j, fm, font=F_EX if is_ex else F_BASE, fmt="0.000" if j < 16 else None, align=CENTER, fill=FILL_GREY if is_ex else None)
    last = first + n - 1
    rngB = lambda c: f"${L(c)}${first + 1}:${L(c)}${last}"
    r = last + 2
    band(ws, r, "C. TGA–FTIR 발생가스 기록 (같은 승온속도 5 K/min 권장)", 9)
    header(ws, r + 1, ["재료", "분위기", "가스", "대역 (cm⁻¹)", "최대 발생 온도 (°C)", "적분 면적 (a.u.·s)", "시료 질량 (mg)", "면적/mg", "비고"], height=30)
    firstc = r + 2
    rr = firstc
    rc(ws, rr, 1, "예시(가상값)", font=F_EX, fill=FILL_GREY)
    for j, v in enumerate(["공기", "HCN", "713", 455, 0.84, 7.0], start=2):
        rc(ws, rr, j, v, font=F_EX, fill=FILL_GREY, align=CENTER)
    rc(ws, rr, 8, f'=IF(OR(F{rr}="",G{rr}=""),"",F{rr}/G{rr})', font=F_EX, fill=FILL_GREY, fmt="0.000", align=CENTER)
    rr += 1
    for m in ALL:
        for atm in ("공기", "N₂"):
            for gas, bands in EGA[:6]:
                rc(ws, rr, 1, LABEL[m], font=F_BOLD)
                rc(ws, rr, 2, atm, align=CENTER)
                rc(ws, rr, 3, gas, align=CENTER)
                rc(ws, rr, 4, bands, align=CENTER)
                for j in (5, 6, 7):
                    rc(ws, rr, j, None, font=F_INPUT, fill=FILL_TODO, align=CENTER)
                rc(ws, rr, 8, f'=IF(OR(F{rr}="",G{rr}=""),"",F{rr}/G{rr})', fmt="0.000", align=CENTER)
                rc(ws, rr, 9, None)
                rr += 1
    lastc = rr - 1
    rngC = lambda c: f"${L(c)}${firstc + 1}:${L(c)}${lastc}"
    r = lastc + 2
    band(ws, r, "D. 재료별 요약 (예시 행 제외)", 7)
    header(ws, r + 1, ["재료", "이소시아누레이트 지수", "난연(P) 지수", "식별", "HCN 면적/mg (공기)", "CO 면적/mg (공기)", "게이트 검증"], height=30)
    for k, m in enumerate(ALL):
        rr = r + 2 + k
        lab = LABEL[m]
        rc(ws, rr, 1, lab, font=F_BOLD)
        rc(ws, rr, 2, f'=IFERROR(AVERAGEIFS({rngB(12)},{rngB(1)},$A{rr},{rngB(6)},">0",{rngB(7)},">0"),"")', fmt="0.000", align=CENTER)
        rc(ws, rr, 3, f'=IFERROR(AVERAGEIFS({rngB(14)},{rngB(1)},$A{rr},{rngB(9)},">0",{rngB(10)},">0"),"")', fmt="0.000", align=CENTER)
        rc(ws, rr, 4, f'=IF(COUNTIFS({rngB(1)},$A{rr},{rngB(16)},"불일치")>0,"불일치 있음",IF(COUNTIFS({rngB(1)},$A{rr},{rngB(16)},"일치")>0,"일치",""))', align=CENTER)
        rc(ws, rr, 5, f'=IFERROR(AVERAGEIFS({rngC(8)},{rngC(1)},$A{rr},{rngC(2)},"공기",{rngC(3)},"HCN",{rngC(6)},">0"),"")', fmt="0.000", align=CENTER)
        rc(ws, rr, 6, f'=IFERROR(AVERAGEIFS({rngC(8)},{rngC(1)},$A{rr},{rngC(2)},"공기",{rngC(3)},"CO",{rngC(6)},">0"),"")', fmt="0.000", align=CENTER)
        rc(ws, rr, 7, f'=IF(D{rr}="불일치 있음","제품 확인 → 전이함수 적용 보류",IF(D{rr}="일치",IF(AND(C{rr}<>"",C{rr}>0),"난연제 확인","난연제 미검출 → 가연성으로 다룸"),"(미측정)"))', fill=FILL_KEY)
    note(ws, r + 9, "EGA 대표 대역(문헌값): " + " · ".join(f"{g} {b}" for g, b in EGA) + ". 정량은 기준 가스 보정이 필요하며, 보정 전에는 같은 조건의 재료 간 상대 비교로만 씁니다.")


# --------------------------------------------------------------------------------------------- 5 중규모
def midscale(wb):
    ws = wb.create_sheet("5_L2_중규모")
    title(ws, "⑤ L2 중규모 성능시험 (중공층 100 mm · 외단열 100 mm · 30 kW 20분 · 30분 평가) — 시험별 값",
          "t600 = 3000 mm에서 ΔT ≥ 600 K가 30 s 이상 유지되기 시작한 시각. t_ig·t_sp: coupled_analysis(착화·확산 분해). P·S·R·A: cavity_fire_index. B: integrated_model.")
    cols = ["시험", "재료", "재료명", "등급", "정량평가 사용", "t_ig (s)", "t_sp (s)", "t600 (s)", "ln t600", "(ln t600)²", "FRI = 900/t600",
            "P", "D600 (분)", "S", "R = P×S", "급발 A", "자립 B (1/0)", "H600 (K·min)", "최고온도 3000 mm (°C)"]
    header(ws, 4, cols, widths=[10, 9, 16, 6, 9, 8, 8, 8, 8, 9, 9, 5, 8, 5, 7, 7, 8, 9, 10], height=40)
    tests = pd.read_csv(BASE.parent / "facade_fra" / "inputs" / "test_inputs.csv")
    idx = pd.read_csv(ROOT / "cavity_fire_index" / "results" / "index_3000mm.csv").set_index("test")
    rm = pd.read_csv(ROOT / "cavity_fire_index" / "results" / "risk_model.csv").set_index("test")
    ig = pd.read_csv(ROOT / "coupled_analysis" / "results" / "ignition_spread_inverse.csv").set_index("test")
    r = 5
    for t in tests.itertuples():
        rc(ws, r, 1, t.test, font=F_BOLD)
        rc(ws, r, 2, t.material, align=CENTER)
        rc(ws, r, 3, t.label)
        rc(ws, r, 4, t.grade, align=CENTER)
        rc(ws, r, 5, "예" if t.used else "아니오", font=F_INPUT, align=CENTER)
        tig = ig.loc[t.test, "t_ig_s"] if t.test in ig.index else None
        tsp = ig.loc[t.test, "t_spread_s"] if t.test in ig.index else None
        rc(ws, r, 6, None if tig is None or pd.isna(tig) else float(tig), font=F_INPUT, fmt="0", align=CENTER)
        rc(ws, r, 7, None if tsp is None or pd.isna(tsp) else float(tsp), font=F_INPUT, fmt="0", align=CENTER)
        rc(ws, r, 8, None if pd.isna(t.t600_s) else float(t.t600_s), font=F_INPUT, fmt="0", align=CENTER)
        rc(ws, r, 9, f'=IF(H{r}="","",LN(H{r}))', fmt="0.0000", align=CENTER)
        rc(ws, r, 10, f'=IF(H{r}="","",LN(H{r})^2)', fmt="0.0000", align=CENTER)
        rc(ws, r, 11, f'=IF(H{r}="","",t_ref/H{r})', fmt="0.00", align=CENTER)
        if t.test in rm.index:
            q = rm.loc[t.test]
            for j, v in ((12, int(q.P)), (13, float(q.D600_min)), (14, int(q.S)), (15, int(q.R)), (16, "A" if q.abrupt_flag == "A" else "")):
                rc(ws, r, j, v, font=F_INPUT, fmt="0.0" if j == 13 else None, align=CENTER)
        else:
            for j in range(12, 17):
                rc(ws, r, j, None)
        rc(ws, r, 17, int(t.flag_B), font=F_INPUT, align=CENTER)
        rc(ws, r, 18, float(idx.loc[t.test, "excess_600K_Kmin"]), font=F_INPUT, fmt="0", align=CENTER)
        rc(ws, r, 19, float(idx.loc[t.test, "peak_C"]), font=F_INPUT, fmt="0", align=CENTER)
        r += 1
    last = r - 1
    for name, col in (("l2_mat", 2), ("l2_use", 5), ("l2_t600", 8), ("l2_ln", 9), ("l2_lnsq", 10), ("l2_B", 17)):
        define(wb, name, "5_L2_중규모", f"${L(col)}$5:${L(col)}${last}")
    note(ws, last + 2, "PIR(PUR) 가연성은 30 s 만에 전면 착화(착화·확산 분해 불가), EPS 준불연 #2는 30분 안에 600 K 미도달(층간 확산 없음으로 계산). "
                       "EPS 가연성은 기록 4분 → 정량평가 제외(재시험 필요).")


# --------------------------------------------------------------------------------------------- 6 전이함수
TRANSFER = [
    ("T1", "TGA → 확산", "t̂600 = t_ig,sys + C_sys / r_max,air", "준불연 · 탄화형 (f400,air ≥ 0.5)", "보정·교차검증(제품 제외 오차 −14 – +29 %)", "D2"),
    ("T2", "TGA → 지속", "p̂_B ~ Beta(n₀m, n₀(1−m)), m = 램프(f400,air; 0.60 → 0.70)", "준불연 · 탄화형", "잠정(제품 3종, B̂ 경계 0.65)", "D4"),
    ("T3", "TGA → 급발", "f400,air < 0.5 → 용융형: 확산식 미적용, A 예상", "전체", "관측 일치(EPS 준불연 A 2/3)", "D5"),
    ("T4", "콘 → 게이트", "준불연 기준 미달 → 가연성: TGA 미적용, 중규모 필수", "전체", "규정 기준 · 콘 자료 필요", "G"),
    ("T5", "콘 → 확산", "M = Q̇″^(1/3)/TRP → t̂600 = 1079 s (M_ref/M)², b > 0 → 가속", "준불연", "제안(콘 자료로 보정)", "D2"),
    ("T6", "콘 → 발열", "THR_600 ↔ H600 (3000 mm 초과 열량)", "전체", "제안", "D3"),
    ("T7", "FT-IR → 게이트", "식별 일치 + 난연(P) 지수 > 0 → 준불연 처리 확인; 이소시아누레이트 지수 ↔ T_onset·f400", "전체", "제안", "G"),
    ("T8", "FT-IR·콘 → 유해", "HCN·CO 발생량 → FDS 수율 → 위층 FED", "전체", "제안", "D6"),
    ("S", "중규모 → 건물", "τ_F = s · t600 · h / 3 m (s: 실물/시험 비)", "전체", "가정 LogU(0.25, 1) → L3·실대형으로 대체", "D2"),
]


def transfer(wb, tga_rows):
    ws = wb.create_sheet("6_전이함수")
    title(ws, "⑥ 전이함수 — 상수, 시험 반복 편차, TGA 예측 편차 (수식)",
          "소형시료 지표를 중규모 등가값(t600, p_B)으로 바꾸는 식과 상수입니다. σ_rep·τ·제품 간 SD는 5_L2_중규모의 값에서 수식으로 계산됩니다.")
    r = 4
    band(ws, r, "상수", 6)
    header(ws, r + 1, ["이름", "항목", "값", "단위", "출처"], widths=[14, 44, 12, 12, 52])
    consts = [("t_ig_sys", "시스템 착화 지연 (준불연 6회 중앙값)", D["constants"]["t_ig_sys"], "s", "integrated_model system_constants"),
              ("C_sys", "확산 상수 t_sp × r_max,air (6회 평균)", D["constants"]["c_sys"], "s·%/°C", "integrated_model (변동계수 12.75 %)"),
              ("t_ref", "FRI 기준시간", 900, "s", "cavity_fire_index RISK_MODEL"),
              ("f_lo", "p_B 램프 시작 f400,air", D["constants"]["f_ramp"][0], "–", "B̂ 경계 0.65 − 0.05"),
              ("f_hi", "p_B 램프 끝 f400,air", D["constants"]["f_ramp"][1], "–", "B̂ 경계 0.65 + 0.05"),
              ("n0", "TGA p_B 사전분포 무게", D["constants"]["n0"], "회", "잠정"),
              ("f_melt", "용융형 경계 f400,air", D["constants"]["f_melt"], "–", "integrated_model 식 4"),
              ("tau_min", "τ 하한", 0.05, "–", "자료 3종 → 과소추정 방지"),
              ("s_lo", "s 범위 하한", D["constants"]["s_range"][0], "–", "facade_fra PARAMS"),
              ("s_hi", "s 범위 상한", D["constants"]["s_range"][1], "–", "facade_fra PARAMS"),
              ("crit_conf", "판정 신뢰수준", D["constants"]["crit_conf"], "–", "facade_fra 예시")]
    for k, (name, item, val, unit, src) in enumerate(consts, start=r + 2):
        rc(ws, k, 1, name, font=F_BOLD)
        rc(ws, k, 2, item)
        rc(ws, k, 3, val, font=F_INPUT, fill=FILL_TODO, fmt="0.0000" if isinstance(val, float) and val < 10 else "0.0", align=CENTER)
        rc(ws, k, 4, unit, align=CENTER)
        rc(ws, k, 5, src, font=F_NOTE)
        define(wb, name, "6_전이함수", A(3, k))
    r = r + 2 + len(consts) + 1
    band(ws, r, "제품 제외(LOO) 보정 상수와 시험 통계 — 준불연 탄화형", 12)
    cols = ["제품", "t_ig,sys LOO (s)", "C_sys LOO", "r_max,air", "t̂600 LOO (s)", "시험 수 n", "ȳ = 평균 ln t600", "Σ (ln t600)²",
            "SS = Σ(ln t − ȳ)²", "잔차 r = ȳ − ln t̂", "r² − σ_rep²/n", "시험 기하평균 (s)", "검산: t_ig + C/r_max (s)"]
    header(ws, r + 1, cols, height=32)
    loo = pd.read_csv(ROOT / "integrated_model" / "results" / "loo_validation.csv").groupby("material").first()
    first = r + 2
    for k, m in enumerate(CHAR):
        rr = first + k
        rc(ws, rr, 1, m, font=F_BOLD)
        rc(ws, rr, 2, float(loo.loc[m, "t_ig_sys_loo_s"]), font=F_INPUT, fmt="0.0", align=CENTER)
        rc(ws, rr, 3, float(loo.loc[m, "C_sys_loo"]), font=F_INPUT, fmt="0.00", align=CENTER)
        rc(ws, rr, 4, f"='3_L1_TGA'!F{tga_rows[m]}", font=F_LINK, fmt="0.0000", align=CENTER)
        rc(ws, rr, 5, float(loo.loc[m, "t600_hat_loo_s"]), font=F_INPUT, fmt="0.00", align=CENTER, fill=FILL_KEY)
        rc(ws, rr, 13, f"=B{rr}+C{rr}/D{rr}", fmt="0.00", align=CENTER)
        crit = f'l2_mat,A{rr},l2_use,"예",l2_t600,">0"'
        rc(ws, rr, 6, f"=COUNTIFS({crit})", align=CENTER)
        rc(ws, rr, 7, f"=SUMIFS(l2_ln,{crit})/F{rr}", fmt="0.00000", align=CENTER)
        rc(ws, rr, 8, f"=SUMIFS(l2_lnsq,{crit})", fmt="0.00000", align=CENTER)
        rc(ws, rr, 9, f"=H{rr}-F{rr}*G{rr}^2", fmt="0.000000", align=CENTER)
        rc(ws, rr, 10, f"=G{rr}-LN(E{rr})", fmt="0.0000", align=CENTER)
        rc(ws, rr, 11, f"=J{rr}^2-sigma_rep^2/F{rr}", fmt="0.00000", align=CENTER)
        rc(ws, rr, 12, f"=EXP(G{rr})", fmt="0.0", align=CENTER)
    last = first + len(CHAR) - 1
    note(ws, last + 1, "t̂600 LOO: integrated_model이 해당 제품을 빼고 t_ig,sys·C_sys를 다시 맞춰 낸 예측값(원값). 마지막 열은 반올림 상수로 다시 계산한 검산값입니다.")
    define(wb, "loo_key", "6_전이함수", f"$A${first}:$A${last}")
    define(wb, "loo_t600", "6_전이함수", f"$E${first}:$E${last}")
    r = last + 3
    band(ws, r, "통계 (수식)", 6)
    stats = [("sigma_rep", "시험 반복 편차 σ_rep (ln t600, 반복쌍 합동)", f"=SQRT(SUM(I{first}:I{last})/(SUM(F{first}:F{last})-COUNT(F{first}:F{last})))"),
             ("tau2_raw", "TGA 예측 편차² (제품 평균) = 평균(r² − σ_rep²/n)", f"=AVERAGE(K{first}:K{last})"),
             ("tau_tga", "TGA 예측 편차 τ (ln t600)", "=SQRT(MAX(tau2_raw,tau_min^2))"),
             ("sd_between", "준불연 제품 간 SD (ln t600, 등급만 아는 경우의 사전분포)", f"=STDEV(G{first}:G{last})")]
    for k, (name, item, f) in enumerate(stats, start=r + 1):
        rc(ws, k, 1, name, font=F_BOLD)
        rc(ws, k, 2, item)
        rc(ws, k, 3, f, fmt="0.0000", align=CENTER, fill=FILL_KEY)
        rc(ws, k, 4, "–", align=CENTER)
        define(wb, name, "6_전이함수", A(3, k))
    r = r + len(stats) + 2
    band(ws, r, "전이함수 목록", 6)
    header(ws, r + 1, ["번호", "연결", "식", "적용 범위", "상태", "위험 차원"])
    for k, row in enumerate(TRANSFER, start=r + 2):
        for j, v in enumerate(row, start=1):
            rc(ws, k, j, v, font=F_BOLD if j == 1 else F_BASE, align=CENTER if j in (1, 6) else WRAP)
        ws.row_dimensions[k].height = 30
    return dict(first=first, rows={m: first + k for k, m in enumerate(CHAR)})


# --------------------------------------------------------------------------------------------- 7 통합(베이즈)
def bayes(wb, tga_rows, tr):
    ws = wb.create_sheet("7_통합(베이즈)")
    title(ws, "⑦ 통합 — TGA 사전분포 × 중규모 시험 (베이즈 결합)과 척도 간 정합성",
          "ln t600: 정규–정규 결합(사전 N(ln t̂_TGA, τ²), 시험 n회 N(μ, σ_rep²)). p_B: 베타–이항 결합. 정합성: |z| ≤ 2, 예측확률 ≥ 0.05.")
    r = 4
    band(ws, r, "A. 확산 (t600) — 준불연 탄화형", 18)
    cols = ["제품", "μ₀ = ln t̂_TGA", "τ", "ȳ (시험)", "n", "σ_rep", "사전 정밀도 1/τ²", "자료 정밀도 n/σ²", "TGA 무게 w", "μ_post",
            "σ_post", "t600 TGA (s)", "t600 시험 (s)", "t600 통합 (s)", "90 % 하한", "90 % 상한", "z = (ȳ − μ₀)/√(τ² + σ²/n)", "정합"]
    header(ws, r + 1, cols, widths=[10, 10, 8, 10, 5, 8, 10, 10, 9, 9, 8, 10, 10, 10, 9, 9, 14, 8], height=40)
    rows_a = {}
    for k, m in enumerate(CHAR):
        rr = r + 2 + k
        rows_a[m] = rr
        t = tr["rows"][m]
        rc(ws, rr, 1, m, font=F_BOLD)
        rc(ws, rr, 2, f"=LN('6_전이함수'!E{t})", font=F_LINK, fmt="0.0000", align=CENTER)
        rc(ws, rr, 3, "=tau_tga", font=F_LINK, fmt="0.0000", align=CENTER)
        rc(ws, rr, 4, f"='6_전이함수'!G{t}", font=F_LINK, fmt="0.0000", align=CENTER)
        rc(ws, rr, 5, f"='6_전이함수'!F{t}", font=F_LINK, align=CENTER)
        rc(ws, rr, 6, "=sigma_rep", font=F_LINK, fmt="0.0000", align=CENTER)
        rc(ws, rr, 7, f"=1/C{rr}^2", fmt="0.0", align=CENTER)
        rc(ws, rr, 8, f"=E{rr}/F{rr}^2", fmt="0.0", align=CENTER)
        rc(ws, rr, 9, f"=G{rr}/(G{rr}+H{rr})", fmt="0.000", align=CENTER)
        rc(ws, rr, 10, f"=I{rr}*B{rr}+(1-I{rr})*D{rr}", fmt="0.0000", align=CENTER)
        rc(ws, rr, 11, f"=1/SQRT(G{rr}+H{rr})", fmt="0.0000", align=CENTER)
        rc(ws, rr, 12, f"=EXP(B{rr})", fmt="0", align=CENTER)
        rc(ws, rr, 13, f"=EXP(D{rr})", fmt="0", align=CENTER)
        rc(ws, rr, 14, f"=EXP(J{rr})", fmt="0", align=CENTER, fill=FILL_KEY)
        rc(ws, rr, 15, f"=EXP(J{rr}-1.645*K{rr})", fmt="0", align=CENTER)
        rc(ws, rr, 16, f"=EXP(J{rr}+1.645*K{rr})", fmt="0", align=CENTER)
        rc(ws, rr, 17, f"=(D{rr}-B{rr})/SQRT(C{rr}^2+F{rr}^2/E{rr})", fmt="0.00", align=CENTER, fill=FILL_KEY)
        rc(ws, rr, 18, f'=IF(ABS(Q{rr})<=2,"정합","불일치")', align=CENTER)
    r = r + 2 + len(CHAR) + 1
    band(ws, r, "B. 게이트 확인 — TGA 전이함수를 게이트 없이 쓰면 (가연성 · 용융형)", 9)
    header(ws, r + 1, ["제품", "등급", "재료형", "t̂600 전체 보정 (s)", "시험 기하평균 (s, 도달 시험)", "도달/시험 수", "z", "판정", "해석"], height=32)
    ws.merge_cells(start_row=r + 1, start_column=9, end_row=r + 1, end_column=18)
    rows_b = {}
    for k, m in enumerate(["PIR_C", "EPS_SC"]):
        rr = r + 2 + k
        rows_b[m] = rr
        tg = tga_rows[m]
        crit = f'l2_mat,"{m}",l2_use,"예",l2_t600,">0"'
        rc(ws, rr, 1, m, font=F_BOLD)
        rc(ws, rr, 2, f"='3_L1_TGA'!C{tg}", font=F_LINK, align=CENTER)
        rc(ws, rr, 3, f"='3_L1_TGA'!M{tg}", font=F_LINK, align=CENTER)
        rc(ws, rr, 4, f"='3_L1_TGA'!O{tg}", font=F_LINK, fmt="0", align=CENTER)
        rc(ws, rr, 5, f"=EXP(SUMIFS(l2_ln,{crit})/COUNTIFS({crit}))", fmt="0", align=CENTER)
        rc(ws, rr, 6, f'=COUNTIFS({crit})&" / "&COUNTIFS(l2_mat,"{m}",l2_use,"예")', align=CENTER)
        rc(ws, rr, 7, f"=(LN(E{rr})-LN(D{rr}))/SQRT(tau_tga^2+sigma_rep^2/COUNTIFS({crit}))", fmt="0.00", align=CENTER, fill=FILL_KEY)
        rc(ws, rr, 8, f'=IF(ABS(G{rr})<=2,"정합","불일치")', align=CENTER)
        rc(ws, rr, 9, {"PIR_C": "가연성: 기상 난연 차이를 TGA가 못 봄 → 위험을 6배 낮게 추정(게이트 필요)",
                       "EPS_SC": "용융형: 녹아 물러난 뒤 순간 연소 → 확산식이 위험을 높게 추정(미도달 1회 포함 시 더 큼)"}[m])
        ws.merge_cells(start_row=rr, start_column=9, end_row=rr, end_column=18)
        ws.row_dimensions[rr].height = 22
    r = r + 2 + 2 + 1
    band(ws, r, "C. 지속 (p_B) — 베타–이항 결합", 14)
    cols = ["제품", "f400,air", "m (램프)", "a₀", "b₀", "B 횟수 k", "시험 수 n", "a = a₀ + k", "b = b₀ + n − k", "p_B 통합", "p_B 시험만 (Jeffreys)",
            "p_B TGA만", "예측확률 P(k | 사전)", "정합"]
    header(ws, r + 1, cols, height=40)
    rows_c = {}
    for k, m in enumerate(["EPS_SC"] + CHAR):
        rr = r + 2 + k
        rows_c[m] = rr
        tg = tga_rows[m]
        rc(ws, rr, 1, m, font=F_BOLD)
        rc(ws, rr, 2, f"='3_L1_TGA'!H{tg}", font=F_LINK, fmt="0.000", align=CENTER)
        rc(ws, rr, 3, f"='3_L1_TGA'!S{tg}", font=F_LINK, fmt="0.000", align=CENTER)
        rc(ws, rr, 4, f"='3_L1_TGA'!T{tg}", font=F_LINK, fmt="0.000", align=CENTER)
        rc(ws, rr, 5, f"='3_L1_TGA'!U{tg}", font=F_LINK, fmt="0.000", align=CENTER)
        rc(ws, rr, 6, f'=SUMIFS(l2_B,l2_mat,A{rr},l2_use,"예")', align=CENTER)
        rc(ws, rr, 7, f'=COUNTIFS(l2_mat,A{rr},l2_use,"예")', align=CENTER)
        rc(ws, rr, 8, f"=D{rr}+F{rr}", fmt="0.000", align=CENTER)
        rc(ws, rr, 9, f"=E{rr}+G{rr}-F{rr}", fmt="0.000", align=CENTER)
        rc(ws, rr, 10, f"=H{rr}/(H{rr}+I{rr})", fmt="0.000", align=CENTER, fill=FILL_KEY)
        rc(ws, rr, 11, f"=(F{rr}+0.5)/(G{rr}+1)", fmt="0.000", align=CENTER)
        rc(ws, rr, 12, f"=D{rr}/(D{rr}+E{rr})", fmt="0.000", align=CENTER)
        rc(ws, rr, 13, f"=COMBIN(G{rr},F{rr})*EXP(GAMMALN(F{rr}+D{rr})+GAMMALN(G{rr}-F{rr}+E{rr})-GAMMALN(G{rr}+D{rr}+E{rr})"
                       f"-GAMMALN(D{rr})-GAMMALN(E{rr})+GAMMALN(D{rr}+E{rr}))", fmt="0.000", align=CENTER, fill=FILL_KEY)
        rc(ws, rr, 14, f'=IF(M{rr}>=0.05,"정합","불일치")', align=CENTER)
    note(ws, r + 7, "가연성 PIR(PUR)은 버너 소화 전에 연료가 소진되어 B 판정 대상이 아닙니다(p_B는 위험도에 영향 없음).")
    return rows_a, rows_b, rows_c


# --------------------------------------------------------------------------------------------- 8 시뮬레이션
def simulation(wb, tr):
    ws = wb.create_sheet("8_L3_시뮬레이션")
    title(ws, "⑧ L3 시뮬레이션 (FDS/PyroSim) — 중규모 재현 검증과 건물 외벽 시나리오 → s 추정 (입력 양식)",
          "1단계: 중규모 시험 재현으로 모델을 검증합니다. 2단계: 검증된 모델로 건물 외벽(3–6개 층)을 계산해 층별 도달시각에서 s를 구합니다. 회색 기울임 = 예시(가상값).")
    r = 4
    band(ws, r, "검증 허용 기준 (예시값 — 사전협의로 확정)", 4)
    for k, (name, item, val, fmt) in enumerate([("tol_t600", "t600 상대오차 허용", 0.30, "0%"), ("tol_peak", "3000 mm 최고온도 오차 허용 (K)", 100, "0"),
                                                 ("h_storey", "층고 (m)", 3.0, "0.0")], start=r + 1):
        rc(ws, k, 1, name, font=F_BOLD)
        rc(ws, k, 2, item)
        rc(ws, k, 3, val, font=F_INPUT, fill=FILL_TODO, fmt=fmt, align=CENTER)
        define(wb, name, "8_L3_시뮬레이션", A(3, k))
    r = r + 5
    band(ws, r, "A. 1단계 — 중규모 시험 재현", 8)
    header(ws, r + 1, ["재료", "시험 t600 (s)", "FDS t600 (s)", "상대오차", "시험 최고온도 3000 mm (°C)", "FDS 최고온도 (°C)", "온도 오차 (K)", "판정"],
           widths=[18, 12, 12, 10, 14, 12, 12, 14], height=32)
    peak = {p["material"]: p.get("peak_med") for p in D["profile"]}
    for k, m in enumerate(MATS):
        rr = r + 2 + k
        crit = f'l2_mat,"{m}",l2_use,"예",l2_t600,">0"'
        rc(ws, rr, 1, LABEL[m], font=F_BOLD)
        rc(ws, rr, 2, f"=EXP(SUMIFS(l2_ln,{crit})/COUNTIFS({crit}))", font=F_LINK, fmt="0", align=CENTER)
        rc(ws, rr, 3, None, font=F_INPUT, fill=FILL_TODO, align=CENTER)
        rc(ws, rr, 4, f'=IF(C{rr}="","",(C{rr}-B{rr})/B{rr})', fmt="0%", align=CENTER)
        rc(ws, rr, 5, float(peak[m]), font=F_INPUT, fmt="0", align=CENTER)
        rc(ws, rr, 6, None, font=F_INPUT, fill=FILL_TODO, align=CENTER)
        rc(ws, rr, 7, f'=IF(F{rr}="","",F{rr}-E{rr})', fmt="0", align=CENTER)
        rc(ws, rr, 8, f'=IF(OR(C{rr}="",F{rr}=""),"",IF(AND(ABS(D{rr})<=tol_t600,ABS(G{rr})<=tol_peak),"검증 통과","보완 필요"))', align=CENTER, fill=FILL_KEY)
    r = r + 2 + len(MATS) + 1
    band(ws, r, "B. 2단계 — 건물 외벽 시나리오: 발화층 위 j층 도달시각 → 층별 s_j = ΔT_j / (t600 · h/3 m) → s 추정 → 판정 예측(M0)", 20)
    cols = ["재료", "시나리오", "시험 t600 (s)"] + [f"T{j} (s)" for j in range(1, 7)] + [f"s{j}" for j in range(1, 7)] + \
           ["s 추정 (중앙값)", "충족 하한 s_meet", "불충족 상한 s_fail", "판정 예측 (M0)"]
    header(ws, r + 1, cols, height=34)
    for j in range(4, 20):
        ws.column_dimensions[L(j)].width = 8
    thr = {(t["material"], t["mitigation"]): t for t in D["voi_thresholds"]}
    rr = r + 2
    entries = [("예시(가상값)", "IE1 창 분출", 1000.0, [480, 1000, 1500, 2050, None, None], None)] + \
              [(m, sc, None, [None] * 6, m) for m in MATS for sc in ("IE1 창 분출", "IE2 외부 화원")]
    for name, sc, t600v, ts, m in entries:
        is_ex = m is None
        fnt = F_EX if is_ex else F_INPUT
        rc(ws, rr, 1, name if is_ex else LABEL[m], font=F_EX if is_ex else F_BOLD, fill=FILL_GREY if is_ex else None)
        rc(ws, rr, 2, sc, align=CENTER, fill=FILL_GREY if is_ex else None)
        if is_ex:
            rc(ws, rr, 3, t600v, font=F_EX, fill=FILL_GREY, fmt="0", align=CENTER)
        else:
            crit = f'l2_mat,"{m}",l2_use,"예",l2_t600,">0"'
            rc(ws, rr, 3, f"=EXP(SUMIFS(l2_ln,{crit})/COUNTIFS({crit}))", font=F_LINK, fmt="0", align=CENTER)
        for j, v in enumerate(ts, start=4):
            rc(ws, rr, j, v, font=fnt, fill=FILL_GREY if is_ex else FILL_TODO, align=CENTER)
        for j in range(6):
            tj, tprev = f"{L(4 + j)}{rr}", (f"{L(3 + j)}{rr}" if j > 0 else None)
            if j == 0:
                f = f'=IF({tj}="","",{tj}/($C{rr}*h_storey/3))'
            else:
                f = f'=IF(OR({tj}="",{tprev}=""),"",({tj}-{tprev})/($C{rr}*h_storey/3))'
            rc(ws, rr, 10 + j, f, fmt="0.00", align=CENTER, fill=FILL_GREY if is_ex else None)
        rc(ws, rr, 16, f'=IFERROR(MEDIAN(J{rr}:O{rr}),"")', fmt="0.00", align=CENTER, fill=FILL_KEY)
        if is_ex:
            sm, sf = 0.69, 0.38
        else:
            t = thr[(m, "M0")]
            sm, sf = t["s_meet"], t["s_fail"]
        rc(ws, rr, 17, sm, font=F_EX if is_ex else F_INPUT, fmt="0.00", align=CENTER, fill=FILL_GREY if is_ex else None)
        rc(ws, rr, 18, sf, font=F_EX if is_ex else F_INPUT, fmt="0.00", align=CENTER, fill=FILL_GREY if is_ex else None)
        rc(ws, rr, 19, f'=IF(P{rr}="","",IF(AND(Q{rr}<>"",P{rr}>=Q{rr}),"충족",IF(AND(R{rr}<>"",P{rr}<=R{rr}),"불충족","판단 유보")))',
           align=CENTER, fill=FILL_KEY)
        rr += 1
    note(ws, rr + 1, "s_meet·s_fail: 9_L3_정보가치(통합 L1+L2, 스프링클러 설치, 대책 없음). 빈칸 = 범위(0.25–1) 안에서 해당 경계 없음. "
                     "층별 s_j가 크게 다르면 '층마다 t600' 가정 대신 층별 시간을 위험도 모델에 직접 넣습니다.")
    note(ws, rr + 2, "추가 산출(입력란 별도 기록 권장): 위층 창 열류속(kW/m²), 창 파손 시각, 위층 세대 FED(CO·HCN; 4_L1_FT-IR·콘 CO 수율로 수율 설정).")


# --------------------------------------------------------------------------------------------- 9 정보가치
def voi(wb):
    ws = wb.create_sheet("9_L3_정보가치")
    title(ws, "⑨ 시뮬레이션(L3)의 정보가치 — s를 알면 판정이 어떻게 갈리는가 (통합 L1+L2, 스프링클러 설치)",
          f"s를 고정하고 나머지 불확실성으로 몬테카를로(인식론적 {D['voi_outer']} × 우연적 {D['voi_inner']})를 돌린 기준 충족 비율(값). "
          "P(…)는 현재 s 사전분포 LogU(0.25, 1)에서의 확률(수식).")
    grid = D["s_grid"]
    vo = pd.DataFrame(D["voi"])
    r = 4
    for g, gk in (("M0", "대책 없음"), ("M1", "차단재 + 층간 배리어")):
        band(ws, r, f"{g} {gk} — 기준 충족 비율", 2 + len(grid))
        header(ws, r + 1, ["재료", "s →"] + [f"{s:.2f}" for s in grid], widths=[18, 6] + [8] * len(grid))
        for k, m in enumerate(MATS):
            rr = r + 2 + k
            rc(ws, rr, 1, LABEL[m], font=F_BOLD)
            rc(ws, rr, 2, "", align=CENTER)
            for j, s in enumerate(grid):
                v = float(vo[(vo.material == m) & (vo.mitigation == g) & (abs(vo.s - s) < 1e-9)].conf_meets_F2.iloc[0])
                rc(ws, rr, 3 + j, v, font=F_INPUT, fmt="0.00", align=CENTER,
                   fill=FILL_HAVE if v >= 0.9 else (FILL_MISS if v <= 0.1 else FILL_PART))
        r = r + 2 + len(MATS) + 1
    band(ws, r, "판정 경계와 L3 이후 판정이 확정될 확률", 8)
    header(ws, r + 1, ["재료", "대책", "s_meet (이상이면 충족)", "s_fail (이하면 불충족)", "P(충족 확정)", "P(불충족 확정)", "P(판정 확정)", "해석"],
           widths=[18, 6, 12, 12, 10, 10, 10, 50], height=32)
    rr = r + 2
    for t in D["voi_thresholds"]:
        rc(ws, rr, 1, LABEL[t["material"]], font=F_BOLD)
        rc(ws, rr, 2, t["mitigation"], align=CENTER)
        rc(ws, rr, 3, t["s_meet"], font=F_INPUT, fmt="0.00", align=CENTER)
        rc(ws, rr, 4, t["s_fail"], font=F_INPUT, fmt="0.00", align=CENTER)
        rc(ws, rr, 5, f'=IF(C{rr}="",0,MIN(1,MAX(0,LN(s_hi/C{rr})/LN(s_hi/s_lo))))', fmt="0%", align=CENTER)
        rc(ws, rr, 6, f'=IF(D{rr}="",0,MIN(1,MAX(0,1-LN(s_hi/D{rr})/LN(s_hi/s_lo))))', fmt="0%", align=CENTER)
        rc(ws, rr, 7, f"=MIN(1,E{rr}+F{rr})", fmt="0%", align=CENTER, fill=FILL_KEY)
        m, g = t["material"], t["mitigation"]
        if m == "PIR_C":
            txt = "s와 무관하게 불충족(가연성)"
        elif g == "M1":
            txt = "s와 무관하게 충족 → 대책 적용 시 L3 불필요"
        elif t["s_meet"] is None:
            txt = f"s ≤ 1 범위에서 충족 불가 · s ≤ {t['s_fail']:.2f}이면 불충족"
        else:
            txt = f"s ≥ {t['s_meet']:.2f}이면 충족, s ≤ {t['s_fail']:.2f}이면 불충족, 그 사이는 유보"
        rc(ws, rr, 8, txt)
        rr += 1


# --------------------------------------------------------------------------------------------- 10 결과
def results(wb):
    ws = wb.create_sheet("10_위험도결과")
    title(ws, "⑩ 증거 수준별 건물 위험도 (몬테카를로, 스프링클러 설치 15층 공동주택, 값)",
          f"multiscale.py: 인식론적 {D['n_outer']} × 우연적 {D['n_inner']}. L2(중규모만)는 facade_fra/qfra.py와 같은 난수로 결과가 일치합니다(상대차 {D['l2_check_rel_err']:.1e}).")
    cols = ["증거 수준", "재료", "대책", "R_S 점추정", "R_S 평균", "R_S 5 %", "R_S 95 %", "F(n≥2) 평균", "재료 인자 MF (중앙값)", "MF 5 %", "MF 95 %",
            "기준 충족 비율", "판정 (신뢰 90 %)"]
    header(ws, 4, cols, widths=[18, 18, 6, 11, 11, 11, 11, 11, 11, 9, 9, 10, 12], height=32)
    r = 5
    order = ["L0", "L1", "L1u", "L2", "L1+L2"]
    risk = pd.DataFrame(D["risk"])
    for g in ("M0", "M1"):
        for mode in order:
            for m in ["NC"] + MATS:
                q = risk[(risk["mode"] == mode) & (risk.material == m) & (risk.mitigation == g)].iloc[0]
                vals = [D["mode_ko"][mode], LABEL[m], g, q.R_S_point, q.R_S_mean, q.R_S_p05, q.R_S_p95, q.F_n2_mean, q.MF_p50, q.MF_p05, q.MF_p95,
                        q.conf_meets_F2, q.verdict]
                for j, v in enumerate(vals, start=1):
                    fmt = SCI if j in (4, 5, 6, 7, 8) else ("0.0" if j in (9, 10, 11) else ("0.0%" if j == 12 else None))
                    fill = None
                    if j == 13:
                        fill = FILL_HAVE if v == "충족" else (FILL_MISS if v == "불충족" else FILL_PART)
                    rc(ws, r, j, float(v) if isinstance(v, (int, float)) and j > 3 else v, font=F_INPUT if j > 3 else F_BASE,
                       fmt=fmt, align=CENTER if j > 2 else WRAP, fill=fill)
                r += 1
    r += 1
    band(ws, r, "재료 인자 분해 (통합 L1+L2, 몬테카를로 평균): MF = R_S / R_S(불연 기준안), 증가분 = 확산 몫 + 자립 연소(지속) 몫", 7)
    header(ws, r + 1, ["재료", "대책", "R_S 평균", "R_S (p_B = 0)", "MF", "확산 몫", "지속 몫"])
    rr = r + 2
    for d in D["decomposition"]:
        vals = [LABEL[d["material"]], d["mitigation"], d["R_S_mean"], d["R_S_mean_noB"], d["MF_mean"], d["spread_part"], d["persist_part"]]
        for j, v in enumerate(vals, start=1):
            rc(ws, rr, j, v, font=F_INPUT if j > 2 else F_BASE, fmt=[None, None, SCI, SCI, "0.0", "0%", "0%"][j - 1], align=CENTER if j > 1 else WRAP)
        rr += 1
    r = rr + 1
    band(ws, r, "재료 평균 t600의 불확실성 (90 % 구간 상한/하한 비) — 증거가 쌓일수록 좁아짐", 6)
    header(ws, r + 1, ["제품", "등급만 (L0)", "TGA (L1)", "중규모 (L2)", "통합 (L1+L2)", "비고"])
    rr = r + 2
    for w in D["width"]:
        vals = [LABEL[w["material"]], w["ratio90_L0"], w["ratio90_L1"], w["ratio90_L2"], w["ratio90_L1+L2"], "건물 위험도 구간은 s·빈도가 지배(별도)"]
        for j, v in enumerate(vals, start=1):
            rc(ws, rr, j, v, font=F_INPUT if 1 < j < 6 else F_BASE, fmt="0.00" if 1 < j < 6 else None, align=CENTER if 1 < j < 6 else WRAP)
        rr += 1
    ws.freeze_panes = "C5"


# --------------------------------------------------------------------------------------------- 11 EFFECT
def effect(wb):
    ws = wb.create_sheet("11_EFFECT비교")
    title(ws, "⑪ EFFECT(Lamont & Ingolfsson 2018) Tier 2를 이 시험 재료에 적용하면 — 범주형 재료 평가의 한계",
          "가정: 불연 외장재, 공기층 100 mm(> 75 mm), 15층(30–50 m) 주거 all-out, 스프링클러 설치, 외벽 하부 발화원 있음.")
    cols = ["재료", "EFFECT 단열재 계열", "연료(Fig. 33·34)", "발화원", "수직 연결", "가능성", "결과", "EFFECT 판정", "시험 t600 (s)", "R = P×S (최대)",
            "R_S 중규모 기반 (평균)", "재료 인자 MF"]
    header(ws, 4, cols, widths=[16, 18, 9, 14, 14, 8, 18, 12, 10, 10, 12, 10], height=32)
    prof = {p["material"]: p for p in D["profile"]}
    risk = pd.DataFrame(D["risk"])
    for k, e in enumerate(D["effect"], start=5):
        m = e["material"]
        p = prof[m]
        vals = [LABEL[m], e["family"], e["fuel"], e["ignition"], e["vertical"], e["likelihood"], e["consequence"], e["outcome"]]
        for j, v in enumerate(vals, start=1):
            rc(ws, k, j, v, font=F_BOLD if j == 1 else F_BASE, align=CENTER if j > 2 else WRAP)
        if m in MATS:
            q = risk[(risk["mode"] == "L2") & (risk.material == m) & (risk.mitigation == "M0")].iloc[0]
            rc(ws, k, 9, p["t600_obs"], font=F_INPUT, fmt="0", align=CENTER)
            rc(ws, k, 10, p["R_max"], font=F_INPUT, align=CENTER)
            rc(ws, k, 11, float(q.R_S_mean), font=F_INPUT, fmt=SCI, align=CENTER)
            rc(ws, k, 12, float(q.MF_p50), font=F_INPUT, fmt="0", align=CENTER, fill=FILL_KEY)
        else:
            for j in (9, 10, 11, 12):
                rc(ws, k, j, "—", align=CENTER)
        ws.row_dimensions[k].height = 30
    note(ws, 12, "EFFECT는 정성적 순위 도구입니다(보고서 초록: 통계·시험자료 부족으로 정성 평가 채택). 같은 'D'라도 시험 t600은 30–1350 s, 재료 인자는 33–630배로 갈립니다.")
    note(ws, 13, "EFFECT 후속 과제(13장): '외벽 시스템 화재시험으로 외장재·단열재·공기층 조합의 점수 보강' → 이 계산서의 중규모·소형시료 자료가 그 근거 역할을 합니다.")


# --------------------------------------------------------------------------------------------- 12 시험계획
PLAN = [
    ("L1", "콘칼로리미터", "ISO 5660-1 (KS F ISO 5660-1)", "50 kW/m² 10분(준불연 판정) + 25·35·65 kW/m²(TRP·CHF)", "100×100 mm, 면재 포함·심재, 각 3회",
     "t_ig, pHRR, THR_600, EHC, TSP, CO", "D1·D2·D3·D6·G", "가연/준불연 게이트, 착화·발열(TGA 사각지대)", "1"),
    ("L1", "FT-IR (ATR)", "고체 정성분석 (예: ASTM E1252)", "4000–400 cm⁻¹, 분해능 4 cm⁻¹, 32회 누적", "심재·면재 각 3 위치",
     "고분자 식별, 이소시아누레이트·난연(P) 지수", "G (D1·D2)", "제품 식별·난연제 유무 → 전이함수 적용 여부", "1"),
    ("L1", "TGA–FTIR (발생가스)", "—", "N₂·공기, 5 K/min (기존 TGA와 같게), 전송관 200–250 °C", "2회", "HCN·CO·NCO·HCl 발생 온도·상대량", "D6",
     "유해가스 → FDS 수율·FED", "2"),
    ("L1", "TGA–DSC 보완", "ISO 11358-1 · ASTM E1269 / ISO 11357-4", "초기 질량 기록, 블랭크·사파이어 보정, EPS 가연성 공기 재측정", "2–3회",
     "잔류량(NU_MATL), c_p, 반응열", "D3·D4 (+L3 입력)", "'조건부·종점 의존' 판정 해소", "2"),
    ("L1", "(선택) MCC", "ASTM D7309", "1 K/s, 750 °C", "3회", "HRC, THR", "D2·D3", "r_max 대신 HRC로 확산식 개선", "3"),
    ("L1", "(선택) 총발열량", "ISO 1716", "봄베 열량계", "3회", "PCS (MJ/kg)", "D3", "A2/B 분류 근거", "3"),
    ("L1", "(선택) SEM/EDS·XRF", "—", "원소 분석", "2 위치", "P·Br·Al·Mg·Sb", "G", "난연제 확인(FT-IR 보완)", "3"),
    ("L2", "중공층 시험 보완", "현행 방법", "버너 소화 후 60분 관찰, CO/CO₂·열류속 계측, 반복 3회", "재료당 3회", "B·지속시간, 가스", "D4·D5·D6",
     "p_B 표본 확대(현재 2–3회)", "1"),
    ("L2", "EPS 가연성 재시험", "현행 방법", "30분 평가 완주", "2–3회", "t600, B", "D2·D4", "정량평가 제외 해소", "1"),
    ("L2", "배리어·차단재 포함 시험", "현행 방법 + 층간 배리어", "배리어 위·아래 열전대", "재료 2–3종", "배리어 통과 시각 t_b", "—", "M1 효과 p_cb·p_e1 근거", "1"),
    ("L3", "FDS 중규모 재현", "FDS 6.9 (PyroSim)", "격자 민감도, TGA 반응 입력(PyroSim 자료), 실측 k·ρ·c", "재료 5종", "t600·온도이력 오차", "D1·D2",
     "모델 검증(NFPA 551 8.2.2)", "1"),
    ("L3", "FDS 건물 외벽 시나리오", "FDS 6.9", "IE1 창 분출화염, IE2 외부 화원, 3–6개 층", "재료 5종 × 2", "층별 도달시각 → s, 창 열류속, FED", "D2·D3·D6",
     "s (건물 위험도의 최대 불확실성)", "1"),
    ("L3+", "(선택) 실대형 시험", "KS F 8414", "2–3종", "1회", "층간 확산시간", "D2", "s 직접 보정", "2"),
]


def plan(wb):
    ws = wb.create_sheet("12_시험계획")
    title(ws, "⑫ 시험·해석 계획 — 어떤 불확실성을 줄이는가", "우선순위 1 = 판정을 바꿀 수 있는 항목(민감도·정보가치 기준).")
    cols = ["척도", "시험·해석", "규격", "조건", "시편·반복", "산출 지표", "위험 차원", "해결하는 불확실성", "우선순위"]
    header(ws, 4, cols, widths=[6, 18, 22, 34, 16, 26, 12, 30, 8], height=28)
    for k, row in enumerate(PLAN, start=5):
        for j, v in enumerate(row, start=1):
            rc(ws, k, j, v, font=F_BOLD if j == 2 else F_BASE, align=CENTER if j in (1, 9) else WRAP,
               fill=FILL_KEY if (j == 9 and v == "1") else None)
        ws.row_dimensions[k].height = 32


# --------------------------------------------------------------------------------------------- 13 검증
def verify(wb, rows_a, rows_b, rows_c):
    ws = wb.create_sheet("13_검증")
    title(ws, "⑬ 검증 — 엑셀 수식과 파이썬(multiscale.py) 값 비교", "두 구현이 같은 값을 내면 계산 단계의 오류가 없다는 근거입니다(NFPA 551 8.2.1.1).")
    header(ws, 4, ["항목", "파이썬", "엑셀", "상대차"], widths=[40, 14, 14, 12])
    st = D["stats"]
    fusion = {f["material"]: f for f in D["fusion"]}
    cons = {c["material"]: c for c in D["consistency"]}
    items = [("σ_rep", st["s_rep"], "=sigma_rep"), ("τ (TGA 예측 편차)", st["tau"], "=tau_tga"), ("준불연 제품 간 SD", st["sd_between"], "=sd_between")]
    for m in CHAR:
        items.append((f"{m} t600 통합 (s)", fusion[m]["t600_post"], f"='7_통합(베이즈)'!N{rows_a[m]}"))
        items.append((f"{m} TGA 무게 w", fusion[m]["w_tga"], f"='7_통합(베이즈)'!I{rows_a[m]}"))
        items.append((f"{m} z (확산 정합)", cons[m]["z_spread"], f"='7_통합(베이즈)'!Q{rows_a[m]}"))
        items.append((f"{m} p_B 통합", fusion[m]["pB_post"], f"='7_통합(베이즈)'!J{rows_c[m]}"))
        items.append((f"{m} 예측확률 P(k)", cons[m]["p_pred_B"], f"='7_통합(베이즈)'!M{rows_c[m]}"))
    for m in ("PIR_C", "EPS_SC"):
        items.append((f"{m} z (게이트 없음)", cons[m]["z_spread"], f"='7_통합(베이즈)'!G{rows_b[m]}"))
    items.append(("EPS_SC 예측확률 P(k)", cons["EPS_SC"]["p_pred_B"], f"='7_통합(베이즈)'!M{rows_c['EPS_SC']}"))
    r = 5
    for name, py, f in items:
        rc(ws, r, 1, name, font=F_BOLD)
        rc(ws, r, 2, float(py), font=F_INPUT, fmt="0.000000", align=CENTER)
        rc(ws, r, 3, f, font=F_LINK, fmt="0.000000", align=CENTER)
        rc(ws, r, 4, f"=IF(MAX(ABS(B{r}),ABS(C{r}))=0,0,ABS(B{r}-C{r})/MAX(ABS(B{r}),ABS(C{r})))", fmt="0.0E+00", align=CENTER)
        r += 1
    rc(ws, r + 1, 1, "최대 상대차", font=F_BOLD)
    rc(ws, r + 1, 2, f"=MAX(D5:D{r - 1})", fmt="0.0E+00", align=CENTER, fill=FILL_KEY)
    rc(ws, r + 1, 3, f'=IF(B{r + 1}<1E-6,"일치","불일치 — 입력 확인")', font=F_BOLD, align=CENTER, fill=FILL_KEY)


def main():
    wb = Workbook()
    readme(wb)
    matrix(wb)
    cone(wb)
    tga_rows = tga(wb)
    ftir(wb)
    midscale(wb)
    tr = transfer(wb, tga_rows)
    rows_a, rows_b, rows_c = bayes(wb, tga_rows, tr)
    simulation(wb, tr)
    voi(wb)
    results(wb)
    effect(wb)
    plan(wb)
    verify(wb, rows_a, rows_b, rows_c)
    for ws in wb.worksheets:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
    dv = DataValidation(type="list", formula1='"c,sc"', allow_blank=True)
    wb["3_L1_TGA"].add_data_validation(dv)
    dv.add(f"C5:C{4 + len(ALL)}")
    wb.save(OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
