"""Cone calorimeter replicate statistics (mean ± SD) for the PUF skin-layer specimens.

Reads the per-quantity workbooks (141B_HRR/THR/SPR/CO/CO2.xlsx), each holding one sheet
per configuration (BSK, SKM, S1H, S2H, S1V, S2V) with the individual test curves
(0-600 s, 5 s step), and writes a workbook in the layout of the lab's STDEV.P template:

    individual test values (formulas on the raw-data sheets)
    -> STDEV.P / AVERAGE / ± / mean-SD / (mean-SD)/mean / 93 % check, plus STDEV.S

It also builds a revised Table 2 (mean ± SD with n, selectable STDEV.S / STDEV.P), a
comparison with the manuscript's current Table 2 (parsed from the .docx when given) and a
table of individual replicate values. Every statistic is an Excel formula, so the
workbook recalculates if a raw value is corrected. Tests can be excluded (--exclude BSK-1):
the raw data stay in the workbook, the test is marked '제외' on the '시험 포함 여부' sheet
(switchable back to '포함') and the exclusion is written into the Table 2 footnote.

Usage:
    python build_table2_stats.py --data-dir data --manuscript data/manuscript.docx \
        --out output/141B_PUF_cone_replicate_stats.xlsx [--exclude BSK-1 --reason "BSK-1=..."]
"""

import argparse
import re
from pathlib import Path

import numpy as np
import openpyxl
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

SAMPLES = ["BSK", "SKM", "S1H", "S2H", "S1V", "S2V"]  # Table 2 order
TIMES = list(range(0, 601, 5))
SRC_ROWS = range(3, 3 + len(TIMES))  # rows holding t = 0..600 s in the source sheets

# key -> (file, column-header prefix in the source sheets, unit, description)
QUANTITIES = {
    "HRR": ("141B_HRR.xlsx", "HRR", "kW m⁻²", "열방출률"),
    "THR": ("141B_THR.xlsx", "THR", "MJ m⁻²", "총열방출량(누적)"),
    "SPR": ("141B_SPR.xlsx", "SPR", "m² s⁻¹", "연기발생률"),
    "CO": ("141B_CO.xlsx", "COP", "g s⁻¹", "CO 생성속도"),
    "CO2": ("141B_CO2.xlsx", "CO2P", "g s⁻¹", "CO₂ 생성속도"),
}

RAW_FIRST = 4  # output raw sheets: row 1 title, row 2 labels, row 3 source column
RAW_LAST = RAW_FIRST + len(TIMES) - 1

SHEET_HRR = "STDEV(1) HRR·THR"
SHEET_SPR = "STDEV(2) SPR"
SHEET_GAS = "STDEV(3) CO·CO2"
SHEET_T2 = "Table 2 (mean±SD)"
SHEET_CMP = "Table 2 비교"
SHEET_IND = "개별값 (Table S)"
SHEET_INC = "시험 포함 여부"
SHEET_OUT = "이상치 검정"
INC_FIRST = 3  # first test row on the inclusion sheet
DASH = "–"

# id, sheet, block title, raw quantity, per-test reduction, number format, ± decimals,
# Table 2 header, default Table 2 decimals
METRICS = [
    ("HRRpeak", SHEET_HRR, "HRRpeak (kW m⁻²) — 시험별 HRR 곡선의 최대값", "HRR", "max", "0.00", 2,
     "HRRpeak (kW m⁻²)", 2),
    ("HRRmean", SHEET_HRR, "HRRmean (kW m⁻²) — 시험별 HRR의 0–600 s 평균", "HRR", "mean", "0.00", 2,
     "HRRmean (kW m⁻²)", 2),
    ("THRmean", SHEET_HRR, "THRmean (MJ m⁻²) — 시험별 THR 곡선의 0–600 s 시간평균 (원고 Table 2 정의)",
     "THR", "mean", "0.00", 2, "THRmean (MJ m⁻²)", 2),
    ("THR600", SHEET_HRR, "THR at 600 s (MJ m⁻²) — 시험 종료(600 s) 시점 총열방출량 [참고: 통상적인 THR]",
     "THR", "end", "0.00", 2, "THR at 600 s (MJ m⁻²)\n[참고]", 2),
    ("SPRpeak", SHEET_SPR, "SPRpeak (m² s⁻¹) — 시험별 SPR 곡선의 최대값", "SPR", "max", "0.00000", 4,
     "SPRpeak (m² s⁻¹)", 4),
    ("SPRmean", SHEET_SPR, "SPRmean (m² s⁻¹) — 시험별 SPR의 0–600 s 평균", "SPR", "mean", "0.000000", 5,
     "SPRmean (m² s⁻¹)", 5),
    ("COmean", SHEET_GAS, "COmean (g s⁻¹) — 시험별 CO 생성속도의 0–600 s 평균", "CO", "mean", "0.000000", 5,
     "COmean (g s⁻¹)", 5),
    ("CO2mean", SHEET_GAS, "CO₂mean (g s⁻¹) — 시험별 CO₂ 생성속도의 0–600 s 평균", "CO2", "mean",
     "0.000000", 5, "CO₂mean (g s⁻¹)", 5),
    ("COpeak", SHEET_GAS, "COpeak (g s⁻¹) — 시험별 CO 생성속도 최대값 [참고: 일부 시험은 후반 훈소 구간에서 최대]", "CO",
     "max", "0.000000", 5, "COpeak (g s⁻¹)\n[참고]", 5),
    ("CO2peak", SHEET_GAS, "CO₂peak (g s⁻¹) — 시험별 CO₂ 생성속도 최대값 [참고]", "CO2", "max", "0.00000", 4,
     "CO₂peak (g s⁻¹)\n[참고]", 4),
    ("COCO2", SHEET_GAS, "CO/CO₂ (–) — 시험별 CO 평균 / CO₂ 평균 (0–600 s) [참고]", "CO", "ratio", "0.000", 3,
     "CO/CO₂ (–)\n[참고]", 3),
]
TABLE2_ORDER = ["HRRpeak", "HRRmean", "THRmean", "SPRpeak", "SPRmean", "COmean", "CO2mean", "THR600",
                "COpeak", "CO2peak", "COCO2"]
N_TABLE2 = 7  # columns that correspond to the manuscript's Table 2; the rest are reference columns
MANUSCRIPT_HEADERS = {  # manuscript Table 2 header -> metric id
    "HRRpeak": "HRRpeak", "HRRmean": "HRRmean", "THRmean": "THRmean", "SPRpeak": "SPRpeak",
    "SPRmean": "SPRmean", "COmean": "COmean", "CO₂mean": "CO2mean", "CO2mean": "CO2mean",
}
# how the manuscript Table 2 values were obtained (summary sheets of the source workbooks)
# why a recomputed mean can differ from the manuscript value (besides excluded tests)
MANUSCRIPT_CAUSE = {
    "HRRpeak": "원고 값 ≠ 시험별 최대값의 평균 (초기 원고 260913은 시험 평균 곡선의 최대값을 사용)",
    "HRRmean": "원고 요약시트의 시작점(0–5 s) 음수 기저선 0 처리 차이",
}
MANUSCRIPT_METHOD = {
    "HRRpeak": "초기 원고(260913): 시험 평균 곡선의 최대값 (141B_HRR 'HRR' 시트 J3:O3)",
    "HRRmean": "시험 평균 곡선의 0–600 s 평균, 첫 1–2개 점 0 처리 (141B_HRR 'HRR' 시트 J4:O4)",
    "THRmean": "시험 평균 THR 곡선의 0–600 s 시간평균 (141B_THR 'SPR' 시트 J3:O3)",
    "SPRpeak": "시험 평균 곡선의 최대값 (141B_SPR 'SPR' 시트 J3:O3)",
    "SPRmean": "시험 평균 곡선의 0–600 s 평균 (141B_SPR 'SPR' 시트 J4:O4)",
    "COmean": "시험 평균 곡선의 0–600 s 평균 (141B_CO 'CO' 시트 K5:P5)",
    "CO2mean": "시험 평균 곡선의 0–600 s 평균 (141B_CO2 'CO2' 시트 K4:P4)",
}

FONT = "맑은 고딕"
F_BASE = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_TITLE = Font(name=FONT, size=12, bold=True)
F_INPUT = Font(name=FONT, size=10, color="0000FF")  # hardcoded inputs
F_LINK = Font(name=FONT, size=10, color="008000")  # formulas pulling from another sheet
F_NOTE = Font(name=FONT, size=9, color="595959")
F_RED = Font(name=FONT, size=10, color="C00000", bold=True)
FILL_HEAD = PatternFill("solid", fgColor="D9E1F2")
FILL_BLOCK = PatternFill("solid", fgColor="F2F2F2")
FILL_SELECT = PatternFill("solid", fgColor="FFFF00")
FILL_FLAG = PatternFill("solid", fgColor="F8CBAD")
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def q(sheet):
    """Sheet name quoted for a formula reference."""
    return "'" + sheet.replace("'", "''") + "'"


def read_runs(path, prefix):
    """{sample: [(source column letter, np.array over TIMES), ...]} from one quantity workbook."""
    wb = openpyxl.load_workbook(path, data_only=True)
    runs = {}
    for s in SAMPLES:
        ws = wb[s]
        t = [ws.cell(r, 1).value for r in SRC_ROWS]
        if t != TIMES:
            raise ValueError(f"{path.name} [{s}]: time column is not 0..600 s in 5 s steps")
        cols = []
        for c in range(2, ws.max_column + 1):
            head = ws.cell(2, c).value
            if not (isinstance(head, str) and head.split("(")[0].strip() == prefix):
                continue
            y = [ws.cell(r, c).value for r in SRC_ROWS]
            if not all(isinstance(v, (int, float)) for v in y):
                raise ValueError(f"{path.name} [{s}] column {get_column_letter(c)}: non-numeric value")
            cols.append((get_column_letter(c), np.array(y, float)))
        runs[s] = cols
    return runs


def read_manuscript_table2(docx_path):
    """{metric id: {sample: (mean, decimals shown, SD or None)}} from the manuscript table headed 'Sample'."""
    import docx

    for table in docx.Document(docx_path).tables:
        rows = [[c.text.strip() for c in r.cells] for r in table.rows]
        head = next((r for r in rows if r and r[0] == "Sample"), None)
        if not head or not any("HRRpeak" in h for h in head):
            continue
        out = {}
        for j, h in enumerate(head[1:], start=1):
            key = MANUSCRIPT_HEADERS.get(h.split("(")[0].strip())
            if key:
                out[key] = {}
                for r in rows:
                    if r and r[0] in SAMPLES:
                        m = re.match(r"\s*([\d.]+)(?:\s*±\s*([\d.]+))?", r[j])  # plain value or "mean ± SD"
                        mean, sd = m.group(1), m.group(2)
                        out[key][r[0]] = (float(mean), len(mean.partition(".")[2]), float(sd) if sd else None)
        return out
    raise ValueError(f"{docx_path}: Table 2 (header 'Sample', 'HRRpeak ...') not found")


def style_range(ws, cells, font=None, fill=None, fmt=None, align=None):
    for c in cells:
        if font is not None:
            c.font = font
        if fill is not None:
            c.fill = fill
        if fmt is not None:
            c.number_format = fmt
        if align is not None:
            c.alignment = align


def write_raw_sheet(wb, key, runs, src_file):
    ws = wb.create_sheet(f"원자료_{key}")
    _, _, unit, desc = QUANTITIES[key]
    ws["A1"] = f"{key} ({unit}) {desc} — 개별 시험 곡선, 원본: {src_file} 의 시료별 시트 (0–600 s, 5 s 간격, 값 수정 없음)"
    ws["A1"].font = F_BOLD
    ws["A2"], ws["A3"] = "Time (s)", "원본 시트!열"
    col_of = {}
    col = 2
    for s in SAMPLES:
        for i, (src_col, y) in enumerate(runs[s], start=1):
            ws.cell(2, col, f"{s}-{i}")
            ws.cell(3, col, f"{s}!{src_col}")
            for k, v in enumerate(y):
                ws.cell(RAW_FIRST + k, col, float(v)).font = F_INPUT
            col_of[(s, i)] = get_column_letter(col)
            col += 1
    for k, t in enumerate(TIMES):
        ws.cell(RAW_FIRST + k, 1, t).font = F_INPUT
    style_range(ws, [c for r in ws.iter_rows(min_row=2, max_row=3) for c in r], font=F_BOLD, fill=FILL_HEAD,
                align=CENTER)
    ws.column_dimensions["A"].width = 12
    for c in range(2, col):
        ws.column_dimensions[get_column_letter(c)].width = 11
    ws.freeze_panes = "B4"
    return col_of


def per_test_formula(raw_sheet, col, kind):
    rng = f"{q(raw_sheet)}!{col}{RAW_FIRST}:{col}{RAW_LAST}"
    if kind == "max":
        return f"=MAX({rng})"
    if kind == "mean":
        return f"=AVERAGE({rng})"
    if kind == "ratio":  # CO/CO2: mean CO production over mean CO2 production of the same test
        co2 = f"{q('원자료_CO2')}!{col}{RAW_FIRST}:{col}{RAW_LAST}"
        return f"=AVERAGE({rng})/AVERAGE({co2})"
    return f"={q(raw_sheet)}!{col}{RAW_LAST}"


def peak_time_formula(raw_sheet, cols, cells, prefix=""):
    """'<prefix>피크 시각(s): 10 / 30' built from the raw curves."""
    parts = []
    for col, cell in zip(cols, cells):
        parts.append(f"IFERROR(INDEX({q(raw_sheet)}!$A${RAW_FIRST}:$A${RAW_LAST},"
                     f"MATCH({cell},{q(raw_sheet)}!{col}{RAW_FIRST}:{col}{RAW_LAST},0)),\"제외\")")
    return f'="{prefix}피크 시각(s): "&' + '&" / "&'.join(parts)


def write_inclusion_sheet(wb, runs, excluded):
    """One row per test with a 포함/제외 switch; returns {(sample, k): row}."""
    ws = wb.create_sheet(SHEET_INC)
    ws["A1"] = ("시험별 포함 여부 — D열을 '제외'로 두면 그 시험은 모든 항목(HRR·THR·SPR·CO·CO₂)의 통계에서 빠짐 "
                "(원자료는 그대로 유지, '포함'으로 바꾸면 다시 계산됨)")
    ws["A1"].font = F_BOLD
    for j, h in enumerate(["Sample", "Test", "원본 열", "포함 여부", "제외 사유 (논문 각주에 들어감)", "각주용"], start=1):
        ws.cell(2, j, h)
    style_range(ws, ws[2], font=F_BOLD, fill=FILL_HEAD, align=CENTER)
    dv = DataValidation(type="list", formula1='"포함,제외"', allow_blank=False)
    ws.add_data_validation(dv)
    row_of = {}
    r = INC_FIRST
    for s in SAMPLES:
        for k, (src_col, _) in enumerate(runs["HRR"][s], start=1):
            ws.cell(r, 1, s).font = F_BOLD
            ws.cell(r, 2, k)
            ws.cell(r, 3, f"{s}!{src_col}").font = F_NOTE
            c = ws.cell(r, 4, "제외" if (s, k) in excluded else "포함")
            c.font, c.fill, c.alignment = F_INPUT, FILL_SELECT, Alignment(horizontal="center")
            dv.add(c.coordinate)
            c = ws.cell(r, 5, excluded.get((s, k), ""))
            c.font, c.fill = F_INPUT, FILL_SELECT
            ws.cell(r, 6, f'=IF(D{r}="제외",A{r}&"-"&B{r}&IF(E{r}="",""," ("&E{r}&")"),"")').font = F_NOTE
            row_of[(s, k)] = r
            r += 1
    last = r - 1
    ws.conditional_formatting.add(f"A{INC_FIRST}:D{last}", FormulaRule(formula=[f'$D{INC_FIRST}="제외"'],
                                                                         fill=FILL_FLAG))
    ws.conditional_formatting.add(f"E{INC_FIRST}:E{last}",
                                  FormulaRule(formula=[f'AND($D{INC_FIRST}="제외",$E{INC_FIRST}="")'],
                                              fill=PatternFill("solid", fgColor="FF7C80", bgColor="FF7C80")))
    ws.cell(last + 2, 1, "제외 사유가 비어 있으면 빨간색으로 표시됨. 시험을 빼려면 결과와 무관한 기술적 사유(시편 결함, 장비·절차 이상 "
                         "등)를 적고, 논문에 제외 사실과 사유를 밝혀야 함.").font = F_NOTE
    for col, w in zip("ABCDEF", [8, 6, 9, 10, 60, 30]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A3"
    return row_of


def write_stdev_sheets(wb, runs, col_of, values, inc_row, excluded):
    """Template-style blocks; returns {(metric, sample): (sheet, row)}."""
    headers = ["항목 / 시료", 1, 2, 3, 4, "n", "표준편차\n(STDEV.P)", "평균값", "± (STDEV.P)", "평균값-표준편차",
               "(평균값-표준편차)\n/평균값", "판정 (≥93%)", "표준편차\n(STDEV.S, n−1)", "비고"]
    widths = [34, 11, 11, 11, 11, 5, 12, 12, 12, 13, 14, 11, 13, 46]
    loc = {}
    next_row = {}
    for sheet in (SHEET_HRR, SHEET_SPR, SHEET_GAS):
        ws = wb.create_sheet(sheet)
        for j, h in enumerate(headers, start=1):
            ws.cell(1, j, h)
            ws.column_dimensions[get_column_letter(j)].width = widths[j - 1]
        style_range(ws, ws[1], font=F_BOLD, fill=FILL_HEAD, align=CENTER)
        ws.row_dimensions[1].height = 30
        ws.freeze_panes = "B2"
        next_row[sheet] = 2

    for mid, sheet, title, raw_key, kind, fmt, pm_dec, _, _ in METRICS:
        ws = wb[sheet]
        raw_sheet = f"원자료_{raw_key}"
        r0 = next_row[sheet]
        ws.cell(r0, 1, title)
        style_range(ws, [ws.cell(r0, j) for j in range(1, 15)], font=F_BOLD, fill=FILL_BLOCK)
        for i, s in enumerate(SAMPLES, start=1):
            r = r0 + i
            ws.cell(r, 1, s).font = F_BOLD
            n = len(runs[raw_key][s])
            cells, cols = [], []
            for k in range(1, n + 1):
                flag = f"{q(SHEET_INC)}!$D${inc_row[(s, k)]}"
                expr = per_test_formula(raw_sheet, col_of[(s, k)], kind)[1:]
                c = ws.cell(r, 1 + k, f'=IF({flag}="제외","",{expr})')
                c.font, c.number_format = F_LINK, fmt
                cells.append(c.coordinate)
                cols.append(col_of[(s, k)])
            rng = f"B{r}:E{r}"
            ws.cell(r, 6, f"=COUNT({rng})")
            ws.cell(r, 7, f'=IF(F{r}>=2,STDEVP({rng}),"{DASH}")').number_format = fmt
            ws.cell(r, 8, f'=IF(F{r}>=1,AVERAGE({rng}),"{DASH}")').number_format = fmt
            ws.cell(r, 9, f'=IF(F{r}>=2,"±"&TEXT(G{r},"0.{"0" * pm_dec}"),"{DASH}")')
            ws.cell(r, 10, f'=IF(F{r}>=2,H{r}-G{r},"{DASH}")').number_format = fmt
            ws.cell(r, 11, f'=IF(F{r}>=2,(H{r}-G{r})/H{r},"{DASH}")').number_format = "0.0%"
            ws.cell(r, 12, f'=IF(F{r}<2,"판정 불가(n<2)",IF(K{r}>=0.93,"충족","93%미만"))')
            ws.cell(r, 13, f'=IF(F{r}>=2,STDEV({rng}),"{DASH}")').number_format = fmt
            for j in range(6, 14):
                ws.cell(r, j).font = F_BASE
                ws.cell(r, j).alignment = Alignment(horizontal="right")
            # notes: excluded tests, S2H split, peak times
            notes = [f"{s}-{k} 제외" for k in range(1, n + 1) if (s, k) in excluded]
            v = values[(mid, s)]
            if s == "S2H" and n == 4 and min(v[2:]) > max(v[:2]):
                notes.append("3·4번 시험 > 1·2번 시험")
            note = ", ".join(notes)
            if kind == "max":
                ws.cell(r, 14, peak_time_formula(raw_sheet, cols, cells, f"{note}; " if note else ""))
            else:
                ws.cell(r, 14, note)
            ws.cell(r, 14).font = F_NOTE
            loc[(mid, s)] = (sheet, r)
        ws.conditional_formatting.add(f"L{r0 + 1}:L{r0 + 6}",
                                      CellIsRule(operator="equal", formula=['"93%미만"'], font=F_RED))
        next_row[sheet] = r0 + 8
    return loc


def write_table2(wb, loc, n_tests):
    ws = wb.create_sheet(SHEET_T2)
    ws["A1"] = ("Table 2 (revised). Combustion characteristics of PUF specimens with HCFC-141b blowing agent "
                "at an external radiant heat flux of 50 kW m⁻² (ISO 5660-1): mean ± SD of individual tests")
    ws["A1"].font = F_TITLE
    ws["A2"], ws["B2"] = "SD 종류 (선택)", "STDEV.S"
    ws["C2"] = "← STDEV.S = 표본표준편차(n−1, 논문 보고 시 일반적) / STDEV.P = 모집단 표준편차(원본 양식)"
    ws["A3"] = "소수 자릿수 (입력)"
    ws["A2"].font = ws["A3"].font = F_BOLD
    ws["B2"].font, ws["B2"].fill = F_INPUT, FILL_SELECT
    ws["C2"].font = F_NOTE
    dv = DataValidation(type="list", formula1='"STDEV.S,STDEV.P"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("B2")

    meta = {m[0]: m for m in METRICS}
    first_col = 3  # C
    ws.cell(4, 1, "Sample")
    ws.cell(4, 2, "n")
    for j, mid in enumerate(TABLE2_ORDER):
        col = first_col + j
        L = get_column_letter(col)
        ws.cell(3, col, meta[mid][8]).font = F_INPUT
        ws.cell(3, col).fill = FILL_SELECT
        ws.cell(4, col, meta[mid][7])
        fmt = f'IF({L}$3=0,"0","0."&REPT("0",{L}$3))'
        for i, s in enumerate(SAMPLES):
            sheet, r = loc[(mid, s)]
            sd = f'IF($B$2="STDEV.P",{q(sheet)}!$G${r},{q(sheet)}!$M${r})'
            n_ref, mean = f"{q(sheet)}!$F${r}", f"{q(sheet)}!$H${r}"
            ws.cell(5 + i, col, f'=IF({n_ref}>=2,TEXT({mean},{fmt})&" ± "&TEXT({sd},{fmt}),'
                                f'IF({n_ref}=1,TEXT({mean},{fmt})&"ᵃ","{DASH}"))').font = F_LINK
            ws.cell(5 + i, col).alignment = Alignment(horizontal="center")
    for i, s in enumerate(SAMPLES):
        sheet, r = loc[("HRRpeak", s)]
        ws.cell(5 + i, 1, s).font = F_BOLD
        ws.cell(5 + i, 2, f"={q(sheet)}!$F${r}").font = F_LINK
        ws.cell(5 + i, 2).alignment = Alignment(horizontal="center")
    last_col = first_col + len(TABLE2_ORDER) - 1
    style_range(ws, [ws.cell(4, j) for j in range(1, last_col + 1)], font=F_BOLD, fill=FILL_HEAD, align=CENTER)
    ws.row_dimensions[4].height = 32
    ws["A11"] = ('="Values are mean ± "&IF($B$2="STDEV.P","population standard deviation",'
                 '"standard deviation (n − 1)")&" of the individual replicate tests; n = number of tests per '
                 'configuration. Each parameter was first determined for every test and then averaged."')
    ws["A12"] = ("HRRpeak and SPRpeak: maximum of each test curve. HRRmean, SPRmean, COmean and CO₂mean: average "
                 "over 0–600 s. THRmean: time-average of the THR curve over 0–600 s (definition used in the "
                 "current Table 2); THR at 600 s: total heat released at the end of the test (reference column).")
    inc_last = INC_FIRST + n_tests - 1
    ws["A13"] = (f'=IF(COUNTIF($B$5:$B$10,1)>0,"ᵃ Single test (n = 1); SD not available. ","")'
                 f'&IF(COUNTIF({q(SHEET_INC)}!$D${INC_FIRST}:$D${inc_last},"제외")=0,"",'
                 f'"Excluded test(s): "&_xlfn.TEXTJOIN("; ",TRUE,{q(SHEET_INC)}!$F${INC_FIRST}:$F${inc_last})&".")')
    for r in (11, 12, 13):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
        ws.cell(r, 1).font = F_NOTE
        ws.cell(r, 1).alignment = WRAP
        ws.row_dimensions[r].height = 27

    # CV table (checks the manuscript sentence "typically within 5% of the reported mean")
    r_cv = 14
    ws.cell(r_cv, 1, "CV (%) = SD / 평균 × 100 — 원고 2.5절 'triplicate reproducibility typically within 5%' 확인용"
                     " (5% 초과 칸 강조)").font = F_BOLD
    numeric_tables = [("CV", r_cv), ("평균값", r_cv + 10), ("표준편차 (위 SD 종류)", r_cv + 20)]
    for label, r_top in numeric_tables:
        if label != "CV":
            ws.cell(r_top, 1, f"{label} — 숫자 (그래프 오차막대용)").font = F_BOLD
        ws.cell(r_top + 1, 1, "Sample")
        ws.cell(r_top + 1, 2, "n")
        for j, mid in enumerate(TABLE2_ORDER):
            col = first_col + j
            ws.cell(r_top + 1, col, meta[mid][7])
            for i, s in enumerate(SAMPLES):
                sheet, r = loc[(mid, s)]
                mean = f"{q(sheet)}!$H${r}"
                sd = f'IF($B$2="STDEV.P",{q(sheet)}!$G${r},{q(sheet)}!$M${r})'
                n_ref = f"{q(sheet)}!$F${r}"
                if label == "CV":
                    f, fmt = f'=IF({n_ref}>=2,{sd}/{mean}*100,"{DASH}")', "0.0"
                elif label == "평균값":
                    f, fmt = f"={mean}", meta[mid][5]
                else:
                    f, fmt = f'=IF({n_ref}>=2,{sd},"{DASH}")', meta[mid][5]
                c = ws.cell(r_top + 2 + i, col, f)
                c.font, c.number_format = F_LINK, fmt
        for i, s in enumerate(SAMPLES):
            ws.cell(r_top + 2 + i, 1, s).font = F_BOLD
            ws.cell(r_top + 2 + i, 2, f"=$B${5 + i}").font = F_BASE
        style_range(ws, [ws.cell(r_top + 1, j) for j in range(1, last_col + 1)], font=F_BOLD, fill=FILL_HEAD,
                    align=CENTER)
        ws.row_dimensions[r_top + 1].height = 32
    cv_rng = f"C{r_cv + 2}:{get_column_letter(first_col + N_TABLE2 - 1)}{r_cv + 7}"  # manuscript Table 2 columns only
    ws.conditional_formatting.add(f"C{r_cv + 2}:{get_column_letter(last_col)}{r_cv + 7}",
                                  FormulaRule(formula=[f"AND(ISNUMBER(C{r_cv + 2}),C{r_cv + 2}>5)"],
                                              fill=FILL_FLAG))
    ws.cell(r_cv + 8, 1, f'="CV 5% 초과: "&COUNTIF({cv_rng},">5")&" / "&COUNT({cv_rng})&" 칸 (Table 2의 7개 항목 × 6개 시료 중 n≥2인 칸)"')
    ws.cell(r_cv + 8, 1).font = F_BOLD

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 6
    for j in range(first_col, last_col + 1):
        ws.column_dimensions[get_column_letter(j)].width = 21
    ws.freeze_panes = "C5"


def write_comparison(wb, loc, manuscript, n_tests):
    """Manuscript Table 2 vs recomputed means, graded 일치 / 근사 일치 / 불일치 with the cause."""
    ws = wb.create_sheet(SHEET_CMP)
    ws["A1"] = "원고 Table 2 값 vs 개별 시험 기반 재계산 평균 (재계산 = 각 시험에서 값을 구한 뒤 평균)"
    ws["A1"].font = F_TITLE
    heads = ["항목", "시료", "원고 Table 2", "재계산 평균\n(개별 시험 평균)", "재계산\n(원고 자릿수 반올림)",
             "차이\n(재계산−원고)", "차이 (%)", "원고\n소수 자릿수", "끝자리 차이\n(단위 수)", "판정", "차이 원인",
             "원고 값의 계산 방식", "n", "원고 SD", "재계산 SD\n(STDEV.S)", "SD 판정\n(원고 자릿수)"]
    for j, h in enumerate(heads, start=1):
        ws.cell(3, j, h)
    style_range(ws, ws[3], font=F_BOLD, fill=FILL_HEAD, align=CENTER)
    ws.row_dimensions[3].height = 32
    inc = q(SHEET_INC)
    inc_last = INC_FIRST + n_tests - 1
    meta = {m[0]: m for m in METRICS}
    first = r = 4
    for mid in TABLE2_ORDER:
        if mid not in manuscript:
            continue
        for s in SAMPLES:
            sheet, rr = loc[(mid, s)]
            fmt = meta[mid][5]
            value, decimals, man_sd = manuscript[mid][s]
            ws.cell(r, 1, mid).font = F_BOLD
            ws.cell(r, 2, s).font = F_BOLD
            c = ws.cell(r, 3, value)
            c.font, c.number_format = F_INPUT, fmt
            c = ws.cell(r, 4, f"={q(sheet)}!$H${rr}")
            c.font, c.number_format = F_LINK, fmt
            ws.cell(r, 5, f'=IF(ISNUMBER(D{r}),ROUND(D{r},H{r}),"{DASH}")').number_format = fmt
            ws.cell(r, 6, f'=IF(ISNUMBER(D{r}),D{r}-C{r},"{DASH}")').number_format = fmt
            ws.cell(r, 7, f'=IF(AND(ISNUMBER(D{r}),C{r}<>0),(D{r}-C{r})/C{r},"{DASH}")').number_format = "0.0%"
            ws.cell(r, 8, decimals).font = F_INPUT
            ws.cell(r, 9, f'=IF(ISNUMBER(E{r}),ROUND(ABS(E{r}-C{r})*10^H{r},0),"{DASH}")')
            ws.cell(r, 10, f'=IF(ISNUMBER(I{r}),IF(I{r}=0,"일치",IF(I{r}<=2,"근사 일치","불일치")),"{DASH}")')
            excl = (f'COUNTIFS({inc}!$A${INC_FIRST}:$A${inc_last},"{s}",'
                    f'{inc}!$D${INC_FIRST}:$D${inc_last},"제외")>0')
            cause = MANUSCRIPT_CAUSE.get(mid, "")
            if cause:
                ws.cell(r, 11, f'=IF(J{r}="일치","",IF({excl},"{s} 시험 제외 영향. ","")&"{cause}")')
            else:
                ws.cell(r, 11, f'=IF(J{r}="일치","",IF({excl},"{s} 시험 제외 영향.","원인 확인 필요"))')
            ws.cell(r, 12, MANUSCRIPT_METHOD[mid]).font = F_NOTE
            ws.cell(r, 13, f"={q(sheet)}!$F${rr}").font = F_LINK
            if man_sd is not None:  # manuscript also reports an SD: compare it with STDEV.S of the tests
                c = ws.cell(r, 14, man_sd)
                c.font, c.number_format = F_INPUT, fmt
                c = ws.cell(r, 15, f"={q(sheet)}!$M${rr}")
                c.font, c.number_format = F_LINK, fmt
                ws.cell(r, 16, f'=IF(ISNUMBER(O{r}),IF(ABS(ROUND(O{r},H{r})-N{r})<10^-(H{r}+3),"일치","불일치"),"{DASH}")')
                ws.cell(r, 16).alignment = Alignment(horizontal="center")
            for j in (5, 6, 7, 9, 10, 11):
                ws.cell(r, j).font = F_BASE
            for j in (8, 9, 10, 13):
                ws.cell(r, j).alignment = Alignment(horizontal="center")
            r += 1
        r += 1
    last = r - 1
    ws["A2"] = (f'="판정 합계 — 일치 "&COUNTIF(J{first}:J{last},"일치")&" · 근사 일치 "&COUNTIF(J{first}:J{last},"근사 일치")'
                f'&" · 불일치 "&COUNTIF(J{first}:J{last},"불일치")')
    ws["A2"].font = F_BOLD
    ws.conditional_formatting.add(f"J{first}:J{last}", CellIsRule(operator="equal", formula=['"불일치"'],
                                                                  fill=FILL_FLAG, font=F_RED))
    ws.conditional_formatting.add(f"J{first}:J{last}", CellIsRule(operator="equal", formula=['"근사 일치"'],
                                                                  fill=PatternFill("solid", fgColor="FFF2CC",
                                                                                   bgColor="FFF2CC")))
    notes = [
        "판정: 재계산 평균을 원고와 같은 소수 자릿수로 반올림해서 비교. 일치 = 같음 · 근사 일치 = 원고 마지막 자리에서 1~2 차이"
        "(예: 15.25 vs 15.27, 반올림만으로 같아지지는 않지만 무시할 수준) · 불일치 = 그보다 큰 차이.",
        "원고 Table 2 값: 원고 파일(BKCS_PUF_SkinLayer_Final_260913.docx) Table 2에서 읽어 입력한 값. "
        f"재계산 평균은 '{SHEET_INC}' 시트에서 '제외'한 시험을 뺀 값.",
    ]
    for i, text in enumerate(notes):
        ws.cell(r + i, 1, text).font = F_NOTE
    ws.conditional_formatting.add(f"P{first}:P{last}", CellIsRule(operator="equal", formula=['"불일치"'],
                                                                  fill=FILL_FLAG, font=F_RED))
    for col, w in zip("ABCDEFGHIJKLMNOP", [10, 7, 13, 15, 14, 13, 10, 9, 10, 10, 58, 58, 5, 11, 12, 11]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "C4"


def write_outlier_check(wb, loc):
    """Dixon Q test (95 %) on the included tests of every metric and configuration."""
    ws = wb.create_sheet(SHEET_OUT)
    ws["A1"] = ("이상치 검정 — Dixon Q (95%, 임계값 n=3: 0.970, n=4: 0.829; Rorabacher 1991). "
                "포함된 시험만 사용, 모든 시료·항목에 같은 기준 적용")
    ws["A1"].font = F_TITLE
    heads = ["항목", "시료", "n", "최소", "최대", "Dixon Q", "임계값 (95%)", "판정", "의심값"]
    for j, h in enumerate(heads, start=1):
        ws.cell(2, j, h)
    style_range(ws, ws[2], font=F_BOLD, fill=FILL_HEAD, align=CENTER)
    meta = {m[0]: m for m in METRICS}
    first = r = 3
    for mid in TABLE2_ORDER:
        for s in SAMPLES:
            sheet, rr = loc[(mid, s)]
            rng = f"{q(sheet)}!$B${rr}:$E${rr}"
            fmt = meta[mid][5]
            gap_hi, gap_lo = f"(LARGE({rng},1)-LARGE({rng},2))", f"(SMALL({rng},2)-SMALL({rng},1))"
            ws.cell(r, 1, mid).font = F_BOLD
            ws.cell(r, 2, s).font = F_BOLD
            ws.cell(r, 3, f"={q(sheet)}!$F${rr}").font = F_LINK
            ws.cell(r, 4, f'=IF(C{r}>=1,MIN({rng}),"{DASH}")').number_format = fmt
            ws.cell(r, 5, f'=IF(C{r}>=1,MAX({rng}),"{DASH}")').number_format = fmt
            ws.cell(r, 6, f'=IF(C{r}<3,"{DASH}",IF(E{r}=D{r},0,MAX({gap_hi},{gap_lo})/(E{r}-D{r})))')
            ws.cell(r, 6).number_format = "0.000"
            ws.cell(r, 7, f'=IF(C{r}=3,0.97,IF(C{r}=4,0.829,IF(C{r}=5,0.71,"{DASH}")))').number_format = "0.000"
            ws.cell(r, 8, f'=IF(C{r}<3,"검정 불가 (n<3)",IF(F{r}>G{r},"이상치 후보","이상치 없음"))')
            ws.cell(r, 9, f'=IF(H{r}="이상치 후보",IF({gap_hi}>={gap_lo},E{r},D{r}),"")').number_format = fmt
            for j in range(4, 10):
                ws.cell(r, j).font = F_BASE
            for j in (3, 6, 7, 8):
                ws.cell(r, j).alignment = Alignment(horizontal="center")
            r += 1
        r += 1
    ws.conditional_formatting.add(f"H{first}:H{r}", CellIsRule(operator="equal", formula=['"이상치 후보"'],
                                                               fill=FILL_FLAG, font=F_RED))
    ws.cell(r, 1, "이상치 후보가 나와도 자동으로 빼지 않음. 제외는 기술적 사유가 확인될 때 '시험 포함 여부' 시트에서 하고 "
                  "논문에 밝힘. n=2 이하는 어느 값이 벗어났는지 통계로 판단할 수 없음.").font = F_NOTE
    for col, w in zip("ABCDEFGHI", [10, 7, 5, 12, 12, 10, 12, 16, 12]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "C3"


def write_individual(wb, runs, col_of, inc_row):
    ws = wb.create_sheet(SHEET_IND)
    ws["A1"] = ("개별 시험값 (Supplementary table 용) — 원자료 곡선에서 직접 계산, "
                f"'{SHEET_INC}'에서 '제외'한 시험도 값은 보이되 통계에는 쓰이지 않음")
    ws["A1"].font = F_TITLE
    heads = ["Sample", "Test", "원본 열", "포함 여부"] + [m[7].replace("\n[참고]", "") for m in METRICS]
    for j, h in enumerate(heads, start=1):
        ws.cell(2, j, h)
    style_range(ws, ws[2], font=F_BOLD, fill=FILL_HEAD, align=CENTER)
    ws.row_dimensions[2].height = 32
    r = 3
    for s in SAMPLES:
        n = len(runs["HRR"][s])
        for k in range(1, n + 1):
            ws.cell(r, 1, s).font = F_BOLD
            ws.cell(r, 2, k)
            ws.cell(r, 3, f"{s}!{runs['HRR'][s][k - 1][0]}").font = F_NOTE
            ws.cell(r, 4, f"={q(SHEET_INC)}!$D${inc_row[(s, k)]}").font = F_LINK
            ws.cell(r, 4).alignment = Alignment(horizontal="center")
            for j, m in enumerate(METRICS):
                c = ws.cell(r, 5 + j, per_test_formula(f"원자료_{m[3]}", col_of[(s, k)], m[4]))
                c.font, c.number_format = F_LINK, m[5]
            r += 1
    ws.conditional_formatting.add(f"A3:{get_column_letter(4 + len(METRICS))}{r - 1}",
                                  FormulaRule(formula=['$D3="제외"'], fill=FILL_FLAG))
    ws.cell(r + 1, 1, "원본 열 = 각 파일(141B_HRR/THR/SPR/CO/CO2.xlsx) 시료별 시트의 같은 열. HRR–THR 시험 순서 일치는 "
                      "THR(600 s) = ∫HRR dt로 확인했고, SPR·CO·CO₂ 파일도 같은 순서로 정리되어 있다고 가정함.").font = F_NOTE
    for j, w in enumerate([8, 6, 9, 9] + [15] * len(METRICS), start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "D3"


def write_readme(ws, n_by_sample, manuscript_given, peak_shift, split_samples, excluded, outliers):
    n_txt = ", ".join(f"{s} {n}" for s, n in n_by_sample.items())
    n_used = {s: n - sum(1 for (es, _) in excluded if es == s) for s, n in n_by_sample.items()}
    excl_lines = [(f"· {s}-{k}: {reason or '제외 사유 미입력'}  → {s} 통계는 {n_used[s]}회 시험 기준"
                   + (" (표준편차 계산 불가)" if n_used[s] < 2 else ""), None)
                  for (s, k), reason in sorted(excluded.items())]
    shift_txt = ", ".join(f"{s} {a:.2f} → {b:.2f}" for s, (a, b) in peak_shift.items()) or "없음"
    lines = [
        ("141B(HCFC-141b) PUF 6종 콘칼로리미터 반복시험 통계 — 리뷰어 코멘트(Table 2 표준편차) 대응용", F_TITLE),
        ("", None),
        ("구성", F_BOLD),
        (f"· {SHEET_HRR} / {SHEET_SPR} / {SHEET_GAS}: 원본 STDEV.P 양식 — 항목별 블록, 시료 6종 × 시험 1~4회,", None),
        ("   표준편차(STDEV.P) · 평균값 · ± · 평균값-표준편차 · (평균값-표준편차)/평균값 · 93% 판정, 추가로 STDEV.S", None),
        (f"· {SHEET_T2}: 논문 Table 2에 넣을 '평균 ± 표준편차' 문자열과 n", None),
        ("   (B2에서 SD 종류, 3행에서 소수 자릿수 선택) + CV(%) 표 + 오차막대용 숫자 표", None),
        (f"· {SHEET_CMP}: 현재 원고 Table 2 값과 개별 시험 기반 재계산 평균 비교 — 일치 / 근사 일치(끝자리 1~2 차이) / "
         "불일치 판정과 차이 원인" if manuscript_given else f"· {SHEET_CMP}: (원고 파일 미지정으로 생략)", None),
        (f"· {SHEET_OUT}: 모든 시료·항목에 Dixon Q 이상치 검정(95%)을 같은 기준으로 적용한 결과", None),
        (f"· {SHEET_IND}: 시험별 개별값 (Supplementary table 용)", None),
        (f"· {SHEET_INC}: 시험별 포함/제외 스위치와 제외 사유 (제외 시 모든 항목의 통계에서 빠지고 Table 2 각주에 자동 기재)",
         None),
        ("· 원자료_HRR/THR/SPR/CO/CO2: 원본 파일 시료별 시트의 시험 곡선을 그대로 옮김 (값 수정 없음, 음수 기저선 값 포함)", None),
        ("", None),
        ("계산 방법", F_BOLD),
        ("· 각 시험 곡선에서 값을 먼저 구한 뒤(최대값, 0–600 s 평균, 600 s 값) 시료별로 평균·표준편차를 계산", None),
        ("   (ISO 5660-1처럼 개별 시편 결과의 평균).", None),
        ("· (평균값-표준편차)/평균값 = 1 − CV(STDEV.P 기준). '판정 ≥93%'는 CV ≤ 7%와 같음.", None),
        ("· 모든 통계값은 수식이므로 원자료 시트 값을 고치면 자동으로 다시 계산됨.", None),
        ("· 글자색: 파란색 = 원자료·입력값, 초록색 = 다른 시트를 참조하는 수식, 검정 = 같은 시트 수식. 노란 칸 = 선택·입력 칸.", None),
        ("", None),
        *([("제외한 시험 (원자료는 그대로 두고 통계에서만 제외)", F_BOLD)] + excl_lines + [("", None)] if excluded else []),
        ("확인이 필요한 사항", F_BOLD),
        (f"1. 원본 파일의 시료별 시험 수: {n_txt}", None),
        ("   — 원고 2.2·2.5절의 'triplicate'와 다름. 원 시험 기록을 확인한 뒤 실제 n을 표기해야 함.", None),
        ("2. 원고 HRRpeak는 '시험 평균 곡선의 최대값'이라, 시험마다 피크 시각이 다르면 개별 시험 피크의 평균보다 낮게 나옴.", None),
        (f"   평균 곡선 최대값 → 개별 시험 피크 평균 (1% 넘게 다른 시료, 제외 시험 반영): {shift_txt}", None),
        (f"   평균 ± SD로 쓰려면 평균도 개별 시험 기준으로 바꿔야 서로 맞음 ({SHEET_CMP} 참조).", None),
        ("3. 원고 2.5절 'triplicate reproducibility typically within 5% of the reported mean'", None),
        (f"   — 실제 CV는 여러 항목에서 5%를 넘음 ({SHEET_T2} 시트의 CV 표 참조).", None),
        *[(f"4. {s}는 {n_by_sample[s]}회 중 3번째 이후 시험이 1·2번보다 모든 항목(HRR·THR·SPR·CO·CO₂)에서 큼 "
            "— 시편·시험 조건 차이 여부 확인 권장.", None) for s in split_samples],
        ("5. THRmean(THR 누적곡선의 시간평균)은 통상 쓰이지 않는 지표 — 시험 종료 시점 THR(600 s)을 참고 열로 함께 계산함.",
         None),
        ("6. 이상치 검정(Dixon Q, 95%, n≥3인 시료): " + (", ".join(f"{m} {s} {v:.4g}" for m, s, v in outliers)
                                                    if outliers else "모든 항목에서 이상치 후보 없음"), None),
    ]
    for i, (text, font) in enumerate(lines, start=1):
        c = ws.cell(i, 1, text)
        c.font = font or F_BASE
        c.alignment = Alignment(wrap_text=False, vertical="top")
    ws.column_dimensions["A"].width = 120


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    here = Path(__file__).resolve().parent
    ap.add_argument("--data-dir", type=Path, default=here / "data")
    ap.add_argument("--manuscript", type=Path, default=None, help="manuscript .docx holding Table 2 (optional)")
    ap.add_argument("--out", type=Path, default=here / "output" / "141B_PUF_cone_replicate_stats.xlsx")
    ap.add_argument("--exclude", action="append", default=[], metavar="SAMPLE-k",
                    help="test to leave out of the statistics, e.g. BSK-1 (repeatable)")
    ap.add_argument("--reason", action="append", default=[], metavar="SAMPLE-k=TEXT",
                    help="reason for an excluded test, shown in the Table 2 footnote")
    args = ap.parse_args()

    runs = {k: read_runs(args.data_dir / f, prefix) for k, (f, prefix, _, _) in QUANTITIES.items()}
    n_by_sample = {s: len(runs["HRR"][s]) for s in SAMPLES}
    for k in QUANTITIES:
        counts = {s: len(runs[k][s]) for s in SAMPLES}
        if counts != n_by_sample:
            raise ValueError(f"test counts differ between HRR and {k}: {counts} vs {n_by_sample}")
    reasons = dict(item.split("=", 1) for item in args.reason)
    excluded = {}
    for test in args.exclude:
        s, _, k = test.rpartition("-")
        if s not in n_by_sample or not k.isdigit() or not 1 <= int(k) <= n_by_sample[s]:
            raise ValueError(f"--exclude {test}: no such test")
        excluded[(s, int(k))] = reasons.get(test, "")
    n_tests = sum(n_by_sample.values())

    # per-test values in numpy (used for notes and printed as a cross-check of the Excel formulas)
    reduce = {"max": np.max, "mean": np.mean, "end": lambda y: y[-1]}
    values = {}
    for m in METRICS:
        for s in SAMPLES:
            if m[4] == "ratio":
                values[(m[0], s)] = np.array([np.mean(a) / np.mean(b) for (_, a), (_, b) in zip(runs["CO"][s], runs["CO2"][s])])
            else:
                values[(m[0], s)] = np.array([reduce[m[4]](y) for _, y in runs[m[3]][s]])

    wb = openpyxl.Workbook()
    readme = wb.active
    readme.title = "설명"
    col_of = None
    raw_sheets = {}
    for k, (f, *_rest) in QUANTITIES.items():
        raw_sheets[k] = write_raw_sheet(wb, k, runs[k], f)
        col_of = col_of or raw_sheets[k]
    inc_row = write_inclusion_sheet(wb, runs, excluded)
    loc = write_stdev_sheets(wb, runs, col_of, values, inc_row, excluded)
    write_table2(wb, loc, n_tests)
    manuscript = read_manuscript_table2(args.manuscript) if args.manuscript else None
    if manuscript:
        write_comparison(wb, loc, manuscript, n_tests)
    write_outlier_check(wb, loc)
    write_individual(wb, runs, col_of, inc_row)
    # peak of the test-averaged HRR curve (manuscript method) vs mean of the individual peaks
    peak_shift = {}
    for s in SAMPLES:
        curves = np.array([y for _, y in runs["HRR"][s]])
        kept = [k for k in range(len(curves)) if (s, k + 1) not in excluded]
        of_avg, avg_of = curves.mean(axis=0).max(), curves[kept].max(axis=1).mean()
        if abs(avg_of - of_avg) > 0.01 * of_avg:
            peak_shift[s] = (of_avg, avg_of)
    # configurations whose later tests exceed the first two in every metric
    split_samples = [s for s in SAMPLES if n_by_sample[s] >= 4 and all(
        min(values[(m[0], s)][2:]) > max(values[(m[0], s)][:2]) for m in METRICS if m[4] != "ratio")]
    # Dixon Q (95 %) on the included tests, same rule as the outlier sheet
    q_crit = {3: 0.970, 4: 0.829, 5: 0.710}
    outliers = []
    for m in METRICS:
        for s in SAMPLES:
            v = np.sort([x for k, x in enumerate(values[(m[0], s)], start=1) if (s, k) not in excluded])
            if len(v) in q_crit and v[-1] > v[0]:
                hi, lo = v[-1] - v[-2], v[1] - v[0]
                if max(hi, lo) / (v[-1] - v[0]) > q_crit[len(v)]:
                    outliers.append((m[0], s, v[-1] if hi >= lo else v[0]))
    write_readme(readme, n_by_sample, bool(manuscript), peak_shift, split_samples, excluded, outliers)

    order = ["설명", SHEET_INC, SHEET_HRR, SHEET_SPR, SHEET_GAS, SHEET_T2] + ([SHEET_CMP] if manuscript else []) + \
        [SHEET_OUT, SHEET_IND] + [f"원자료_{k}" for k in QUANTITIES]
    wb._sheets = [wb[n] for n in order]
    wb.active = 0
    for ws in wb.worksheets:  # print each summary sheet landscape, one page wide
        if not ws.title.startswith("원자료_"):
            ws.page_setup.orientation = "landscape"
            ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
            ws.sheet_properties.pageSetUpPr.fitToPage = True
    args.out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(args.out)

    print(f"saved {args.out}")
    print("tests per configuration:", n_by_sample, "| excluded:", [f"{s}-{k}" for s, k in excluded])
    for m in METRICS:
        for s in SAMPLES:
            v = np.array([x for k, x in enumerate(values[(m[0], s)], start=1) if (s, k) not in excluded])
            sd = (f"sdP={v.std():.4g} sdS={v.std(ddof=1):.4g}" if len(v) >= 2 else "sd n/a (n<2)")
            print(f"{m[0]:8s} {s}: n={len(v)} mean={v.mean():.6g} {sd}")


if __name__ == "__main__":
    main()
