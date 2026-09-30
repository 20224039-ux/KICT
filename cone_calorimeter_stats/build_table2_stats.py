"""Cone calorimeter replicate statistics (mean ± SD) for the PUF skin-layer specimens.

Reads the per-quantity workbooks (141B_HRR/THR/SPR/CO/CO2.xlsx), each holding one sheet
per configuration (BSK, SKM, S1H, S2H, S1V, S2V) with the individual test curves
(0-600 s, 5 s step), and writes a workbook in the layout of the lab's STDEV.P template:

    individual test values (formulas on the raw-data sheets)
    -> STDEV.P / AVERAGE / ± / mean-SD / (mean-SD)/mean / 93 % check, plus STDEV.S

It also builds a revised Table 2 (mean ± SD with n, selectable STDEV.S / STDEV.P), a
comparison with the manuscript's current Table 2 (parsed from the .docx when given) and a
table of individual replicate values. Every statistic is an Excel formula, so the
workbook recalculates if a raw value is corrected.

Usage:
    python build_table2_stats.py --data-dir data --manuscript data/manuscript.docx \
        --out output/141B_PUF_cone_replicate_stats.xlsx
"""

import argparse
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
]
TABLE2_ORDER = ["HRRpeak", "HRRmean", "THRmean", "SPRpeak", "SPRmean", "COmean", "CO2mean", "THR600"]
MANUSCRIPT_HEADERS = {  # manuscript Table 2 header -> metric id
    "HRRpeak": "HRRpeak", "HRRmean": "HRRmean", "THRmean": "THRmean", "SPRpeak": "SPRpeak",
    "SPRmean": "SPRmean", "COmean": "COmean", "CO₂mean": "CO2mean", "CO2mean": "CO2mean",
}
# how the manuscript Table 2 values were obtained (summary sheets of the source workbooks)
MANUSCRIPT_METHOD = {
    "HRRpeak": "시험 평균 곡선의 최대값 (141B_HRR 'HRR' 시트 J3:O3)",
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
    """{metric id: {sample: (value, decimals shown)}} from the manuscript table headed 'Sample'."""
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
                out[key] = {r[0]: (float(r[j]), len(r[j].partition(".")[2])) for r in rows
                            if r and r[0] in SAMPLES}
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
    return f"={q(raw_sheet)}!{col}{RAW_LAST}"


def peak_time_formula(raw_sheet, cols, cells, prefix=""):
    """'<prefix>피크 시각(s): 10 / 30' built from the raw curves."""
    parts = []
    for col, cell in zip(cols, cells):
        parts.append(f"INDEX({q(raw_sheet)}!$A${RAW_FIRST}:$A${RAW_LAST},"
                     f"MATCH({cell},{q(raw_sheet)}!{col}{RAW_FIRST}:{col}{RAW_LAST},0))")
    return f'="{prefix}피크 시각(s): "&' + '&" / "&'.join(parts)


def write_stdev_sheets(wb, runs, col_of, values):
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
                c = ws.cell(r, 1 + k, per_test_formula(raw_sheet, col_of[(s, k)], kind))
                c.font, c.number_format = F_LINK, fmt
                cells.append(c.coordinate)
                cols.append(col_of[(s, k)])
            rng = f"B{r}:E{r}"
            ws.cell(r, 6, f"=COUNT({rng})")
            ws.cell(r, 7, f"=STDEVP({rng})").number_format = fmt
            ws.cell(r, 8, f"=AVERAGE({rng})").number_format = fmt
            ws.cell(r, 9, f'="±"&TEXT(G{r},"0.{"0" * pm_dec}")')
            ws.cell(r, 10, f"=H{r}-G{r}").number_format = fmt
            ws.cell(r, 11, f"=(H{r}-G{r})/H{r}").number_format = "0.0%"
            ws.cell(r, 12, f'=IF(K{r}>=0.93,"충족","93%미만")')
            ws.cell(r, 13, f"=STDEV({rng})").number_format = fmt
            for j in range(6, 14):
                ws.cell(r, j).font = F_BASE
            ws.cell(r, 9).alignment = Alignment(horizontal="right")
            # notes: sample size, S2H split, peak times
            notes = []
            if n != 3:
                notes.append(f"n={n}")
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


def write_table2(wb, loc):
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
            ws.cell(5 + i, col, f"=TEXT({q(sheet)}!$H${r},{fmt})&\" ± \"&TEXT({sd},{fmt})").font = F_LINK
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
    for r in (11, 12):
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
                if label == "CV":
                    f, fmt = f"={sd}/{mean}*100", "0.0"
                elif label == "평균값":
                    f, fmt = f"={mean}", meta[mid][5]
                else:
                    f, fmt = f"={sd}", meta[mid][5]
                c = ws.cell(r_top + 2 + i, col, f)
                c.font, c.number_format = F_LINK, fmt
        for i, s in enumerate(SAMPLES):
            ws.cell(r_top + 2 + i, 1, s).font = F_BOLD
            ws.cell(r_top + 2 + i, 2, f"=$B${5 + i}").font = F_BASE
        style_range(ws, [ws.cell(r_top + 1, j) for j in range(1, last_col + 1)], font=F_BOLD, fill=FILL_HEAD,
                    align=CENTER)
        ws.row_dimensions[r_top + 1].height = 32
    cv_rng = f"C{r_cv + 2}:{get_column_letter(last_col - 1)}{r_cv + 7}"  # Table 2 columns (without reference)
    ws.conditional_formatting.add(f"C{r_cv + 2}:{get_column_letter(last_col)}{r_cv + 7}",
                                  CellIsRule(operator="greaterThan", formula=["5"], fill=FILL_FLAG))
    ws.cell(r_cv + 8, 1, f'="CV 5% 초과: "&COUNTIF({cv_rng},">5")&" / "&COUNT({cv_rng})&" 칸 (Table 2의 7개 항목 × 6개 시료)"')
    ws.cell(r_cv + 8, 1).font = F_BOLD

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 6
    for j in range(first_col, last_col + 1):
        ws.column_dimensions[get_column_letter(j)].width = 21
    ws.freeze_panes = "C5"


def write_comparison(wb, loc, manuscript):
    ws = wb.create_sheet(SHEET_CMP)
    ws["A1"] = "원고 Table 2 값 vs 개별 시험 기반 재계산 평균 (재계산 = 각 시험에서 값을 구한 뒤 평균)"
    ws["A1"].font = F_TITLE
    heads = ["항목", "시료", "원고 Table 2", "재계산 평균\n(개별 시험 평균)", "차이\n(재계산−원고)", "차이 (%)",
             "원고\n소수 자릿수", "원고 자릿수로\n반올림 시 일치", "원고 값의 계산 방식", "n"]
    for j, h in enumerate(heads, start=1):
        ws.cell(2, j, h)
    style_range(ws, ws[2], font=F_BOLD, fill=FILL_HEAD, align=CENTER)
    ws.row_dimensions[2].height = 32
    meta = {m[0]: m for m in METRICS}
    r = 3
    for mid in TABLE2_ORDER:
        if mid not in manuscript:
            continue
        for s in SAMPLES:
            sheet, rr = loc[(mid, s)]
            fmt = meta[mid][5]
            value, decimals = manuscript[mid][s]
            ws.cell(r, 1, mid).font = F_BOLD
            ws.cell(r, 2, s).font = F_BOLD
            c = ws.cell(r, 3, value)
            c.font, c.number_format = F_INPUT, fmt
            c = ws.cell(r, 4, f"={q(sheet)}!$H${rr}")
            c.font, c.number_format = F_LINK, fmt
            ws.cell(r, 5, f"=D{r}-C{r}").number_format = fmt
            ws.cell(r, 6, f"=IF(C{r}=0,\"\",(D{r}-C{r})/C{r})").number_format = "0.0%"
            ws.cell(r, 7, decimals).font = F_INPUT
            ws.cell(r, 8, f'=IF(ABS(ROUND(D{r},G{r})-C{r})<10^-(G{r}+3),"일치","불일치")')
            ws.cell(r, 9, MANUSCRIPT_METHOD[mid]).font = F_NOTE
            ws.cell(r, 10, f"={q(sheet)}!$F${rr}").font = F_LINK
            for j in (5, 6, 8):
                ws.cell(r, j).font = F_BASE
            for j in (7, 8, 10):
                ws.cell(r, j).alignment = Alignment(horizontal="center")
            r += 1
        r += 1
    ws.conditional_formatting.add(f"H3:H{r}", CellIsRule(operator="equal", formula=['"불일치"'], fill=FILL_FLAG,
                                                          font=F_RED))
    ws.cell(r, 1, "원고 Table 2 값: 원고 파일(BKCS_PUF_SkinLayer_Final_260913.docx) Table 2에서 읽어 입력한 값. "
                  "'불일치' = 재계산 평균을 원고와 같은 자릿수로 반올림해도 원고 값과 다름.").font = F_NOTE
    for col, w in zip("ABCDEFGHIJ", [10, 7, 13, 15, 13, 10, 9, 13, 62, 5]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "C3"


def write_individual(wb, loc, runs):
    ws = wb.create_sheet(SHEET_IND)
    ws["A1"] = "개별 시험값 (Supplementary table 용) — 각 항목의 시험별 값은 STDEV 시트의 1~4열을 참조"
    ws["A1"].font = F_TITLE
    heads = ["Sample", "Test", "원본 열"] + [m[7].replace("\n[참고]", "") for m in METRICS]
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
            for j, m in enumerate(METRICS):
                sheet, rr = loc[(m[0], s)]
                c = ws.cell(r, 4 + j, f"={q(sheet)}!{get_column_letter(1 + k)}{rr}")
                c.font, c.number_format = F_LINK, m[5]
            r += 1
    ws.cell(r + 1, 1, "원본 열 = 각 파일(141B_HRR/THR/SPR/CO/CO2.xlsx) 시료별 시트의 같은 열. HRR–THR 시험 순서 일치는 "
                      "THR(600 s) = ∫HRR dt로 확인했고, SPR·CO·CO₂ 파일도 같은 순서로 정리되어 있다고 가정함.").font = F_NOTE
    for j, w in enumerate([8, 6, 9] + [15] * len(METRICS), start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "D3"


def write_readme(ws, n_by_sample, manuscript_given, peak_shift, split_samples):
    n_txt = ", ".join(f"{s} {n}" for s, n in n_by_sample.items())
    shift_txt = ", ".join(f"{s} {a:.2f} → {b:.2f}" for s, (a, b) in peak_shift.items()) or "없음"
    lines = [
        ("141B(HCFC-141b) PUF 6종 콘칼로리미터 반복시험 통계 — 리뷰어 코멘트(Table 2 표준편차) 대응용", F_TITLE),
        ("", None),
        ("구성", F_BOLD),
        (f"· {SHEET_HRR} / {SHEET_SPR} / {SHEET_GAS}: 원본 STDEV.P 양식 — 항목별 블록, 시료 6종 × 시험 1~4회,", None),
        ("   표준편차(STDEV.P) · 평균값 · ± · 평균값-표준편차 · (평균값-표준편차)/평균값 · 93% 판정, 추가로 STDEV.S", None),
        (f"· {SHEET_T2}: 논문 Table 2에 넣을 '평균 ± 표준편차' 문자열과 n", None),
        ("   (B2에서 SD 종류, 3행에서 소수 자릿수 선택) + CV(%) 표 + 오차막대용 숫자 표", None),
        (f"· {SHEET_CMP}: 현재 원고 Table 2 값과 개별 시험 기반 재계산 평균 비교" if manuscript_given
         else f"· {SHEET_CMP}: (원고 파일 미지정으로 생략)", None),
        (f"· {SHEET_IND}: 시험별 개별값 (Supplementary table 용)", None),
        ("· 원자료_HRR/THR/SPR/CO/CO2: 원본 파일 시료별 시트의 시험 곡선을 그대로 옮김 (값 수정 없음, 음수 기저선 값 포함)", None),
        ("", None),
        ("계산 방법", F_BOLD),
        ("· 각 시험 곡선에서 값을 먼저 구한 뒤(최대값, 0–600 s 평균, 600 s 값) 시료별로 평균·표준편차를 계산", None),
        ("   (ISO 5660-1처럼 개별 시편 결과의 평균).", None),
        ("· (평균값-표준편차)/평균값 = 1 − CV(STDEV.P 기준). '판정 ≥93%'는 CV ≤ 7%와 같음.", None),
        ("· 모든 통계값은 수식이므로 원자료 시트 값을 고치면 자동으로 다시 계산됨.", None),
        ("· 글자색: 파란색 = 원자료·입력값, 초록색 = 다른 시트를 참조하는 수식, 검정 = 같은 시트 수식. 노란 칸 = 선택·입력 칸.", None),
        ("", None),
        ("확인이 필요한 사항", F_BOLD),
        (f"1. 시료별 시험 수(n): {n_txt}", None),
        ("   — 원고 2.2·2.5절의 'triplicate'와 다름. 원 시험 기록을 확인한 뒤 실제 n을 표기해야 함.", None),
        ("2. 원고 HRRpeak는 '시험 평균 곡선의 최대값'이라, 시험마다 피크 시각이 다르면 개별 시험 피크의 평균보다 낮게 나옴.", None),
        (f"   평균 곡선 최대값 → 개별 시험 피크 평균 (1% 넘게 다른 시료): {shift_txt}", None),
        (f"   평균 ± SD로 쓰려면 평균도 개별 시험 기준으로 바꿔야 서로 맞음 ({SHEET_CMP} 참조).", None),
        ("3. 원고 2.5절 'triplicate reproducibility typically within 5% of the reported mean'", None),
        (f"   — 실제 CV는 여러 항목에서 5%를 넘음 ({SHEET_T2} 시트의 CV 표 참조).", None),
        *[(f"4. {s}는 {n_by_sample[s]}회 중 3번째 이후 시험이 1·2번보다 모든 항목(HRR·THR·SPR·CO·CO₂)에서 큼 "
            "— 시편·시험 조건 차이 여부 확인 권장.", None) for s in split_samples],
        ("5. THRmean(THR 누적곡선의 시간평균)은 통상 쓰이지 않는 지표 — 시험 종료 시점 THR(600 s)을 참고 열로 함께 계산함.",
         None),
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
    args = ap.parse_args()

    runs = {k: read_runs(args.data_dir / f, prefix) for k, (f, prefix, _, _) in QUANTITIES.items()}
    n_by_sample = {s: len(runs["HRR"][s]) for s in SAMPLES}
    for k in QUANTITIES:
        counts = {s: len(runs[k][s]) for s in SAMPLES}
        if counts != n_by_sample:
            raise ValueError(f"test counts differ between HRR and {k}: {counts} vs {n_by_sample}")

    # per-test values in numpy (used for notes and printed as a cross-check of the Excel formulas)
    reduce = {"max": np.max, "mean": np.mean, "end": lambda y: y[-1]}
    values = {(m[0], s): np.array([reduce[m[4]](y) for _, y in runs[m[3]][s]]) for m in METRICS for s in SAMPLES}

    wb = openpyxl.Workbook()
    readme = wb.active
    readme.title = "설명"
    col_of = None
    raw_sheets = {}
    for k, (f, *_rest) in QUANTITIES.items():
        raw_sheets[k] = write_raw_sheet(wb, k, runs[k], f)
        col_of = col_of or raw_sheets[k]
    loc = write_stdev_sheets(wb, runs, col_of, values)
    write_table2(wb, loc)
    manuscript = read_manuscript_table2(args.manuscript) if args.manuscript else None
    if manuscript:
        write_comparison(wb, loc, manuscript)
    write_individual(wb, loc, runs)
    # peak of the test-averaged HRR curve (manuscript method) vs mean of the individual peaks
    peak_shift = {}
    for s in SAMPLES:
        curves = np.array([y for _, y in runs["HRR"][s]])
        of_avg, avg_of = curves.mean(axis=0).max(), curves.max(axis=1).mean()
        if abs(avg_of - of_avg) > 0.01 * of_avg:
            peak_shift[s] = (of_avg, avg_of)
    # configurations whose later tests exceed the first two in every metric
    split_samples = [s for s in SAMPLES if n_by_sample[s] >= 4 and all(
        min(values[(m[0], s)][2:]) > max(values[(m[0], s)][:2]) for m in METRICS)]
    write_readme(readme, n_by_sample, bool(manuscript), peak_shift, split_samples)

    order = ["설명", SHEET_HRR, SHEET_SPR, SHEET_GAS, SHEET_T2] + ([SHEET_CMP] if manuscript else []) + \
        [SHEET_IND] + [f"원자료_{k}" for k in QUANTITIES]
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
    print("n per configuration:", n_by_sample)
    for m in METRICS:
        for s in SAMPLES:
            v = values[(m[0], s)]
            print(f"{m[0]:8s} {s}: mean={v.mean():.6g} sdP={v.std():.4g} sdS={v.std(ddof=1):.4g}")


if __name__ == "__main__":
    main()
