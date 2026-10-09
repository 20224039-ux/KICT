#!/usr/bin/env python3
"""중공층(100T) + 외단열 100 mm 화재시험 – 재료별 화재확산 경향과 3000 mm 기준 at² 형 위험지수.

Cavity fire tests (100 mm cavity, 100 mm external insulation): 18 thermocouples
(three per height at 500, 1000, 1500, 2000, 2500, 3000 mm; CH1-3 = 500 mm ... CH16-18 = 3000 mm),
1 Hz logs. Materials: EPS combustible / semi-noncombustible, PIR (sc), PF board exterior / interior.

    python analysis.py data/EPS.xlsx data/PF.xlsx data/PIR_SC_2_max.xlsx

Sheet names identify the tests: '(n) EPS_C', '(n) EPS_SC', '(n) PF_semi_out', '(n) PF_semi_in'
(column 'm' followed by CH1..CH18) and 'sc_2 max' (column 's' followed by one max column per height).
Writes figures/*.png, results/*.csv and results/summary.json next to this file.
"""
from __future__ import annotations

import argparse
import csv
import json
import re
from itertools import combinations
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
import openpyxl  # noqa: E402
from scipy import stats  # noqa: E402

HERE = Path(__file__).resolve().parent
HEIGHTS = (500, 1000, 1500, 2000, 2500, 3000)
TOP = len(HEIGHTS) - 1                 # index of the 3000 mm level
EVAL_END_S = 1800                      # 30-min evaluation window (common to every full-length test)
SUSTAIN_S = 30                         # "for at least 30 s" (BS 8414 / BR 135 / KS F 8414 style)
THETA_REF = 600                        # K, temperature rise of the fire-spread criterion
THRESHOLDS = (300, 400, 500, 600, 700, 800)
FRONT_THETA = 500                      # K, flame-front marker for the vertical trajectory
GRADE_TG = (450, 900, 1800)            # s: t600 boundaries (7.5 / 15 / 30 min, factor-2 ladder like NFPA t_g)
GRADE_NAMES = ("IV", "III", "II", "I")
MAX_GAP_S = 60                         # interior sensor gaps up to this long are interpolated
MIN_ISLAND_S = 10                      # valid runs shorter than this between faults are faults too
BAD_LOW_C, BAD_HIGH_C = -20.0, 1400.0  # physically impossible readings -> sensor fault
COOL_RATE = -5.0                       # °C/s, 11-s moving average: collapse of a lower level
COOL_LEVELS = 2                        # of the 500-1500 mm levels, together = lower-level cooling event
BURNER_PHASE_S = 120                   # front arrivals earlier than this are the burner flame itself
ROBUST_THETA = (400, 500, 600, 700)
ROBUST_HOLD = (1, 10, 30)
EARLY_S = 120

# material key -> (Korean label, English label, colour, marker); colours validated all-pairs (CVD dE >= 13)
MATERIALS = {
    "EPS_C": ("EPS 가연성", "EPS combustible", "#eda100", "o"),
    "EPS_SC": ("EPS 준불연", "EPS semi-noncomb.", "#e87ba4", "s"),
    "PIR_SC": ("PIR 준불연(sc)", "PIR (sc)", "#008300", "^"),
    "PF_out": ("PF 외장용 준불연", "PF exterior", "#2a78d6", "D"),
    "PF_in": ("PF 내장용 준불연", "PF interior", "#4a3aa7", "v"),
}
SHEET_KIND = {"EPS_C": "EPS_C", "EPS_SC": "EPS_SC", "PF_semi_out": "PF_out", "PF_semi_in": "PF_in"}
REPEAT_STYLE = {0: (0, (1, 1.5)), 1: "-", 2: (0, (5, 2)), 3: (0, (1.5, 1.5))}

# chart tokens (dataviz reference palette, light surface)
SURFACE, INK, INK2, MUTED = "#fcfcfb", "#0b0b0b", "#52514e", "#898781"
GRID, AXIS, BAND = "#e1e0d9", "#c3c2b7", "#efeee9"
HEIGHT_COLORS = ["#c3c2b7", "#86b6ef", "#5598e7", "#2a78d6", "#1c5cab", "#0d366b"]  # 500 mm grey, ordinal ramp
GRADE_FILL = {"IV": "#f3d9d9", "III": "#f7e6dc", "II": "#f6f1e1", "I": "#eef3ec"}
trapezoid = getattr(np, "trapezoid", None) or np.trapz


# ----------------------------------------------------------------------------- helpers

def num(v):
    if isinstance(v, bool):
        return np.nan
    if isinstance(v, (int, float)):
        return float(v)
    return np.nan


def smooth(y, w=11):
    """Centred moving average that ignores NaN."""
    ok = np.isfinite(y)
    k = np.ones(w)
    numr = np.convolve(np.pad(np.where(ok, y, 0.0), w // 2, mode="edge"), k, mode="valid")
    den = np.convolve(np.pad(ok.astype(float), w // 2, mode="edge"), k, mode="valid")
    with np.errstate(invalid="ignore", divide="ignore"):
        return np.where(den > w / 2, numr / den, np.nan)


def sustained_arrival(dT, theta, hold=SUSTAIN_S):
    """Start (s) of the first window of `hold` s in which dT >= theta throughout; NaN if none."""
    run = 0
    for i, ok in enumerate(np.nan_to_num(dT, nan=-np.inf) >= theta):
        run = run + 1 if ok else 0
        if run >= hold:
            return float(i - hold + 1)
    return np.nan


def sustained_envelope(dT, hold=SUSTAIN_S):
    """Highest rise that has been held for `hold` s up to each time (running max of the trailing minimum)."""
    out = np.full(len(dT), -np.inf)
    x = np.nan_to_num(dT, nan=-np.inf)
    for i in range(hold - 1, len(x)):
        out[i] = x[i - hold + 1:i + 1].min()
    return np.maximum.accumulate(np.where(np.isfinite(out), out, 0.0))


def interpolate_gaps(x, max_gap=MAX_GAP_S):
    """Linear interpolation over interior NaN runs no longer than max_gap samples."""
    x = x.copy()
    bad = ~np.isfinite(x)
    i, n, filled = 0, len(x), []
    while i < n:
        if bad[i]:
            j = i
            while j < n and bad[j]:
                j += 1
            if i > 0 and j < n and j - i <= max_gap:
                x[i:j] = np.interp(np.arange(i, j), [i - 1, j], [x[i - 1], x[j]])
                filled.append((i, j - 1))
            i = j
        else:
            i += 1
    return x, filled


def drop_islands(bad, min_len=MIN_ISLAND_S):
    """Mark valid runs shorter than min_len that sit between faults (or the record start) as faults."""
    bad = bad.copy()
    i, n = 0, len(bad)
    while i < n:
        if not bad[i]:
            j = i
            while j < n and not bad[j]:
                j += 1
            if j - i < min_len and (i == 0 or bad[i - 1]) and j < n and bad[j]:
                bad[i:j] = True
            i = j
        else:
            i += 1
    return bad


def grade_of(t600, record_s):
    if np.isfinite(t600):
        for g, tg in zip(GRADE_NAMES, GRADE_TG):
            if t600 <= tg:
                return g
        return "I"
    return "I" if record_s >= EVAL_END_S else "n/a"


# ----------------------------------------------------------------------------- data

def read_tests(paths):
    """Return {test_id: dict} with raw 18-channel blocks (or 6 max columns for the PIR file)."""
    tests = {}
    for path in paths:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for ws in wb.worksheets:
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            if not rows:
                continue
            title, hdr = ws.title.strip(), [str(h).strip() if h is not None else "" for h in rows[0]]
            m = re.match(r"^\((\d+)\)\s*(EPS_C|EPS_SC|PF_semi_out|PF_semi_in)$", title)
            if m:
                rep, mat = int(m.group(1)), SHEET_KIND[m.group(2)]
                i0 = hdr.index("m")
                width = i0 + 19
                body = [r + [None] * (width - len(r)) for r in rows[1:]]
                tcol = np.array([num(r[i0]) for r in body])
                ch = np.array([[num(v) for v in r[i0 + 1:i0 + 19]] for r in body])
                empty = ~np.isfinite(ch).any(axis=1)
                end = int(np.argmax(empty)) if empty.any() else len(ch)   # first all-empty row ends the log
                tests[f"{mat}-{rep}"] = dict(material=mat, repeat=rep, source=f"{Path(path).name} / {ws.title}",
                                             ch=ch[:end], tcol=tcol[:end], max_only=False)
            elif re.match(r"^sc_(\d+)\s*max$", title, re.I):
                rep = int(re.match(r"^sc_(\d+)", title, re.I).group(1))
                i_s = hdr.index("s")
                cols = [hdr.index(next(h for h in hdr if re.search(rf"_{z}$", h))) for z in HEIGHTS]
                body = [r + [None] * (max(cols) + 1 - len(r)) for r in rows[1:]]
                tcol = np.array([num(r[i_s]) for r in body])
                H = np.array([[num(r[c]) for c in cols] for r in body])
                end = int(np.argmax(~np.isfinite(H).any(axis=1))) if (~np.isfinite(H).any(axis=1)).any() else len(H)
                tests[f"PIR_SC-{rep}"] = dict(material="PIR_SC", repeat=rep, source=f"{Path(path).name} / {ws.title}",
                                              H_raw=H[:end], tcol=tcol[:end], max_only=True)
    return tests


def check_time_base(tcol):
    """Rows are 1 s apart; the 'm' column is in minutes (1/60 steps) but may switch to seconds."""
    t = tcol[np.isfinite(tcol)]
    d = np.diff(t)
    minutes = np.isclose(d, 1 / 60, atol=1e-4)
    seconds = np.isclose(d, 1.0, atol=1e-3)
    other = ~(minutes | seconds)
    note = None
    if minutes.any() and seconds.any():
        note = f"time column switches from minutes to seconds at row {int(np.argmax(seconds)) + 1}"
    return dict(steps_ok=bool(other.sum() <= 1), unit_switch=note)


def find_splices(tests):
    """Rows that belong to another sheet (copy-paste), found by identical 18-channel rows at the same row index.

    For every sheet pair and column shift (-1, 0, +1) a row is 'shared' when >= 12 channels are finite in
    both and all of them agree; the sheet that is discontinuous where the shared block begins is the one
    that was overwritten and is truncated just before it.
    """
    found = []
    names = [k for k, v in tests.items() if not v["max_only"]]
    for a, b in combinations(names, 2):
        A, B = tests[a]["ch"], tests[b]["ch"]
        n = min(len(A), len(B))
        for s in (-1, 0, 1):
            ja = np.arange(18)
            jb = ja + s
            ok = (jb >= 0) & (jb < 18)
            X, Y = A[:n][:, ja[ok]], B[:n][:, jb[ok]]
            fin = np.isfinite(X) & np.isfinite(Y)
            same = np.isclose(X, Y, atol=0.05) & fin
            with np.errstate(invalid="ignore"):
                spread = np.nanmax(np.where(fin, X, np.nan), axis=1) - np.nanmin(np.where(fin, X, np.nan), axis=1)
            shared = (fin.sum(1) >= 12) & (same.sum(1) == fin.sum(1)) & (np.nan_to_num(spread) > 1.0)
            if shared.sum() < 10:
                continue
            idx = np.flatnonzero(shared)
            r0, r1 = int(idx[0]), int(idx[-1])
            if r0 == 0:
                found.append(dict(a=a, b=b, shift=s, rows=[r0, r1], victim=None, note="identical from row 0"))
                continue
            jump = {k: float(np.nansum(np.abs(M[r0] - M[r0 - 1]))) for k, M in ((a, A), (b, B))}
            typical = {k: float(np.nanmedian(np.nansum(np.abs(np.diff(M[:r0], axis=0)), axis=1))) for k, M in ((a, A), (b, B))}
            victim = max(jump, key=jump.get)
            found.append(dict(a=a, b=b, shift=s, rows=[r0, r1], n_shared=int(shared.sum()), victim=victim,
                              donor=b if victim == a else a, jump_victim=jump[victim],
                              jump_donor=jump[b if victim == a else a], typical_victim=typical[victim]))
    return found


def prepare(tests):
    """QA: splice truncation, sensor-fault removal, gap interpolation, height values (max of 3 channels)."""
    qa = {"splices": find_splices(tests), "tests": {}}
    cut = {}
    for s in qa["splices"]:
        if s.get("victim"):
            cut[s["victim"]] = min(cut.get(s["victim"], 10 ** 9), s["rows"][0])
    for k, v in tests.items():
        note = {"source": v["source"]}
        note.update(check_time_base(v["tcol"]))
        if v["max_only"]:
            H = v["H_raw"].copy()
            H[(H < BAD_LOW_C) | (H > BAD_HIGH_C)] = np.nan
            note["channels"] = "max per height only"
            T0 = H[0].copy()
        else:
            ch = v["ch"].copy()
            if k in cut:
                note["truncated_at_s"] = cut[k]
                ch = ch[:cut[k]]
            faults = (ch < BAD_LOW_C) | (ch > BAD_HIGH_C) | ~np.isfinite(ch)
            faults = np.column_stack([drop_islands(faults[:, j]) for j in range(18)])
            note["fault_samples"] = {f"CH{j + 1}": int(faults[:, j].sum()) for j in range(18) if faults[:, j].any()}
            ch[faults] = np.nan
            filled = {}
            for j in range(18):
                ch[:, j], f = interpolate_gaps(ch[:, j])
                if f:
                    filled[f"CH{j + 1}"] = f
            note["interpolated_gaps"] = filled
            still = {f"CH{j + 1}": int((~np.isfinite(ch[:, j])).sum()) for j in range(18) if (~np.isfinite(ch[:, j])).any()}
            note["unrecoverable"] = still
            H = np.column_stack([np.max(ch[:, 3 * i:3 * i + 3], axis=1) for i in range(6)])  # NaN if any channel bad
            with np.errstate(all="ignore"):
                T0 = np.array([np.nanmax(ch[0, 3 * i:3 * i + 3]) if np.isfinite(ch[0, 3 * i:3 * i + 3]).any()
                               else np.nanmedian(ch[0]) for i in range(6)])          # ambient from the valid TCs at t = 0
            v["ch_clean"] = ch
            hottest = []
            for i in range(6):
                sub = ch[:, 3 * i:3 * i + 3]
                ok = np.isfinite(sub).all(axis=1)
                hottest.append(np.bincount(np.argmax(sub[ok], axis=1), minlength=3) / max(ok.sum(), 1))
            note["share_hottest_channel_abc"] = np.round(np.array(hottest), 3).tolist()
        n = min(len(H), EVAL_END_S + 1)
        v["H"] = H[:n]
        v["t"] = np.arange(n, dtype=float)
        v["record_s"] = int(len(H) - 1)
        v["T0"] = T0
        v["dT"] = v["H"] - T0
        v["hot_start"] = bool(np.isfinite(v["dT"][1, 0]) and v["dT"][1, 0] > 300)
        v["full"] = v["record_s"] >= EVAL_END_S and not v["hot_start"]
        note.update(record_s=v["record_s"], used_s=int(v["t"][-1]), T0=np.round(T0, 1).tolist(),
                    hot_start=v["hot_start"], in_index=v["full"])
        qa["tests"][k] = note
    return qa


# ----------------------------------------------------------------------------- metrics

def cooling_event(dT):
    """Lower-level collapse (burner off / fire-source loss?): >= 2 of the 500-1500 mm levels cool at
    <= -5 °C/s (11-s mean, within +-10 s) and lose > 40 % of their rise within the next 60 s."""
    S = [smooth(dT[:, i]) for i in range(3)]
    R = [np.gradient(s) for s in S]
    n = len(S[0])
    for t in range(60, n - 1):
        hit = 0
        for s, r in zip(S, R):
            with np.errstate(invalid="ignore"):
                fast = np.nanmin(r[max(0, t - 10):t + 11]) <= COOL_RATE if np.isfinite(r[max(0, t - 10):t + 11]).any() else False
            if fast and np.isfinite(s[t - 5]) and s[min(t + 60, n - 1)] < 0.6 * s[t - 5]:
                hit += 1
        if hit >= COOL_LEVELS:
            return float(t)
    return np.nan


def height_metrics(v):
    rows = []
    dT, t = v["dT"], v["t"]
    for i, z in enumerate(HEIGHTS):
        x = dT[:, i]
        s = smooth(x)
        r = dict(height_mm=z, T0_C=v["T0"][i], peak_C=float(np.nanmax(v["H"][:, i])),
                 t_peak_s=float(t[np.nanargmax(s)]), dur_600K_s=int(np.sum(np.nan_to_num(x) >= THETA_REF)),
                 excess_600K_Kmin=float(np.nansum(np.clip(x - THETA_REF, 0, None)) / 60.0),
                 exposure_Kmin=float(np.nansum(np.clip(x, 0, None)) / 60.0))
        first = int(np.flatnonzero(np.isfinite(x))[0])
        for th in THRESHOLDS:
            ta = sustained_arrival(x, th)
            r[f"t{th}_s"] = np.nan if first > 0 and ta == first else ta   # censored by a sensor fault at the start
        r["first_valid_s"] = first
        rows.append(r)
    return rows


def hinge_fit(t, y, n):
    """y = b + a*(t - ti)_+^n, ti on a 2-s grid, (b, a) by least squares with a > 0."""
    best = None
    for ti in np.arange(t[0], t[-1] - 10, 2.0):
        X = np.column_stack([np.ones_like(t), np.clip(t - ti, 0, None) ** n])
        coef, *_ = np.linalg.lstsq(X, y, rcond=None)
        if coef[1] <= 0:
            continue
        sse = float(np.sum((X @ coef - y) ** 2))
        if best is None or sse < best[0]:
            best = (sse, ti, coef[0], coef[1])
    if best is None:
        return None
    sse, ti, b, a = best
    return dict(n=n, ti=float(ti), b=float(b), a=float(a), sse=sse,
                r2=float(1 - sse / np.sum((y - y.mean()) ** 2)), rmse=float(np.sqrt(sse / len(t))))


def growth_shape(v):
    """Shape of the 3000 mm rise: hinge power law on the 30-s sustained envelope, 60 s .. 95 % of its maximum."""
    E = sustained_envelope(v["dT"][:, TOP])
    t = v["t"]
    i95 = int(np.argmax(E >= 0.95 * E.max()))
    m = (t >= 60) & (t <= t[i95])
    if m.sum() < 60:
        return dict(E=E, window=None)
    best = None
    for n in np.arange(0.5, 4.01, 0.25):
        f = hinge_fit(t[m], E[m], n)
        if f and (best is None or f["sse"] < best["sse"]):
            best = f
    sq = hinge_fit(t[m], E[m], 2.0)
    return dict(E=E, window=(float(t[m][0]), float(t[m][-1])), best=best, t2=sq)


def index_3000(k, v):
    x = v["dT"][:, TOP]
    t = v["t"]
    arr = {th: sustained_arrival(x, th) for th in THRESHOLDS}
    t600 = arr[THETA_REF]
    alpha = THETA_REF / t600 ** 2 if np.isfinite(t600) else np.nan
    alpha500 = 500 / arr[500] ** 2 if np.isfinite(arr[500]) else np.nan
    E = sustained_envelope(x)
    gate = E >= THETA_REF
    with np.errstate(divide="ignore", invalid="ignore"):
        a_env = float(np.max(E[gate] / np.maximum(t[gate], 1.0) ** 2)) if gate.any() else np.nan
    zs = [z for z in HEIGHTS if z >= 2000]
    ta = [sustained_arrival(v["dT"][:, HEIGHTS.index(z)], FRONT_THETA) for z in zs]
    ok = np.isfinite(ta) & (np.nan_to_num(ta) > BURNER_PHASE_S)
    if ok.sum() >= 2 and np.ptp(np.array(ta)[ok]) >= SUSTAIN_S:
        v_up = float(np.polyfit(np.array(ta)[ok], np.array(zs)[ok], 1)[0])
        simult = False
    elif ok.sum() >= 2:
        v_up, simult = np.nan, True
    else:
        v_up, simult = np.nan, False
    early = x[: EARLY_S + 1]
    if v["max_only"]:
        n_tc = np.nan
    else:
        c = v["ch_clean"][: len(t), 15:18]
        n_tc = int(sum(np.isfinite(sustained_arrival(c[:, j] - c[0, j], THETA_REF)) for j in range(3)))
    return dict(
        test=k, material=v["material"], repeat=v["repeat"], record_s=v["record_s"], in_index=v["full"],
        T0_C=float(v["T0"][TOP]), peak_C=float(np.nanmax(v["H"][:, TOP])), t_peak_s=float(t[np.nanargmax(smooth(x))]),
        **{f"t{th}_s": arr[th] for th in THRESHOLDS},
        alpha600_K_s2=alpha, alpha500_K_s2=alpha500, alpha_env_K_s2=a_env,
        grade=grade_of(t600, v["record_s"]) if not v["hot_start"] else "n/a",
        rise_300_600_s=(t600 - arr[300]) if np.isfinite(t600) and np.isfinite(arr[300]) else np.nan,
        dur_600K_s=int(np.sum(np.nan_to_num(x) >= THETA_REF)), n_tc_600_of_3=n_tc,
        excess_600K_Kmin=float(np.nansum(np.clip(x - THETA_REF, 0, None)) / 60.0),
        exposure_Kmin=float(np.nansum(np.clip(x, 0, None)) / 60.0),
        v_front_2000_3000_mm_s=v_up, front_simultaneous=simult,
        early_max_dT_K=float(np.nanmax(early)), early_mean_dT_60_120_K=float(np.nanmean(x[60:EARLY_S + 1])),
        cooling_event_s=cooling_event(v["dT"]),
    )


def robustness(tests, order):
    """alpha = theta / t_theta^2 for theta x hold variants (max-of-3 signal); ranks and Kendall tau-b vs the primary."""
    variants, table, grades = [], {}, {}
    for th in ROBUST_THETA:
        for hold in ROBUST_HOLD:
            key = f"{th}K/{hold}s"
            vals, gr = [], []
            for k in order:
                v = tests[k]
                ta = sustained_arrival(v["dT"][:, TOP], th, hold)
                vals.append(th / ta ** 2 if np.isfinite(ta) else 0.0)
                gr.append(grade_of(ta, v["record_s"]))
            table[key] = np.array(vals)
            grades[key] = gr
            variants.append(key)
    pkey = f"{THETA_REF}K/{SUSTAIN_S}s"
    prim = table[pkey]
    taus = {key: float(stats.kendalltau(prim, val).statistic) for key, val in table.items()}
    ranks = {key: stats.rankdata(-val, method="min").astype(int) for key, val in table.items()}
    same_grade = {k: float(np.mean([grades[key][j] == grades[pkey][j] for key in variants])) for j, k in enumerate(order)}
    return dict(variants=variants, alpha=table, tau=taus, ranks=ranks, grades=grades, same_grade=same_grade, primary=pkey)


# ----------------------------------------------------------------------------- figures

def style():
    plt.rcParams.update({
        "figure.facecolor": SURFACE, "axes.facecolor": SURFACE, "savefig.facecolor": SURFACE,
        "axes.edgecolor": AXIS, "axes.linewidth": 0.8, "axes.labelcolor": INK2,
        "axes.titlesize": 10, "axes.titleweight": "bold", "axes.titlelocation": "left",
        "axes.titlecolor": INK, "axes.labelsize": 9.5, "font.size": 9.5,
        "xtick.color": AXIS, "ytick.color": AXIS, "xtick.labelcolor": INK2, "ytick.labelcolor": INK2,
        "xtick.labelsize": 8.5, "ytick.labelsize": 8.5,
        "axes.grid": True, "grid.color": GRID, "grid.linewidth": 0.6, "grid.linestyle": "-",
        "axes.axisbelow": True, "axes.spines.top": False, "axes.spines.right": False,
        "legend.frameon": False, "legend.fontsize": 8.5, "legend.labelcolor": INK2,
        "lines.solid_capstyle": "round", "lines.solid_joinstyle": "round",
    })


def ring(**kw):
    return dict(markeredgecolor=SURFACE, markeredgewidth=1.0, **kw)


def label_of(k, tests):
    v = tests[k]
    return f"{MATERIALS[v['material']][1]} #{v['repeat']}"


def fig_histories(tests, panels, idx, path):
    fig, axs = plt.subplots(5, 3, figsize=(13.5, 15.5), sharey=True)
    for r, row in enumerate(panels):
        for c in range(3):
            ax = axs[r, c]
            if c >= len(row):
                ax.axis("off")
                continue
            k = row[c]
            v = tests[k]
            t = v["t"]
            for i, z in enumerate(HEIGHTS):
                ax.plot(t, v["H"][:, i], color=HEIGHT_COLORS[i], lw=1.9 if i == TOP else 1.0, zorder=3 + i,
                        label=f"{z} mm")
            ax.set_title(f"{label_of(k, tests)}  ({k})", pad=14)
            ax.set_ylim(0, 1300)
            xmax = EVAL_END_S if v["record_s"] >= EVAL_END_S else 300
            ax.set_xlim(0, xmax)
            d = idx[k]
            if np.isfinite(d["t600_s"]):
                ax.plot([d["t600_s"]], [v["T0"][TOP] + THETA_REF], "o", color=INK, ms=6, zorder=20, **ring())
                txt = f"3000 mm: dT 600 K held 30 s at {d['t600_s']:.0f} s (dot)"
            elif v["record_s"] >= EVAL_END_S:
                txt = "3000 mm: dT 600 K not reached"
            else:
                txt = f"record ends at {v['record_s']} s"
            ax.text(0.99, 1.01, txt, transform=ax.transAxes, ha="right", va="bottom", fontsize=7.8, color=INK2)
            if np.isfinite(d["cooling_event_s"]):
                ax.axvline(d["cooling_event_s"], color=INK2, lw=0.8, ls=(0, (4, 3)), zorder=2)
                ax.text(d["cooling_event_s"] + (4 if xmax < 1000 else 15), 1270, f"lower levels\ncollapse {d['cooling_event_s']:.0f} s",
                        fontsize=7.4, color=INK2, va="top")
            if c == 0:
                ax.set_ylabel("Temperature (°C)")
            ax.set_xlabel("Time from ignition (s)")
    h, l = axs[0, 0].get_legend_handles_labels()
    fig.legend(h, l, loc="upper center", ncol=6, bbox_to_anchor=(0.5, 0.995), title="Height above floor (max of 3 TCs)",
               title_fontsize=8.5)
    fig.tight_layout(rect=(0, 0, 1, 0.975))
    fig.savefig(path, dpi=150)
    plt.close(fig)


def fig_fronts(tests, order, hm, path):
    fig, axs = plt.subplots(1, 3, figsize=(15, 5.6))
    for k in order:
        v = tests[k]
        col, mk = MATERIALS[v["material"]][2], MATERIALS[v["material"]][3]
        ls = REPEAT_STYLE[v["repeat"]]
        rows = hm[k]
        ta = np.array([r[f"t{FRONT_THETA}_s"] for r in rows])
        ok = np.isfinite(ta)
        axs[0].plot(ta[ok] / 60, np.array(HEIGHTS)[ok], ls=ls, color=col, lw=1.6, marker=mk, ms=6, **ring(),
                    label=label_of(k, tests))
        axs[1].plot([r["peak_C"] for r in rows], HEIGHTS, ls=ls, color=col, lw=1.6, marker=mk, ms=6, **ring())
        axs[2].plot([r["dur_600K_s"] / 60 for r in rows], HEIGHTS, ls=ls, color=col, lw=1.6, marker=mk, ms=6, **ring())
    axs[0].set_title(f"(a) Fire front: dT {FRONT_THETA} K held 30 s")
    axs[0].set_xlabel("Arrival time (min)")
    axs[0].set_xlim(0, 30)
    axs[1].set_title("(b) Peak temperature")
    axs[1].set_xlabel("Peak temperature (°C)")
    axs[2].set_title("(c) Time with dT >= 600 K")
    axs[2].set_xlabel("Duration (min)")
    for ax in axs:
        ax.set_ylabel("Height (mm)")
        ax.set_yticks(HEIGHTS)
        ax.set_ylim(300, 3200)
    fig.legend(*axs[0].get_legend_handles_labels(), loc="lower center", ncol=4, bbox_to_anchor=(0.5, -0.01))
    fig.tight_layout(rect=(0, 0.1, 1, 1))
    fig.savefig(path, dpi=150)
    plt.close(fig)


def fig_at2(tests, order, short, idx, path):
    fig = plt.figure(figsize=(16, 10.8))
    gs = fig.add_gridspec(2, 5, height_ratios=[1, 1.05], hspace=0.42, wspace=0.08)
    tt = np.linspace(0, EVAL_END_S, 400)
    for c, m in enumerate(MATERIALS):
        ax = fig.add_subplot(gs[0, c])
        for tg in GRADE_TG:
            a = THETA_REF / tg ** 2
            ax.plot(tt, a * tt ** 2, color=MUTED, lw=0.9, ls=(0, (5, 3)), zorder=2)
            if tg < EVAL_END_S:
                tx = tg * np.sqrt(1190 / THETA_REF)
                ax.text(tx + 25, 1190, f"{tg / 60:g} min", fontsize=7.4, color=INK2, ha="left", va="top")
            else:
                ax.text(EVAL_END_S - 15, THETA_REF - 25, f"{tg / 60:g} min", fontsize=7.4, color=INK2, ha="right", va="top")
        ax.axhline(THETA_REF, color=INK2, lw=0.8, zorder=2)
        for k in [k for k in order + short if tests[k]["material"] == m]:
            v = tests[k]
            col, mk = MATERIALS[m][2], MATERIALS[m][3]
            ax.plot(v["t"], smooth(v["dT"][:, TOP]), color=col, ls=REPEAT_STYLE[v["repeat"]], lw=1.5, zorder=4,
                    label=f"#{v['repeat']}")
            if np.isfinite(idx[k]["t600_s"]):
                ax.plot([idx[k]["t600_s"]], [THETA_REF], marker=mk, color=col, ms=8, zorder=6, ls="none", **ring())
        ax.set_xlim(0, EVAL_END_S)
        ax.set_ylim(0, 1200)
        ax.set_xticks([0, 600, 1200, 1800])
        ax.set_title(MATERIALS[m][1])
        ax.legend(loc="upper left", fontsize=7.6)
        recs = [tests[k]["record_s"] for k in short if tests[k]["material"] == m]
        if recs:
            ax.text(0.97, 0.06, "records end at\n" + " / ".join(f"{r} s" for r in recs) + " (not graded)",
                    transform=ax.transAxes, ha="right", va="bottom", fontsize=7.8, color=INK2)
        ax.set_xlabel("Time from ignition (s)")
        if c == 0:
            ax.set_ylabel("dT at 3000 mm (K, 11-s mean)")
        else:
            ax.tick_params(labelleft=False)
    fig.text(0.06, 0.955, "(a) 3000 mm temperature rise against equivalent t² curves a·t² that reach 600 K at 7.5 / 15 / 30 min "
             "(dashed = grade boundaries; marker = t600, dT >= 600 K held 30 s; #1 solid, #2 dashed, #3 dotted)",
             fontsize=10, fontweight="bold", color=INK)
    ax = fig.add_subplot(gs[1, :])
    keys = order + short
    x = np.arange(len(keys))
    lo, hi = 1e-4, 1e-2
    bounds = [hi] + [THETA_REF / tg ** 2 for tg in GRADE_TG] + [lo]
    for g, (b1, b0) in zip(GRADE_NAMES, zip(bounds[:-1], bounds[1:])):
        ax.axhspan(b0, b1, color=GRADE_FILL[g], zorder=0, lw=0)
        ax.text(-0.5, np.sqrt(b0 * b1), f"grade {g}", ha="left", va="center", fontsize=8.5, color=INK2)
    for xx, k in zip(x, keys):
        d = idx[k]
        m = tests[k]["material"]
        if np.isfinite(d["alpha600_K_s2"]):
            ax.plot([xx], [d["alpha600_K_s2"]], marker=MATERIALS[m][3], color=MATERIALS[m][2], ms=11, ls="none",
                    zorder=4, **ring())
            ax.text(xx, d["alpha600_K_s2"] * 1.25, f"{d['alpha600_K_s2']:.2e}\nt600 {d['t600_s']:.0f} s", ha="center",
                    va="bottom", fontsize=7.8, color=INK)
        else:
            txt = "not reached\nin 30 min" if d["record_s"] >= EVAL_END_S else f"n/a\nrecord {d['record_s']} s"
            ax.text(xx, 1.25e-4, txt, ha="center", va="bottom", fontsize=7.8, color=INK2)
    ax.set_xticks(x)
    ax.set_xticklabels([label_of(k, tests).replace(" #", "\n#") for k in keys], fontsize=8)
    ax.set_yscale("log")
    ax.set_ylim(lo, hi)
    ax.set_xlim(-0.6, len(keys) - 0.4)
    ax.grid(axis="x", visible=False)
    ax.set_ylabel("a600 = 600 K / t600²  (K/s²)")
    ax.set_title("(b) Equivalent t² growth coefficient a600 at 3000 mm, by test")
    fig.subplots_adjust(left=0.06, right=0.99, top=0.92, bottom=0.08)
    fig.savefig(path, dpi=150)
    plt.close(fig)


def fig_risk_map(tests, order, idx, path):
    fig, ax = plt.subplots(figsize=(9.5, 6.4))
    lo, hi = 1e-4, 1e-2
    nr = 1.35e-4
    bounds = [hi] + [THETA_REF / tg ** 2 for tg in GRADE_TG] + [lo]
    for g, (b1, b0) in zip(GRADE_NAMES, zip(bounds[:-1], bounds[1:])):
        ax.axvspan(b0, b1, color=GRADE_FILL[g], zorder=0, lw=0)
        ax.text(np.sqrt(b0 * b1), 0.98, f"grade {g}", transform=ax.get_xaxis_transform(), ha="center", va="top",
                fontsize=8.5, color=INK2)
    offsets = {}
    for k in order:
        d = idx[k]
        v = tests[k]
        col, mk = MATERIALS[v["material"]][2], MATERIALS[v["material"]][3]
        x = d["alpha600_K_s2"] if np.isfinite(d["alpha600_K_s2"]) else nr
        yv = d["excess_600K_Kmin"]
        ax.plot([x], [yv], marker=mk, color=col, ms=11, ls="none", zorder=5, **ring())
        dy = offsets.get(round(np.log10(x), 1), 0)
        offsets[round(np.log10(x), 1)] = dy + 1
        ax.annotate(f"{label_of(k, tests)}\npeak {d['peak_C']:.0f} °C, {d['dur_600K_s'] / 60:.1f} min > 600 K",
                    (x, yv), xytext=(9, 6 - 22 * dy), textcoords="offset points", fontsize=7.6, color=INK)
    ax.set_xscale("log")
    ax.set_xlim(lo, hi)
    ymax = max(idx[k]["excess_600K_Kmin"] for k in order)
    ax.set_ylim(-0.05 * ymax, 1.18 * ymax)
    ax.set_xticks([nr, 1.85e-4, 7.41e-4, 2.96e-3])
    ax.set_xticklabels(["not\nreached", "1.85e-4\n(30 min)", "7.41e-4\n(15 min)", "2.96e-3\n(7.5 min)"])
    ax.set_xlabel("Speed: a600 = 600 K / t600²  (K/s², log scale; t600 in brackets)")
    ax.set_ylabel("Severity: excess exposure  ∫(dT − 600 K)+ dt  (K·min, 0–30 min)")
    ax.set_title("Fire-spread risk map at 3000 mm (speed x severity)")
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)


def fig_growth_shape(tests, order, shapes, path):
    fig, axs = plt.subplots(2, 4, figsize=(15, 7.2), sharey=True)
    for ax, k in zip(axs.ravel(), order):
        v = tests[k]
        sh = shapes[k]
        col = MATERIALS[v["material"]][2]
        ax.plot(v["t"], v["dT"][:, TOP], color=GRID, lw=0.7, zorder=2)
        ax.plot(v["t"], sh["E"], color=col, lw=1.8, zorder=3, label="30-s sustained envelope")
        if sh["window"]:
            t0, t1 = sh["window"]
            tt = np.linspace(t0, t1, 300)
            for f, ls, c, name in ((sh["t2"], (0, (5, 3)), INK, "t²"), (sh["best"], "-", INK2, "best n")):
                if f:
                    ax.plot(tt, f["b"] + f["a"] * np.clip(tt - f["ti"], 0, None) ** f["n"], ls=ls, color=c, lw=1.1,
                            zorder=4, label=f"{name}: n={f['n']:.2f}, R²={f['r2']:.2f}")
            ax.axvspan(t0, t1, color=BAND, zorder=0, lw=0)
        ax.set_title(label_of(k, tests))
        ax.set_xlim(0, EVAL_END_S)
        ax.set_ylim(0, 1200)
        ax.legend(loc="upper left", fontsize=7.2)
        ax.set_xlabel("Time (s)")
    for ax in axs[:, 0]:
        ax.set_ylabel("dT at 3000 mm (K)")
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)


def fig_robustness(tests, order, rob, path):
    fig, axs = plt.subplots(1, 2, figsize=(14.5, 5.6), gridspec_kw=dict(width_ratios=[1, 1.1]))
    ax = axs[0]
    for j, k in enumerate(order):
        v = tests[k]
        col, mk = MATERIALS[v["material"]][2], MATERIALS[v["material"]][3]
        ys = [rob["alpha"][f"{th}K/{SUSTAIN_S}s"][j] for th in ROBUST_THETA]
        ys = [y if y > 0 else 6e-5 for y in ys]
        ax.plot(ROBUST_THETA, ys, color=col, ls=REPEAT_STYLE[v["repeat"]], marker=mk, ms=7, lw=1.5, **ring(),
                label=label_of(k, tests))
    ax.set_yscale("log")
    ax.set_ylim(5e-5, 2e-2)
    ax.axhline(6e-5, color=AXIS, lw=0.6)
    ax.text(ROBUST_THETA[0], 6.6e-5, "not reached", fontsize=7.6, color=INK2)
    ax.set_xticks(ROBUST_THETA)
    ax.set_xlabel("Criterion dT (K), held 30 s")
    ax.set_ylabel("a = dT / t²  (K/s²)")
    ax.set_title("(a) Equivalent t² coefficient vs. criterion level")
    ax.legend(fontsize=7.6, loc="upper right", ncol=2)
    ax = axs[1]
    R = np.array([rob["ranks"][key] for key in rob["variants"]])  # variants x tests
    prim = rob["ranks"][rob["primary"]]
    y = np.arange(len(order))[::-1]
    for yy, j in zip(y, range(len(order))):
        col = MATERIALS[tests[order[j]]["material"]][2]
        vals = R[:, j]
        cnt = np.bincount(vals, minlength=len(order) + 1)
        for rank, c in enumerate(cnt):
            if c:
                ax.scatter([rank], [yy], s=18 + 9 * c, color=col, alpha=0.85, edgecolor=SURFACE, linewidth=1.0, zorder=3)
        ax.plot([prim[j]], [yy], marker="|", color=INK, ms=16, mew=1.6, zorder=4)
    ax.set_yticks(y)
    ax.set_yticklabels([label_of(k, tests) for k in order])
    ax.set_xticks(range(1, len(order) + 1))
    ax.set_xlim(0.5, len(order) + 0.5)
    ax.set_xlabel("Rank (1 = most hazardous); dot area = number of variants, | = primary definition")
    taus = np.array(list(rob["tau"].values()))
    ax.set_title(f"(b) Rank across {len(rob['variants'])} definitions  (Kendall tau vs primary: "
                 f"min {taus.min():.2f}, median {np.median(taus):.2f})")
    ax.grid(axis="y", visible=False)
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)


# ----------------------------------------------------------------------------- output

def clean(o):
    if isinstance(o, dict):
        return {str(k): clean(v) for k, v in o.items()}
    if isinstance(o, (list, tuple)):
        return [clean(v) for v in o]
    if isinstance(o, (np.floating, float)):
        return None if not np.isfinite(o) else round(float(o), 6)
    if isinstance(o, (np.integer,)):
        return int(o)
    if isinstance(o, np.bool_):
        return bool(o)
    if isinstance(o, np.ndarray):
        return clean(o.tolist())
    return o


def write_csv(path, rows):
    rows = [clean(r) for r in rows]
    keys = list(dict.fromkeys(k for r in rows for k in r))
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=keys)
        w.writeheader()
        for r in rows:
            w.writerow({k: "" if r.get(k) is None else r.get(k) for k in keys})


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("xlsx", nargs="*", type=Path,
                    default=[HERE / "data" / n for n in ("EPS.xlsx", "PF.xlsx", "PIR_SC_2_max.xlsx")])
    args = ap.parse_args()
    figs, res = HERE / "figures", HERE / "results"
    figs.mkdir(exist_ok=True)
    res.mkdir(exist_ok=True)

    tests = read_tests(args.xlsx)
    qa = prepare(tests)
    mat_order = list(MATERIALS)
    keys = sorted(tests, key=lambda k: (mat_order.index(tests[k]["material"]), tests[k]["repeat"]))
    order = [k for k in keys if tests[k]["full"]]                      # 30-min tests used for the index
    short = [k for k in keys if not tests[k]["full"] and not tests[k]["hot_start"]]   # truncated records
    excluded = [k for k in keys if tests[k]["hot_start"]]

    hm = {k: height_metrics(tests[k]) for k in keys}
    idx = {k: index_3000(k, tests[k]) for k in keys}
    shapes = {k: growth_shape(tests[k]) for k in order}
    rob = robustness(tests, order)

    style()
    panels = [[k for k in keys if tests[k]["material"] == m and k not in excluded] for m in mat_order]
    fig_histories(tests, panels, idx, figs / "fig1_temperature_histories.png")
    fig_fronts(tests, order, hm, figs / "fig2_fire_front_and_height_trends.png")
    fig_at2(tests, order, short, idx, figs / "fig3_at2_index_3000mm.png")
    fig_risk_map(tests, order, idx, figs / "fig4_risk_map.png")
    fig_growth_shape(tests, order, shapes, figs / "fig5_growth_shape_3000mm.png")
    fig_robustness(tests, order, rob, figs / "fig6_robustness.png")

    write_csv(res / "index_3000mm.csv", [idx[k] for k in keys])
    write_csv(res / "height_metrics.csv", [dict(test=k, **r) for k in keys for r in hm[k]])
    shape_rows = []
    for k in order:
        sh = shapes[k]
        if sh["window"]:
            shape_rows.append(dict(test=k, window_s=f"{sh['window'][0]:.0f}-{sh['window'][1]:.0f}",
                                   n_best=sh["best"]["n"], ti_best_s=sh["best"]["ti"], r2_best=sh["best"]["r2"],
                                   rmse_best_K=sh["best"]["rmse"], ti_t2_s=sh["t2"]["ti"], a_t2_K_s2=sh["t2"]["a"],
                                   r2_t2=sh["t2"]["r2"], rmse_t2_K=sh["t2"]["rmse"]))
    write_csv(res / "growth_shape_3000mm.csv", shape_rows)
    rob_rows = []
    for key in rob["variants"]:
        r = {"variant": key, "kendall_tau_vs_primary": rob["tau"][key]}
        for j, k in enumerate(order):
            r[f"alpha_{k}"] = rob["alpha"][key][j]
            r[f"rank_{k}"] = int(rob["ranks"][key][j])
            r[f"grade_{k}"] = rob["grades"][key][j]
        rob_rows.append(r)
    write_csv(res / "robustness.csv", rob_rows)

    mat_rows = []
    for m in mat_order:
        ks = [k for k in order if tests[k]["material"] == m]
        if not ks:
            ks_short = [k for k in short if tests[k]["material"] == m]
            mat_rows.append(dict(material=m, label_ko=MATERIALS[m][0], n_index=0, tests=" ".join(ks_short),
                                 note="records shorter than 30 min: not graded"))
            continue
        t6 = np.array([idx[k]["t600_s"] for k in ks])
        al = np.array([idx[k]["alpha600_K_s2"] for k in ks])
        mat_rows.append(dict(
            material=m, label_ko=MATERIALS[m][0], n_index=len(ks), tests=" ".join(ks),
            t600_reached=f"{int(np.isfinite(t6).sum())}/{len(ks)}",
            t600_min_s=np.nanmin(t6) if np.isfinite(t6).any() else np.nan,
            t600_max_s=np.nanmax(t6) if np.isfinite(t6).any() else np.nan,
            alpha600_max=np.nanmax(al) if np.isfinite(al).any() else np.nan,
            alpha600_min=np.nanmin(al) if np.isfinite(al).all() else 0.0,
            grades=" / ".join(idx[k]["grade"] for k in ks),
            peak_mean_C=np.mean([idx[k]["peak_C"] for k in ks]),
            excess_600K_mean_Kmin=np.mean([idx[k]["excess_600K_Kmin"] for k in ks]),
            dur_600K_mean_s=np.mean([idx[k]["dur_600K_s"] for k in ks]),
        ))
    write_csv(res / "material_summary.csv", mat_rows)

    summary = dict(
        definitions=dict(
            signal="max of the three thermocouples per height; dT = T - T(t=0)",
            t_theta=f"start of the first {SUSTAIN_S}-s window with dT >= theta throughout (from ignition, t = 0)",
            alpha600="600 K / t600^2  [K/s^2]  (NFPA-style: alpha = reference / t_g^2)",
            grades={g: f"t600 <= {tg} s (alpha >= {THETA_REF / tg ** 2:.3e})" for g, tg in zip(GRADE_NAMES[:3], GRADE_TG)}
            | {"I": f"not reached within {EVAL_END_S} s"},
            evaluation_window_s=EVAL_END_S,
        ),
        tests_in_index=order, short_records=short, excluded=excluded,
        qa=qa, index_3000mm={k: idx[k] for k in keys}, materials=mat_rows,
        robustness=dict(tau=rob["tau"], tau_min=min(rob["tau"].values()),
                        tau_median=float(np.median(list(rob["tau"].values()))), same_grade_share=rob["same_grade"]),
        growth_shape=shape_rows,
    )
    (res / "summary.json").write_text(json.dumps(clean(summary), ensure_ascii=False, indent=1), encoding="utf-8")

    print("splices:", json.dumps(clean(qa["splices"]), ensure_ascii=False))
    for k in keys:
        d = idx[k]
        print(f"{k:9s} rec={d['record_s']:5d} t300={d['t300_s']:6.0f} t500={d['t500_s']:6.0f} t600={d['t600_s']:6.0f} "
              f"a600={d['alpha600_K_s2']:.2e} grade={d['grade']:3s} peak={d['peak_C']:6.0f} D600={d['dur_600K_s']:4d} "
              f"H600={d['excess_600K_Kmin']:7.0f} v={d['v_front_2000_3000_mm_s']:.2f} sim={d['front_simultaneous']} "
              f"cool={d['cooling_event_s']}")
    print("tau min/median:", summary["robustness"]["tau_min"], summary["robustness"]["tau_median"])
    print("same grade share:", {k: round(x, 2) for k, x in rob["same_grade"].items()})


if __name__ == "__main__":
    main()
