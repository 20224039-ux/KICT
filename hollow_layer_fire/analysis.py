#!/usr/bin/env python3
"""중공층(PIR) 화재시험 – 높이별 온도이력 · 화재확산 · 성장(at²) 경향 분석.

Hollow-layer (cavity) fire test, PIR specimen 2: 1 Hz temperature logs at
500–3000 mm above the fire source (sheet 'sc_2 max', one column per height).

    python analysis.py                      # reads data/PIR_2_max.xlsx
    python analysis.py --xlsx other.xlsx    # same layout: time column 's' + PIR_<n>_<height> columns

Writes figures/*.png, results/*.csv and results/summary.json next to this file.
"""
from __future__ import annotations

import argparse
import csv
import json
import re
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.patheffects as pe  # noqa: E402
import matplotlib.pyplot as plt  # noqa: E402
import matplotlib.ticker  # noqa: E402
import numpy as np  # noqa: E402
import openpyxl  # noqa: E402
from scipy import optimize, stats  # noqa: E402

HERE = Path(__file__).resolve().parent
SHEET = "sc_2 max"
END_OF_RECORD_S = 1800     # from 1807 s every channel falls to ~35 °C within 5–8 s (end of test) -> excluded
PLATEAU_S = (100, 340)     # phase II: quasi-steady plateau under the fire source
ONSET_RISE_C = 30          # phase III starts when the running max exceeds plateau + 30 °C
MIN_GROWTH_RISE_C = 150    # smaller phase-III rise = level already inside the flame (no growth fit)
THRESHOLDS = (300, 400, 500, 600, 700, 800, 900, 1000)
ISOTHERMS = (500, 600, 800)
SUSTAIN_S = 30             # "for at least 30 s" (KS F 8414 / BS 8414 style)
BURNOUT_C = 600            # burn-out front: first drop below this after the peak
EVENT_AFTER_S = 1100       # search window for the simultaneous late cooling
EVENT_RATE = -2.0          # °C/s (11-s moving average)
PROFILE_TIMES = (60, 300, 600, 900, 1200)
POST_EVENT_TIMES = (1230, 1290, 1350)
DWELL_S = 30               # hottest level must hold this long to count as a change

# The earlier Excel chart (test.xlsx, chart3): exponential trendlines with 'Set Intercept'
EXCEL_A = {2000: 350.0, 2500: 250.0, 3000: 150.0}
EXCEL_PRINTED_K = {2000: 0.0012, 2500: 0.0013, 3000: 0.0014}
EXCEL_LEVELS = np.array([400.0, 500.0, 600.0, 700.0, 800.0])
UNIFIED_H_REF, UNIFIED_T_REF = 2000.0, 400.0

# chart tokens (dataviz reference palette, light surface; ramps validated as ordinal)
SURFACE, INK, INK2, MUTED = "#fcfcfb", "#0b0b0b", "#52514e", "#898781"
GRID, AXIS, BAND = "#e1e0d9", "#c3c2b7", "#efeee9"
HEIGHT_COLORS = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#4a3aa7", "#e34948"]
HEIGHT_MARKERS = ["o", "s", "^", "D", "p", "h", "v", "P"]
BLUE_RAMP = ["#86b6ef", "#5598e7", "#2a78d6", "#1c5cab", "#104281"]
ORANGE_RAMP = ["#ef8a5c", "#d9582a", "#9c3d17"]
ISO_COLOR = {500: BLUE_RAMP[0], 600: BLUE_RAMP[2], 700: BLUE_RAMP[3], 800: BLUE_RAMP[4], 900: INK}
PEAK_COLOR, BURN_COLOR = "#eb6834", INK2
trapezoid = getattr(np, "trapezoid", None) or np.trapz


def rms(x):
    return float(np.sqrt(np.mean(np.square(x))))


def smooth(y, w=11):
    """Centred moving average that ignores NaN."""
    ok = np.isfinite(y)
    k = np.ones(w)
    num = np.convolve(np.pad(np.where(ok, y, 0.0), w // 2, mode="edge"), k, mode="valid")
    den = np.convolve(np.pad(ok.astype(float), w // 2, mode="edge"), k, mode="valid")
    with np.errstate(invalid="ignore", divide="ignore"):
        return np.where(den > w / 2, num / den, np.nan)


# ----------------------------------------------------------------------------- data

def read_logs(path: Path):
    """Time vector and {height_mm: temperature} from the first block of PIR_<n>_<height> columns."""
    ws = openpyxl.load_workbook(path, data_only=True)[SHEET]
    header = [(c.column_letter, c.value) for c in ws[1]]
    tcol = next(col for col, v in header if v == "s")
    cols, started = {}, False
    for col, v in header[[c for c, _ in header].index(tcol) + 1:]:
        m = re.fullmatch(r"PIR_\d+_(\d+)", str(v or ""))
        if m:
            cols[int(m.group(1))], started = col, True
        elif started:
            break
    rows = range(2, ws.max_row + 1)
    as_float = lambda v: float(v) if isinstance(v, (int, float)) else np.nan  # noqa: E731
    t = np.array([as_float(ws[f"{tcol}{r}"].value) for r in rows])
    logs = {h: np.array([as_float(ws[f"{c}{r}"].value) for r in rows]) for h, c in sorted(cols.items())}
    keep = np.isfinite(t) & (t <= END_OF_RECORD_S)
    return t[keep], {h: T[keep] for h, T in logs.items()}


# ----------------------------------------------------------------------------- per-level metrics

def first_arrival(t, T, thr):
    idx = np.flatnonzero(T >= thr)
    return float(t[idx[0]]) if idx.size else np.nan


def sustained_arrival(t, T, thr, hold=SUSTAIN_S):
    """Start of the first period with T >= thr lasting at least `hold` seconds."""
    edges = np.diff(np.concatenate(([0], (T >= thr).astype(int), [0])))
    for s, e in zip(np.flatnonzero(edges == 1), np.flatnonzero(edges == -1)):
        if t[e - 1] - t[s] >= hold:
            return float(t[s])
    return np.nan


def phase_metrics(t, T):
    dt = float(np.median(np.diff(t)))
    in_plateau = (t >= PLATEAU_S[0]) & (t <= PLATEAU_S[1])
    plateau = float(T[in_plateau].mean())
    env = np.maximum.accumulate(T)
    i_pk = int(np.argmax(T))
    rate = np.gradient(smooth(T), t)
    after_peak = np.arange(T.size) > i_pk
    burn = np.flatnonzero(after_peak & (T < BURNOUT_C))
    onset = np.flatnonzero((t > PLATEAU_S[1]) & (env > plateau + ONSET_RISE_C))
    return {
        "T0": float(T[0]),
        "t90_initial": float(t[np.flatnonzero(T >= T[0] + 0.9 * (plateau - T[0]))[0]]),
        "rate_initial": float(rate[t <= 80].max()),
        "plateau": plateau,
        "plateau_sd": float(T[in_plateau].std()),
        "dT_plateau": plateau - float(T[0]),
        "onset_III": float(t[onset[0]]) if onset.size else np.nan,
        "T_peak": float(T[i_pk]),
        "t_peak": float(t[i_pk]),
        "rise_III": float(T[i_pk]) - plateau,
        **{f"dur_{thr}": float(np.sum(T >= thr) * dt) for thr in (300, 500, 600, 800)},
        "dose_Cmin": float(trapezoid(T - T[0], t) / 60),
        "burnout": float(t[burn[0]]) if burn.size else np.nan,
        "t_dT600_30s": sustained_arrival(t, T, T[0] + 600.0),
    }


def late_cooling_event(t, logs):
    """First time after EVENT_AFTER_S when at least half of the levels cool faster than EVENT_RATE."""
    rates = {h: np.gradient(smooth(T), t) for h, T in logs.items()}
    cooling = np.sum([r < EVENT_RATE for r in rates.values()], axis=0)
    i = int(np.flatnonzero((t > EVENT_AFTER_S) & (cooling >= len(logs) / 2))[0])
    after = t >= t[i]
    return {
        "t_event": float(t[i]),
        "cooling_levels": [h for h, r in rates.items() if r[i] < EVENT_RATE],
        "max_heating_after": {h: float(r[after].max()) for h, r in rates.items()},
        "t_max_heating_after": {h: float(t[after][np.argmax(r[after])]) for h, r in rates.items()},
        "max_cooling_after": {h: float(r[after].min()) for h, r in rates.items()},
    }


# ----------------------------------------------------------------------------- vertical structure

def isotherm_height(H, M, thr):
    """Highest level with T >= thr, interpolated linearly toward the next level (NaN if below the lowest)."""
    z = np.full(M.shape[1], np.nan)
    for i in range(M.shape[1]):
        col = M[:, i]
        idx = np.flatnonzero(col >= thr)
        if idx.size:
            k = idx.max()
            z[i] = H[k] if k == len(H) - 1 else H[k] + (H[k + 1] - H[k]) * (col[k] - thr) / (col[k] - col[k + 1])
    return z


def hottest_level_changes(t, H, M_smooth):
    """Times at which the hottest level changes (a new level must stay hottest for DWELL_S)."""
    arg = np.argmax(M_smooth, axis=0)
    need = int(DWELL_S / float(np.median(np.diff(t))))
    changes, current = [], arg[0]
    for i in range(1, arg.size):
        if arg[i] != current and np.all(arg[i:i + need] == arg[i]):
            changes.append({"t_s": float(t[i]), "from_mm": int(H[current]), "to_mm": int(H[arg[i]])})
            current = arg[i]
    return H[arg], changes


# ----------------------------------------------------------------------------- fronts

def segments(arr: dict):
    """Consecutive-level velocities for one front {height: time}."""
    hs = [h for h in sorted(arr) if np.isfinite(arr[h])]
    out = []
    for h1, h2 in zip(hs, hs[1:]):
        dt = arr[h2] - arr[h1]
        out.append({"from_mm": h1, "to_mm": h2, "t_from_s": arr[h1], "t_to_s": arr[h2],
                    "dt_s": dt, "v_mm_s": (h2 - h1) / dt if dt > 0 else np.inf})
    return out


def front_fit(arr: dict):
    """Linear z = z0 + v t and log-log z ∝ t^m over all levels reached."""
    hs = [h for h in sorted(arr) if np.isfinite(arr[h])]
    z, ta = np.array(hs, float), np.array([arr[h] for h in hs])
    lin, lg = stats.linregress(ta, z), stats.linregress(np.log(ta), np.log(z))
    return {"levels": len(hs), "v_mm_s": float(lin.slope), "z0_mm": float(lin.intercept), "r2_linear": float(lin.rvalue ** 2),
            "m": float(lg.slope), "r2_loglog": float(lg.rvalue ** 2),
            "p_if_L~Q^0.4": float(2.5 * lg.slope), "p_if_L~Q'^(2/3)": float(1.5 * lg.slope)}


# ----------------------------------------------------------------------------- growth laws

def power_fit(x, y):
    """Least-squares y = a x^n (x > 0), started from the log-log line."""
    ok = y > 0
    n0, lna0 = np.polyfit(np.log(x[ok]), np.log(y[ok]), 1)
    (a, n), _ = optimize.curve_fit(lambda x, a, n: a * x ** n, x, y, p0=(np.exp(lna0), n0), maxfev=20000)
    return float(a), float(n), rms(y - a * x ** n)


def growth_models(t, T, pm):
    """Phase III (onset -> peak) growth of the running-max envelope above the plateau."""
    env = np.maximum.accumulate(T)
    sel = (t > pm["onset_III"]) & (t <= pm["t_peak"])
    x, y = t[sel] - pm["onset_III"], env[sel] - pm["plateau"]
    a, n, rmse_pow = power_fit(x, y)
    a2 = float(np.sum(y * x ** 2) / np.sum(x ** 4))
    r1 = float(np.sum(y * x) / np.sum(x ** 2))
    best = None  # sensitivity: let the time origin float between ignition and the onset
    for ts in np.arange(0.0, pm["onset_III"], 5.0):
        try:
            fit = power_fit(t[sel] - ts, y)
        except RuntimeError:
            continue
        if best is None or fit[2] < best[1][2]:
            best = (ts, fit)
    return {"x": x, "y": y, "n": n, "a_pow": a, "rmse_pow": rmse_pow,
            "a_at2": a2, "rmse_at2": rms(y - a2 * x ** 2), "rate_lin": r1, "rmse_lin": rms(y - r1 * x),
            "free_origin_ts": float(best[0]), "free_origin_n": best[1][1], "free_origin_rmse": best[1][2]}


def arrival_models(t_arr, T_lev, T_amb, A_fixed, k_printed):
    """Models for the time needed to reach each threshold (thresholds fixed, time measured)."""
    L, s = np.log(T_lev), np.sqrt(T_lev - T_amb)
    k_fix = float(np.sum(t_arr * np.log(T_lev / A_fixed)) / np.sum(t_arr ** 2))  # Excel 'Set Intercept'
    r_exp, r_at2, r_lin = stats.linregress(L, t_arr), stats.linregress(s, t_arr), stats.linregress(T_lev, t_arr)
    m = {
        "excel_trendline": {"A": A_fixed, "k": k_fix, "t_hat": np.log(T_lev / A_fixed) / k_fix},
        "excel_printed": {"A": A_fixed, "k": k_printed, "t_hat": np.log(T_lev / A_fixed) / k_printed},
        "exponential": {"A": float(np.exp(-r_exp.intercept / r_exp.slope)), "k": float(1 / r_exp.slope),
                        "t_hat": r_exp.intercept + r_exp.slope * L},
        "at2": {"a": float(1 / r_at2.slope ** 2), "t0": float(r_at2.intercept),
                "t_hat": r_at2.intercept + r_at2.slope * s},
        "linear": {"rate": float(1 / r_lin.slope), "t_hat": r_lin.intercept + r_lin.slope * T_lev},
    }
    for v in m.values():
        res = t_arr - v["t_hat"]
        v["rmse_t"], v["max_abs_t"] = rms(res), float(np.max(np.abs(res)))
    return m


def unified_model(chart):
    """t = tau0 + s (h - 2000) + b ln(T/400): one heating curve shifted in time with height."""
    rows = [(h, T, tt) for h, (ts, Ts) in chart.items() for tt, T in zip(ts, Ts)]
    h, T, y = (np.array(c, float) for c in zip(*rows))
    L = np.log(T / UNIFIED_T_REF)
    dummies = np.column_stack([(h == hh).astype(float) for hh in sorted(chart)])

    def sse(X):
        beta = np.linalg.lstsq(X, y, rcond=None)[0]
        return beta, float(np.sum((y - X @ beta) ** 2))

    X3 = np.column_stack([np.ones_like(y), h - UNIFIED_H_REF, L])
    beta, sse3 = sse(X3)
    _, sse4 = sse(np.column_stack([dummies, L]))
    _, sse6 = sse(np.column_stack([dummies, dummies * L[:, None]]))
    n, dof = y.size, y.size - 3
    se = np.sqrt(np.diag(sse3 / dof * np.linalg.inv(X3.T @ X3)))
    q = stats.t.ppf(0.975, dof)
    f_slope = ((sse4 - sse6) / 2) / (sse6 / (n - 6))
    f_linear = (sse3 - sse4) / (sse4 / (n - 4))
    tau0, s, b = beta
    return {
        "tau0_s": float(tau0), "delay_s_per_mm": float(s), "b_s": float(b), "k": float(1 / b),
        "velocity_mm_s": float(1 / s), "velocity_ci95": [float(1 / (s + q * se[1])), float(1 / (s - q * se[1]))],
        "k_ci95": [float(1 / (b + q * se[2])), float(1 / (b - q * se[2]))],
        "rmse_t": float(np.sqrt(sse3 / n)),
        "F_separate_k": float(f_slope), "p_separate_k": float(stats.f.sf(f_slope, 2, n - 6)),
        "F_nonlinear_delay": float(f_linear), "p_nonlinear_delay": float(stats.f.sf(f_linear, 1, n - 4)),
        "A_equiv": {int(hh): float(UNIFIED_T_REF * np.exp(-(tau0 + s * (hh - UNIFIED_H_REF)) / b)) for hh in sorted(chart)},
    }


def unified_time(u, h, T):
    return u["tau0_s"] + u["delay_s_per_mm"] * (h - UNIFIED_H_REF) + u["b_s"] * np.log(T / UNIFIED_T_REF)


# ----------------------------------------------------------------------------- figures

def style():
    plt.rcParams.update({
        "figure.facecolor": SURFACE, "axes.facecolor": SURFACE, "savefig.facecolor": SURFACE,
        "axes.edgecolor": AXIS, "axes.linewidth": 0.8, "axes.labelcolor": INK2,
        "axes.titlesize": 10.5, "axes.titleweight": "bold", "axes.titlelocation": "left",
        "axes.titlecolor": INK, "axes.labelsize": 9.5, "font.size": 9.5,
        "xtick.color": AXIS, "ytick.color": AXIS, "xtick.labelcolor": INK2, "ytick.labelcolor": INK2,
        "xtick.labelsize": 8.5, "ytick.labelsize": 8.5,
        "axes.grid": True, "grid.color": GRID, "grid.linewidth": 0.6, "grid.linestyle": "-",
        "axes.axisbelow": True, "axes.spines.top": False, "axes.spines.right": False,
        "legend.frameon": False, "legend.fontsize": 8.5, "legend.labelcolor": INK2,
        "lines.solid_capstyle": "round", "lines.solid_joinstyle": "round",
    })


def ring(**kw):
    """Marker with a 2px surface ring."""
    return dict(markeredgecolor=SURFACE, markeredgewidth=1.0, **kw)


def event_line(ax, ev, y, text=True):
    ax.axvline(ev["t_event"], color=INK2, lw=0.9, ls=(0, (4, 3)), zorder=2)
    if text:
        ax.text(ev["t_event"] - 12, y, f"≈{ev['t_event']:.0f} s: lower levels cool,\ntop flares up",
                ha="right", va="top", fontsize=8, color=INK2)


def fig_histories(t, logs, pm, ev, colors, path):
    fig, ax = plt.subplots(figsize=(10.2, 5.6))
    ax.axvspan(0, 70, color=BAND, lw=0, zorder=0)
    for x, label in ((35, "I"), (210, "II  plateau"), (780, "III  upward spread"),
                     (ev["t_event"] + 75, "IV"), (1560, "V  burn-out")):
        ax.text(x, 1262, label, ha="center", fontsize=8.5, color=INK2)
    for thr in (600, 800):
        ax.axhline(thr, color=AXIS, lw=0.7, zorder=1)
        ax.text(END_OF_RECORD_S + 8, thr, f"{thr} °C", va="center", fontsize=8, color=MUTED)
    event_line(ax, ev, 1225)
    for i, (h, T) in enumerate(logs.items()):
        ax.plot(t, T, color=colors[h], lw=1.3, label=f"{h} mm", zorder=3)
        m = pm[h]
        ax.plot(m["t_peak"], m["T_peak"], HEIGHT_MARKERS[i], ms=6.5, color=colors[h], zorder=4, **ring())
        ax.text(150, m["plateau"] + 18, f"{h} mm", fontsize=8, color=INK, zorder=5,
                path_effects=[pe.withStroke(linewidth=3, foreground=SURFACE)])
    ax.set(xlim=(0, END_OF_RECORD_S), ylim=(0, 1300), xlabel="Time from ignition (s)", ylabel="Temperature (°C)")
    ax.set_title("Temperature histories, 500–3000 mm above the fire source (markers = peak)")
    ax.legend(loc="upper right", ncol=2, bbox_to_anchor=(1.0, 0.93))
    fig.tight_layout()
    fig.savefig(path, dpi=200)
    plt.close(fig)


def fig_space_time(t, H, M, iso, pm, ev, path):
    fig, ax = plt.subplots(figsize=(10.6, 5.4))
    edges = np.concatenate(([H[0] - (H[1] - H[0]) / 2], (H[:-1] + H[1:]) / 2, [H[-1] + (H[-1] - H[-2]) / 2]))
    t_edges = np.concatenate(([t[0]], (t[:-1] + t[1:]) / 2, [t[-1]]))
    mesh = ax.pcolormesh(t_edges, edges, M, cmap="inferno", vmin=0, vmax=1200, shading="flat", rasterized=True)
    cb = fig.colorbar(mesh, ax=ax, pad=0.015, fraction=0.035)
    cb.set_label("Temperature (°C)", color=INK2)
    cb.outline.set_visible(False)
    halo = [pe.withStroke(linewidth=2.6, foreground="#0b0b0b")]
    for thr, ls in ((500, ":"), (600, "--"), (800, "-")):
        ax.plot(t, smooth(iso[thr], 31), color="white", lw=1.3, ls=ls, path_effects=halo, label=f"{thr} °C isotherm")
    hs = sorted(pm)
    ax.plot([pm[h]["t_peak"] for h in hs], hs, color="white", lw=1.0, marker="D", ms=6,
            markerfacecolor=PEAK_COLOR, markeredgecolor="white", path_effects=halo, label="peak temperature")
    ax.axvline(ev["t_event"], color="white", lw=1.0, ls=(0, (4, 3)))
    ax.text(ev["t_event"] + 10, 350, f"≈{ev['t_event']:.0f} s", color="white", fontsize=8, path_effects=halo)
    ax.set(xlim=(0, END_OF_RECORD_S), ylim=(edges[0], edges[-1]), xlabel="Time from ignition (s)",
           ylabel="Height above fire source (mm)")
    ax.set_yticks(H)
    ax.grid(False)
    ax.set_title("Space–time temperature map (bands = thermocouple levels; isotherms interpolated between levels)")
    leg = ax.legend(loc="upper left", fontsize=8, labelcolor="white", facecolor="#0b0b0b", framealpha=0.55, frameon=True)
    leg.get_frame().set_edgecolor("none")
    fig.tight_layout()
    fig.savefig(path, dpi=200)
    plt.close(fig)


def fig_profiles(t, H, M_s, hot_level, changes, ev, path):
    fig, (ax, bx) = plt.subplots(1, 2, figsize=(10.6, 5.6), gridspec_kw={"width_ratios": [1, 1.35]})
    for times, ramp, tag in ((PROFILE_TIMES, BLUE_RAMP, ""), (POST_EVENT_TIMES, ORANGE_RAMP, " (after event)")):
        for tt, c in zip(times, ramp):
            i = int(np.argmin(np.abs(t - tt)))
            ax.plot(M_s[:, i], H, color=c, lw=1.5, marker="o", ms=5, label=f"{tt} s{tag}", **ring())
    ax.set(xlim=(0, 1200), ylim=(250, 3250), xlabel="Temperature (°C, 11-s mean)", ylabel="Height above fire source (mm)")
    ax.set_yticks(H)
    ax.set_title("(a) Vertical profiles")
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.13), ncol=4, fontsize=7.8)
    bx.step(t, hot_level, where="post", color=INK, lw=1.3)
    top = -np.inf
    for c in changes:
        if c["to_mm"] <= top:  # label only the first time each level becomes the hottest
            continue
        top = c["to_mm"]
        bx.plot(c["t_s"], c["to_mm"], "o", ms=5, color=PEAK_COLOR, zorder=3, **ring())
        right = c["t_s"] > ev["t_event"]
        bx.annotate(f"{c['t_s']:.0f} s", (c["t_s"], c["to_mm"]), xytext=(6 if right else -6, 5), textcoords="offset points",
                    ha="left" if right else "right", fontsize=7.8, color=INK2)
    event_line(bx, ev, 3200, text=False)
    bx.set(xlim=(0, END_OF_RECORD_S), ylim=(250, 3250), xlabel="Time from ignition (s)")
    bx.set_yticks(H)
    bx.set_title("(b) Hottest level vs time (31-s mean, ≥30 s dwell)")
    fig.tight_layout()
    fig.savefig(path, dpi=200)
    plt.close(fig)


def fig_fronts(t, iso, arrivals, fits, pm, ev, path):
    fig, (ax, bx) = plt.subplots(1, 2, figsize=(11.0, 5.4), sharey=True, gridspec_kw={"width_ratios": [1.6, 1]})
    for thr in ISOTHERMS:
        ax.plot(t, smooth(iso[thr], 31), color=ISO_COLOR[thr], lw=1.1, alpha=0.9, label=f"{thr} °C isotherm height")
    for thr in (600, 800):
        arr = arrivals[thr]
        hh = [h for h in sorted(arr) if np.isfinite(arr[h])]
        ax.plot([arr[h] for h in hh], hh, "o", ms=6, color=ISO_COLOR[thr], **ring())
    f = fits[800]
    tt = np.array([0.0, END_OF_RECORD_S])
    ax.plot(tt, f["z0_mm"] + f["v_mm_s"] * tt, color=ISO_COLOR[800], lw=0.9, ls="--")
    ax.text(30, 3230, f"dashed: 800 °C first-arrival front, linear fit\n≈ {f['v_mm_s']:.2f} mm/s (R² {f['r2_linear']:.3f})",
            fontsize=8, color=INK, va="top")
    hs = sorted(pm)
    ax.plot([pm[h]["t_peak"] for h in hs], hs, color=PEAK_COLOR, lw=1.3, marker="D", ms=6, label="peak temperature", **ring())
    ax.plot([pm[h]["burnout"] for h in hs], hs, color=BURN_COLOR, lw=1.2, ls="--", marker="s", ms=5,
            label=f"burn-out (< {BURNOUT_C} °C)", **ring())
    event_line(ax, ev, 3230, text=False)
    ax.set(xlim=(0, END_OF_RECORD_S), ylim=(0, 3300), xlabel="Time from ignition (s)", ylabel="Height above fire source (mm)")
    ax.set_title("(a) Fronts: isotherm heights, first arrivals (dots), peak, burn-out")
    ax.legend(loc="lower right", fontsize=8)
    for name, arr, color, mk, ls in [(f"{thr} °C", arrivals[thr], ISO_COLOR[thr], "o", "-") for thr in (600, 800, 900)] + [
            ("peak", {h: pm[h]["t_peak"] for h in hs}, PEAK_COLOR, "D", "-"),
            ("burn-out", {h: pm[h]["burnout"] for h in hs}, BURN_COLOR, "s", "--")]:
        seg = [s for s in segments(arr) if np.isfinite(s["v_mm_s"])]
        bx.plot([s["v_mm_s"] for s in seg], [(s["from_mm"] + s["to_mm"]) / 2 for s in seg], ls=ls, color=color,
                lw=0.9, marker=mk, ms=6, label=name, **ring())
    bx.set_xscale("log")
    bx.set(xlim=(0.7, 200), xlabel="Velocity between adjacent levels (mm/s, log)")
    bx.set_xticks([1, 2, 5, 10, 20, 50, 100])
    bx.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: f"{v:g}"))
    bx.xaxis.set_minor_formatter(matplotlib.ticker.NullFormatter())
    bx.axvspan(1.5, 3.5, color=BAND, lw=0, zorder=0)
    bx.text(2.3, 3180, "1.5–3.5 mm/s", ha="center", fontsize=8, color=INK2)
    bx.set_title("(b) Segment velocity (at segment midpoint)")
    bx.legend(loc="lower right", fontsize=8)
    fig.tight_layout()
    fig.savefig(path, dpi=200)
    plt.close(fig)


def fig_growth(t, logs, pm, gm, arrivals, fits, colors, path):
    hs = list(gm)
    fig, axes = plt.subplots(2, 3, figsize=(11.6, 7.2))
    for ax, h in zip(axes.flat, hs):
        m, g = pm[h], gm[h]
        sel = (t > m["onset_III"]) & (t <= m["t_peak"])
        x = g["x"]
        ax.plot(t[sel] - m["onset_III"], logs[h][sel] - m["plateau"], color=colors[h], lw=0.8, alpha=0.35)
        ax.plot(x, g["y"], color=colors[h], lw=1.5, label="measured (running max)")
        ax.plot(x, g["a_pow"] * x ** g["n"], color=INK, lw=1.2, label="power law  a·t'ⁿ")
        ax.plot(x, g["a_at2"] * x ** 2, color=INK, lw=1.2, ls="--", label="at'²")
        ax.plot(x, g["rate_lin"] * x, color=MUTED, lw=1.2, ls=":", label="linear")
        ax.text(0.03, 0.97, f"n = {g['n']:.2f} (RMSE {g['rmse_pow']:.0f} °C)\nat'²: RMSE {g['rmse_at2']:.0f} °C\n"
                            f"linear: RMSE {g['rmse_lin']:.0f} °C", transform=ax.transAxes, va="top", fontsize=7.8, color=INK)
        ax.set_title(f"{h} mm  (onset {m['onset_III']:.0f} s → peak {m['t_peak']:.0f} s)", fontsize=9.5)
        ax.set_xlabel("t' = time since phase-III onset (s)", fontsize=8.5)
        ax.set_ylabel("Rise above plateau (°C)", fontsize=8.5)
        ax.set_ylim(0, max(1000, g["y"].max() * 1.05))
    for ax in list(axes.flat)[len(hs):-1]:
        ax.set_visible(False)
    ax = axes.flat[-1]
    for thr in (600, 800, 900):
        arr = arrivals[thr]
        hh = [k for k in sorted(arr) if np.isfinite(arr[k])]
        ax.plot([arr[k] for k in hh], hh, "o-", color=ISO_COLOR[thr], lw=1.0, ms=5,
                label=f"{thr} °C: m = {fits[thr]['m']:.2f}", **ring())
    tt = np.array([100.0, 1500.0])
    for m_ref, ls, lab in ((0.8, "--", "m = 0.8 (≙ Q ∝ t² if L ∝ Q^0.4)"), (1.0, ":", "m = 1 (constant speed)")):
        ax.plot(tt, 1500 * (tt / 600) ** m_ref, color=MUTED, lw=1.0, ls=ls, label=lab)
    ax.set_xscale("log")
    ax.set_yscale("log")
    ax.set(xlim=(20, 2000), ylim=(400, 3500), xlabel="First-arrival time (s, log)", ylabel="Height (mm, log)")
    ax.set_yticks([500, 1000, 1500, 2000, 3000])
    ax.set_xticks([30, 100, 300, 1000])
    for axis in (ax.xaxis, ax.yaxis):
        axis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: f"{v:g}"))
        axis.set_minor_formatter(matplotlib.ticker.NullFormatter())
    ax.set_title("(f) Front height ∝ tᵐ", fontsize=9.5)
    ax.legend(loc="upper left", fontsize=7.2)
    handles, labels = axes.flat[0].get_legend_handles_labels()
    fig.legend(handles, labels, loc="lower center", ncol=4, bbox_to_anchor=(0.5, 0.0))
    fig.tight_layout(rect=(0, 0.05, 1, 1))
    fig.savefig(path, dpi=200)
    plt.close(fig)


def fig_height_trends(pm, gm, decay, colors, path):
    hs = sorted(pm)
    panels = [
        ("(a) Plateau rise ΔT (100–340 s)", "°C", {h: pm[h]["dT_plateau"] for h in hs}, "{:.0f}"),
        ("(b) Max heating rate, first 80 s", "°C/s", {h: pm[h]["rate_initial"] for h in hs}, "{:.1f}"),
        ("(c) Growth exponent n (phase III)", "n", {h: gm[h]["n"] for h in gm}, "{:.2f}"),
        ("(d) Peak temperature", "°C", {h: pm[h]["T_peak"] for h in hs}, "{:.0f}"),
        ("(e) Time of peak", "s", {h: pm[h]["t_peak"] for h in hs}, "{:.0f}"),
        ("(f) Time above threshold", "s", None, None),
        ("(g) Thermal dose ∫(T−T₀)dt", "°C·min", {h: pm[h]["dose_Cmin"] / 1000 for h in hs}, "{:.1f}k"),
        ("(h) First ΔT > 600 K for ≥ 30 s", "s", {h: pm[h]["t_dT600_30s"] for h in hs}, "{:.0f}"),
    ]
    fig, axes = plt.subplots(2, 4, figsize=(14.2, 6.8))
    for ax, (title, unit, vals, fmt) in zip(axes.flat, panels):
        if vals is None:
            for thr, c in ((600, BLUE_RAMP[2]), (800, BLUE_RAMP[4])):
                v = [pm[h][f"dur_{thr}"] for h in hs]
                ax.plot(hs, v, color=c, lw=1.2, marker="o", ms=5.5, **ring())
                ax.annotate(f"≥ {thr} °C", (hs[-1], v[-1]), xytext=(4, 8), textcoords="offset points",
                            fontsize=8, color=INK2)
        else:
            k = list(vals)
            ax.plot(k, [vals[h] for h in k], color=AXIS, lw=1.0, zorder=2)
            for h in k:
                ax.plot(h, vals[h], HEIGHT_MARKERS[hs.index(h)], ms=7, color=colors[h], zorder=3, **ring())
                ax.annotate(fmt.format(vals[h]), (h, vals[h]), xytext=(0, 7), textcoords="offset points",
                            ha="center", fontsize=7.6, color=INK)
        ax.set_title(title, fontsize=9.2)
        ax.set_ylabel(unit, fontsize=8.5)
        ax.set_xticks(hs)
        ax.tick_params(axis="x", labelsize=7.5)
        ax.set_xlim(hs[0] - 250, hs[-1] + 350)
        ax.margins(y=0.2)
    ax = axes.flat[0]
    zz = np.linspace(decay["z_from"] - 100, hs[-1] + 100, 50)
    ax.plot(zz, np.exp(decay["ln_a"] + decay["slope"] * zz), color=MUTED, lw=1.0, ls="--", zorder=1)
    ax.text(0.97, 0.92, f"z ≥ {decay['z_from']:.0f} mm: λ ≈ {decay['lambda_mm']:.0f} mm", transform=ax.transAxes,
            ha="right", fontsize=7.8, color=INK2)
    ax = axes.flat[2]
    for ref, lab in ((1, "linear"), (2, "t²")):
        ax.axhline(ref, color=AXIS, lw=0.8, zorder=1)
        ax.text(hs[0] - 230, ref, lab, va="bottom", fontsize=7.6, color=MUTED)
    ax.set_ylim(0, 2.3)
    ax = axes.flat[7]
    ax.axhline(900, color=AXIS, lw=0.8)
    ax.text(hs[0] - 230, 900, "15 min", va="bottom", fontsize=7.6, color=MUTED)
    for ax in axes[1]:
        ax.set_xlabel("Height above fire source (mm)", fontsize=8.5)
    fig.tight_layout()
    fig.savefig(path, dpi=200)
    plt.close(fig)


def fig_excel_check(chart, models, uni, colors, path):
    fig, ax = plt.subplots(figsize=(9.6, 5.2))
    for h, (ts, Ts) in chart.items():
        c = colors[h]
        tt = np.linspace(ts.min(), ts.max(), 200)
        ex, pr = models[h]["excel_trendline"], models[h]["excel_printed"]
        ax.plot(tt, ex["A"] * np.exp(ex["k"] * tt), color=c, lw=1.5, zorder=2)
        ax.plot(tt, pr["A"] * np.exp(pr["k"] * tt), color=c, lw=1.2, ls=":", zorder=2)
        TT = np.linspace(400, 800, 100)
        ax.plot(unified_time(uni, h, TT), TT, color=INK2, lw=1.0, ls="--", zorder=2)
        ax.plot(ts, Ts, "o", ms=6.5, color=c, zorder=3, **ring())
        ax.annotate(f"{h} mm\nExcel: {ex['A']:.0f}·e^({ex['k']:.6f} t)\nprinted: {pr['A']:.0f}·e^({pr['k']} t)",
                    (ts[0], Ts[0]), xytext=(0, -46), textcoords="offset points", ha="left", fontsize=7.8, color=INK)
    handles = [plt.Line2D([], [], color=INK2, lw=1.5, label="Excel trendline as drawn (intercept fixed, exact k)"),
               plt.Line2D([], [], color=INK2, lw=1.2, ls=":", label="equation as printed (k rounded)"),
               plt.Line2D([], [], color=INK2, lw=1.0, ls="--", label="unified model: one curve shifted by height")]
    ax.legend(handles=handles, loc="upper left")
    ax.set(xlim=(0, 1300), ylim=(250, 900), xlabel="Arrival time (s)", ylabel="Threshold temperature (°C)")
    ax.set_title("Appendix – exponential trendlines on the 400–800 °C arrival times (2000–3000 mm)")
    fig.tight_layout()
    fig.savefig(path, dpi=200)
    plt.close(fig)


# ----------------------------------------------------------------------------- main

def clean(o):
    """Strict JSON: numpy scalars -> float/int, NaN -> null."""
    if isinstance(o, dict):
        return {k: clean(v) for k, v in o.items()}
    if isinstance(o, (list, tuple, np.ndarray)):
        return [clean(v) for v in o]
    if isinstance(o, (float, np.floating)):
        return float(o) if np.isfinite(o) else None
    if isinstance(o, np.integer):
        return int(o)
    return o


def write_csv(path, rows):
    rows = list(rows)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0]))
        w.writeheader()
        for r in rows:
            w.writerow({k: ((round(v, 6) if np.isfinite(v) else "") if isinstance(v, float) else v) for k, v in r.items()})


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--xlsx", type=Path, default=HERE / "data" / "PIR_2_max.xlsx")
    args = ap.parse_args()
    for d in ("figures", "results"):
        (HERE / d).mkdir(exist_ok=True)

    t, logs = read_logs(args.xlsx)
    hs = sorted(logs)
    H = np.array(hs, float)
    M = np.array([logs[h] for h in hs])
    M11 = np.array([smooth(logs[h]) for h in hs])
    M31 = np.array([smooth(logs[h], 31) for h in hs])
    colors = {h: HEIGHT_COLORS[i] for i, h in enumerate(hs)}

    # ---- stages and events
    pm = {h: phase_metrics(t, logs[h]) for h in hs}
    ev = late_cooling_event(t, logs)

    # ---- vertical structure
    zp = np.array([pm[h]["dT_plateau"] for h in hs])
    z_from = 1000.0 if np.sum(H >= 1000) >= 3 else H[0]
    sel = H >= z_from
    r = stats.linregress(H[sel], np.log(zp[sel]))
    decay = {"z_from": z_from, "lambda_mm": float(-1 / r.slope), "r2": float(r.rvalue ** 2), "slope": float(r.slope),
             "ln_a": float(r.intercept), "loss_per_500mm": float(1 - np.exp(500 * r.slope)),
             "power_exponent": float(stats.linregress(np.log(H[sel]), np.log(zp[sel])).slope)}
    iso = {thr: isotherm_height(H, M, thr) for thr in ISOTHERMS}
    in_plateau = (t >= PLATEAU_S[0]) & (t <= PLATEAU_S[1])
    plateau_iso = {thr: float(np.nanmean(iso[thr][in_plateau])) for thr in ISOTHERMS}
    iso_at = {int(tt): {thr: float(smooth(iso[thr], 31)[int(np.argmin(np.abs(t - tt)))]) for thr in ISOTHERMS}
              for tt in (60, 300, 450, 600, 750, 900, 1050, 1200)}
    hot_level, hot_changes = hottest_level_changes(t, H, M31)
    profiles = {int(tt): {h: float(M11[k, int(np.argmin(np.abs(t - tt)))]) for k, h in enumerate(hs)}
                for tt in PROFILE_TIMES + POST_EVENT_TIMES}

    # ---- fronts
    arrivals = {thr: {h: first_arrival(t, logs[h], thr) for h in hs} for thr in THRESHOLDS}
    sustained = {thr: {h: sustained_arrival(t, logs[h], thr) for h in hs} for thr in THRESHOLDS}
    fronts = {f"{thr} C": arrivals[thr] for thr in THRESHOLDS}
    fronts.update({f"{thr} C sustained {SUSTAIN_S}s": sustained[thr] for thr in (500, 600, 700, 800)})
    fronts["peak"] = {h: pm[h]["t_peak"] for h in hs}
    hot_first = {hs[0]: 0.0}
    for c in hot_changes:
        if c["to_mm"] > max(hot_first):
            hot_first[c["to_mm"]] = c["t_s"]
    fronts["hottest level (first reach)"] = hot_first
    fronts[f"burn-out <{BURNOUT_C} C"] = {h: pm[h]["burnout"] for h in hs}
    fits = {thr: front_fit(arrivals[thr]) for thr in THRESHOLDS if np.sum(np.isfinite(list(arrivals[thr].values()))) >= 4}
    span = {}
    for thr in (600, 700, 800, 900):
        a = arrivals[thr]
        span[f"{thr} C"] = {f"{lo}-{hi}": (hi - lo) / (a[hi] - a[lo]) for lo, hi in ((500, 1500), (1500, 3000), (500, 3000), (1000, 3000))
                            if lo in a and hi in a and np.isfinite(a[lo]) and np.isfinite(a[hi])}

    # ---- growth laws
    gm = {h: growth_models(t, logs[h], pm[h]) for h in hs if pm[h]["rise_III"] >= MIN_GROWTH_RISE_C}

    # ---- appendix: the earlier Excel exponential chart (2000–3000 mm, 400–800 °C)
    chart = {h: (np.array([first_arrival(t, logs[h], T) for T in EXCEL_LEVELS]), EXCEL_LEVELS) for h in EXCEL_A if h in logs}
    am = {h: arrival_models(chart[h][0], chart[h][1], pm[h]["T0"], EXCEL_A[h], EXCEL_PRINTED_K[h]) for h in chart}
    uni = unified_model(chart)
    for h in chart:
        ts, Ts = chart[h]
        tt = np.linspace(ts.min(), ts.max(), 400)
        ex, pr = am[h]["excel_trendline"], am[h]["excel_printed"]
        am[h]["unified_rmse_t"] = rms(ts - unified_time(uni, h, Ts))
        am[h]["printed_vs_drawn_max_C"] = float(np.max(np.abs(pr["A"] * np.exp(pr["k"] * tt) - ex["A"] * np.exp(ex["k"] * tt))))
    k_common_A = {h: float(np.sum(chart[h][0] * np.log(chart[h][1] / 250.0)) / np.sum(chart[h][0] ** 2)) for h in chart}

    # ---- tables
    write_csv(HERE / "results" / "phase_metrics.csv", ({"height_mm": h, **pm[h]} for h in hs))
    write_csv(HERE / "results" / "arrival_times.csv",
              ({"threshold_C": thr, "height_mm": h, "first_arrival_s": arrivals[thr][h],
                f"sustained_{SUSTAIN_S}s_arrival_s": sustained[thr][h]} for thr in THRESHOLDS for h in hs))
    write_csv(HERE / "results" / "spread_velocity.csv",
              ({"front": name, **s} for name, arr in fronts.items() for s in segments(arr)))
    write_csv(HERE / "results" / "front_fits.csv", ({"threshold_C": thr, **f} for thr, f in fits.items()))
    growth_rows = []
    for h, g in gm.items():
        for name, par, err in (("power law", f"a={g['a_pow']:.6g}; n={g['n']:.4f}", g["rmse_pow"]),
                               ("at2", f"a={g['a_at2']:.6g}", g["rmse_at2"]),
                               ("linear", f"rate={g['rate_lin']:.6g}", g["rmse_lin"]),
                               ("power law, free origin", f"ts={g['free_origin_ts']:.0f}; n={g['free_origin_n']:.4f}", g["free_origin_rmse"])):
            growth_rows.append({"height_mm": h, "model": name, "params": par, "rmse_C": err})
    write_csv(HERE / "results" / "growth_fits.csv", growth_rows)
    write_csv(HERE / "results" / "vertical_profiles.csv",
              ({"time_s": tt, **{f"T_{h}mm": v for h, v in prof.items()}} for tt, prof in profiles.items()))
    step = (t % 10 == 0)
    write_csv(HERE / "results" / "isotherm_heights.csv",
              ({"time_s": float(tt), **{f"z_{thr}C_mm": float(iso[thr][i]) for thr in ISOTHERMS}, "hottest_level_mm": float(hot_level[i])}
               for i, tt in zip(np.flatnonzero(step), t[step])))

    summary = {
        "levels_mm": hs, "record_s": [float(t[0]), float(t[-1])],
        "phase_metrics": pm, "late_cooling_event": ev,
        "plateau_decay": decay, "plateau_isotherm_height_mm": plateau_iso, "isotherm_height_mm_at": iso_at,
        "hottest_level_changes": hot_changes, "hottest_level_first_reach_s": hot_first, "profiles_C": profiles,
        "arrival_first_s": arrivals, "arrival_sustained_s": sustained,
        "front_fits": fits, "span_velocity_mm_s": span,
        "phase_III_growth": {h: {k: v for k, v in g.items() if k not in ("x", "y")} for h, g in gm.items()},
        "excel_appendix": {"arrival_points_s": {h: chart[h][0] for h in chart},
                           "models": {h: {k: ({kk: vv for kk, vv in v.items() if kk != "t_hat"} if isinstance(v, dict) else v)
                                          for k, v in am[h].items()} for h in am},
                           "k_with_common_A_250": k_common_A, "unified_model": uni},
    }
    (HERE / "results" / "summary.json").write_text(
        json.dumps(clean(summary), indent=2, ensure_ascii=False, allow_nan=False), encoding="utf-8")

    # ---- figures
    style()
    fig_histories(t, logs, pm, ev, colors, HERE / "figures" / "fig1_temperature_histories.png")
    fig_space_time(t, H, M, iso, pm, ev, HERE / "figures" / "fig2_space_time_map.png")
    fig_profiles(t, H, M11, hot_level, hot_changes, ev, HERE / "figures" / "fig3_vertical_profiles.png")
    fig_fronts(t, iso, arrivals, fits, pm, ev, HERE / "figures" / "fig4_fire_front_velocity.png")
    fig_growth(t, logs, pm, gm, arrivals, fits, colors, HERE / "figures" / "fig5_growth_at2.png")
    fig_height_trends(pm, gm, decay, colors, HERE / "figures" / "fig6_height_trends.png")
    fig_excel_check(chart, am, uni, colors, HERE / "figures" / "figA1_excel_trendline_check.png")

    # ---- console summary
    print(f"levels {hs}, record {t[0]:.0f}-{t[-1]:.0f} s")
    for h in hs:
        m = pm[h]
        print(f"  {h:4d} mm: T0 {m['T0']:.1f}, t90 {m['t90_initial']:.0f} s, rate {m['rate_initial']:.1f} °C/s, plateau {m['plateau']:.0f}±{m['plateau_sd']:.0f}, "
              f"onset {m['onset_III']:.0f}, peak {m['T_peak']:.0f}@{m['t_peak']:.0f}, rise {m['rise_III']:.0f}, "
              f"dur300/500/600/800 {m['dur_300']:.0f}/{m['dur_500']:.0f}/{m['dur_600']:.0f}/{m['dur_800']:.0f}, "
              f"dose {m['dose_Cmin']:.0f} °C·min, burn-out {m['burnout']:.0f}, ΔT600K30s {m['t_dT600_30s']:.0f}")
    print("late cooling event:", ev)
    print("plateau decay:", decay)
    print("plateau isotherm heights:", plateau_iso, "| isotherm heights at:", iso_at)
    print("hottest level changes:", hot_changes)
    print("profiles:", {tt: {h: round(v) for h, v in p.items()} for tt, p in profiles.items()})
    for thr in THRESHOLDS:
        print(f"  {thr:5d} °C first: " + " ".join(f"{h}:{arrivals[thr][h]:.0f}" for h in hs)
              + " | sustained: " + " ".join(f"{h}:{sustained[thr][h]:.0f}" for h in hs))
    for name, arr in fronts.items():
        print(f"  {name:24s} " + "  ".join(f"{s['from_mm']}-{s['to_mm']}:{s['v_mm_s']:.2f}" for s in segments(arr)))
    print("front fits:", {thr: {k: round(v, 3) for k, v in f.items()} for thr, f in fits.items()})
    print("span velocities:", {k: {kk: round(vv, 2) for kk, vv in v.items()} for k, v in span.items()})
    for h, g in gm.items():
        print(f"  growth {h} mm: n={g['n']:.2f} (RMSE {g['rmse_pow']:.1f}), at² a={g['a_at2']:.3e} (RMSE {g['rmse_at2']:.1f}), "
              f"linear {g['rate_lin']:.3f} (RMSE {g['rmse_lin']:.1f}); free origin ts={g['free_origin_ts']:.0f} → n={g['free_origin_n']:.2f}")
    print("appendix arrival points:", {h: chart[h][0].tolist() for h in chart})
    for h in am:
        print(f"  {h}: " + " ".join(f"{k}={v['rmse_t']:.1f}" for k, v in am[h].items() if isinstance(v, dict))
              + f" | unified {am[h]['unified_rmse_t']:.1f} | k={am[h]['excel_trendline']['k']:.6f} | gap {am[h]['printed_vs_drawn_max_C']:.1f} °C"
              + f" | at² t0={am[h]['at2']['t0']:.0f}")
    print("k with A=250:", k_common_A)
    print("unified:", uni)


if __name__ == "__main__":
    main()
