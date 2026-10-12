#!/usr/bin/env python3
"""TGA-DSC of the cavity-test insulation materials (N2 and oxidative atmosphere) and its link to the 3 m cavity fire tests.

PerkinElmer STA exports (columns: Time [min], Unsubtracted Weight [mg], Program/Sample Temperature [°C],
Unsubtracted Heat Flow [mW, endo up], ...), 5 °C/min to 800 °C, 20 mL/min purge.

    python analysis.py                         # reads data/N/*.csv and data/O/*.csv
    python analysis.py --data /path/to/dir     # same layout

The weight signal has no absolute zero in these exports (several runs start below 0 mg) and no blank
(buoyancy / heat-flow) baseline was subtracted, so every mass metric here is normalised to the mass lost
between the start of the ramp and 790 °C, and the DSC is used qualitatively (sign and temperature of events).
Writes figures/*.png and results/*.csv next to this file.
"""
from __future__ import annotations

import argparse
import csv
import glob
import os
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402
from scipy import stats  # noqa: E402
from scipy.signal import find_peaks, savgol_filter  # noqa: E402

HERE = Path(__file__).resolve().parent
FIRE = HERE.parent / "cavity_fire_index" / "results"
GRID_STEP = 0.5            # °C
SMOOTH_RAW = 241           # raw samples (8 Hz, 5 °C/min -> about 2.5 °C)
SMOOTH_DTG = 21            # grid points (10 °C)
HT_SPLIT_C = 400.0         # mass lost above this = high-temperature (char) stage
DSC_LOCAL_WIN = 201        # grid points (100 °C): local heat-flow trend removed for the qualitative DSC view
PEAK_PROM = 0.0015         # 1/°C, DTG peak prominence
PEAK_TMAX_C = 750.0        # main-peak search window upper limit (loss still accelerating at the ramp end for PF in N2)

# file stem -> (material key used in the cavity analysis, Korean label, English label)
MATERIALS = {
    "eps_c": ("EPS_C", "EPS 가연성", "EPS combustible"),
    "eps_sc": ("EPS_SC", "EPS 준불연", "EPS semi-noncomb."),
    "pir_c": ("PIR_C", "PIR 가연성", "PIR combustible"),
    "pir_sc": ("PIR_SC", "PIR 준불연", "PIR semi-noncomb."),
    "pf_outter": ("PF_out", "PF 외장용", "PF exterior"),
    "pf_inner": ("PF_in", "PF 내장용", "PF interior"),
}
ATM = {"N": "N2", "O": "oxidative (O)"}

# chart tokens (dataviz reference palette, light surface); atmosphere = categorical slots 1-2
SURFACE, INK, INK2, MUTED = "#fcfcfb", "#0b0b0b", "#52514e", "#898781"
GRID, AXIS = "#e1e0d9", "#c3c2b7"
ATM_COLOR = {"N": "#2a78d6", "O": "#eb6834"}
MAT_COLOR = {"EPS_SC": "#e87ba4", "PIR_SC": "#008300", "PF_out": "#2a78d6", "PF_in": "#4a3aa7"}
MAT_MARK = {"EPS_SC": "s", "PIR_SC": "^", "PF_out": "D", "PF_in": "v"}


def stem(name):
    s = Path(name).stem.lower()
    mat, rep = s.rsplit("_", 1)
    return mat, int(rep)


def process(path, atm):
    d = pd.read_csv(path)
    d.columns = [c.strip() for c in d.columns]
    t, Tp = d["Time"].to_numpy(), d["Program Temperature"].to_numpy()
    Ts, w, hf = d["Sample Temperature"].to_numpy(), d["Unsubtracted Weight"].to_numpy(), d["Unsubtracted Heat Flow"].to_numpy()
    i0 = int(np.argmax(Tp > Tp[0] + 0.5))                 # ramp start (after the initial isothermal)
    i1 = int(np.argmax(Tp >= Tp.max() - 0.01))            # ramp end
    sl = slice(i0, i1 + 1)
    Tm = np.maximum.accumulate(savgol_filter(Ts[sl], SMOOTH_RAW, 2))
    T = np.arange(np.ceil(Tm[0]) + 5, np.floor(Tm[-1]) - 2, GRID_STEP)
    W = np.interp(T, Tm, savgol_filter(w[sl], SMOOTH_RAW, 2))
    H = np.interp(T, Tm, savgol_filter(hf[sl], SMOOTH_RAW, 2))
    w_ref, w_end = W[:20].mean(), W[-10:].mean()
    dm = w_ref - w_end
    a = (w_ref - W) / dm                                   # conversion: share of the mass lost up to the ramp end
    r = savgol_filter(np.gradient(a, T), SMOOTH_DTG, 2)    # DTG, 1/°C
    h_local = H - savgol_filter(H, DSC_LOCAL_WIN, 2)       # qualitative local heat flow, mW, endo up
    pk, _ = find_peaks(r, prominence=PEAK_PROM, distance=30)
    pk = pk[np.argsort(T[pk])]
    band = (T >= 100) & (T <= PEAK_TMAX_C)                  # main peak away from the ramp ends
    p = int(np.flatnonzero(band)[np.argmax(r[band])])
    valley = pk[pk < p]
    lo = int(np.argmin(r[(valley.max() if valley.size else 0):p + 1]) + (valley.max() if valley.size else 0)) if p > 0 else 0
    t_on = float(T[p] - (a[p] - a[lo]) / r[p]) if r[p] > 0 else np.nan
    at = lambda x: float(T[np.argmax(a >= x)]) if (a >= x).any() else np.nan  # noqa: E731
    win = (T > 100) & (T < 780)
    m = dict(
        t_ramp_start_min=float(t[i0]), T_ramp_end_C=float(T[-1]), dm_tot_mg=float(dm),
        T_a05_C=at(0.05), T_a10_C=at(0.10), T_a50_C=at(0.50), T_a90_C=at(0.90),
        T_onset_C=t_on, T_peak_C=float(T[p]), r_max_pct_C=float(r[p] * 100),
        f_above400=float(1 - np.interp(HT_SPLIT_C, T, a)),
        dtg_peaks=" ".join(f"{T[i]:.0f}" for i in pk if r[i] * 100 >= 0.15),
        dsc_at_main_peak_mW=float(np.mean(h_local[(T > T[p] - 5) & (T < T[p] + 5)])),
        dsc_exo_peak_C=float(T[win][np.argmin(h_local[win])]), dsc_exo_mW=float(h_local[win].min()),
        dsc_endo_peak_C=float(T[win][np.argmax(h_local[win])]), dsc_endo_mW=float(h_local[win].max()),
    )
    return dict(T=T, a=a, r=r, h=h_local, rem=100 * (1 - a), dtg=100 * r, metrics=m)


def style():
    plt.rcParams.update({
        "figure.facecolor": SURFACE, "axes.facecolor": SURFACE, "savefig.facecolor": SURFACE,
        "axes.edgecolor": AXIS, "axes.linewidth": 0.8, "axes.labelcolor": INK2,
        "axes.titlesize": 10, "axes.titleweight": "bold", "axes.titlelocation": "left", "axes.titlecolor": INK,
        "axes.labelsize": 9.5, "font.size": 9.5, "xtick.color": AXIS, "ytick.color": AXIS,
        "xtick.labelcolor": INK2, "ytick.labelcolor": INK2, "xtick.labelsize": 8.5, "ytick.labelsize": 8.5,
        "axes.grid": True, "grid.color": GRID, "grid.linewidth": 0.6, "axes.axisbelow": True,
        "axes.spines.top": False, "axes.spines.right": False, "legend.frameon": False, "legend.fontsize": 8.5,
        "legend.labelcolor": INK2, "lines.solid_capstyle": "round",
    })


def small_multiples(runs, key, ylabel, title, path, ylim=None, note=None, xlim=(50, 800)):
    fig, axs = plt.subplots(2, 3, figsize=(15, 8.4), sharex=True, sharey=ylim is not None)
    for ax, (mat, (_, _, en)) in zip(axs.ravel(), MATERIALS.items()):
        for atm in ("N", "O"):
            for rep in (1, 2):
                run = runs.get((atm, mat, rep))
                if run is None:
                    continue
                ax.plot(run["T"], run[key], color=ATM_COLOR[atm], lw=1.4 if rep == 1 else 1.0,
                        ls="-" if rep == 1 else (0, (4, 2)), label=f"{ATM[atm]} #{rep}")
        ax.set_title(en)
        if ylim:
            ax.set_ylim(*ylim)
        ax.set_xlim(*xlim)
        if xlim != (50, 800):
            ys = [v for (_, m2, _), r2 in runs.items() if m2 == mat
                  for sel in [(r2["T"] >= xlim[0]) & (r2["T"] <= xlim[1])] for v in (r2[key][sel].min(), r2[key][sel].max())]
            pad = 0.08 * (max(ys) - min(ys))
            ax.set_ylim(min(ys) - pad, max(ys) + pad)
    for ax in axs[1]:
        ax.set_xlabel("Sample temperature (°C)")
    for ax in axs[:, 0]:
        ax.set_ylabel(ylabel)
    h, lab = axs[0, 0].get_legend_handles_labels()
    fig.legend(h, lab, loc="upper center", ncol=4, bbox_to_anchor=(0.5, 0.985))
    fig.suptitle(title, x=0.01, y=0.94, ha="left", fontsize=11, fontweight="bold", color=INK)
    if note:
        fig.text(0.01, 0.005, note, fontsize=8.5, color=INK2)
    fig.tight_layout(rect=(0, 0.02, 1, 0.92))
    fig.savefig(path, dpi=150)
    plt.close(fig)


def fire_table():
    """Per-material medians of the 3 m cavity indices (not-reached t600 counted as > 1800 s)."""
    idx = {r["test"]: r for r in csv.DictReader(open(FIRE / "index_3000mm.csv", encoding="utf-8"))}
    rm = {r["test"]: r for r in csv.DictReader(open(FIRE / "risk_model.csv", encoding="utf-8"))}
    out = {}
    for k, r in rm.items():
        mat = r["material"]
        t600 = float(r["t600_s"]) if r["t600_s"] else 1800.0 + 1
        out.setdefault(mat, []).append(dict(t600=t600, D600=float(r["D600_min"]), R=int(r["R"]),
                                            H600=float(idx[k]["excess_600K_Kmin"])))
    return {m: dict(n=len(v), t600_med=float(np.median([x["t600"] for x in v])),
                    t600_min=min(x["t600"] for x in v), t600_max=max(x["t600"] for x in v),
                    D600_med=float(np.median([x["D600"] for x in v])), R_max=max(x["R"] for x in v),
                    H600_med=float(np.median([x["H600"] for x in v]))) for m, v in out.items()}


def fig_fire_link(summary, fire, path):
    fig, axs = plt.subplots(1, 2, figsize=(14, 5.6), sharey=True)
    for ax, atm in zip(axs, ("O", "N")):
        for mat, (key, _, en) in MATERIALS.items():
            if key not in fire:
                continue
            s = summary[(atm, mat)]
            f = fire[key]
            x, xe = s["r_max_pct_C_mean"], s["r_max_pct_C_halfrange"]
            ax.errorbar([x], [f["t600_med"]], xerr=[[xe], [xe]],
                        yerr=[[f["t600_med"] - f["t600_min"]], [min(f["t600_max"], 1800) - f["t600_med"]]],
                        fmt=MAT_MARK[key], color=MAT_COLOR[key], ms=10, mec=SURFACE, mew=1.2, elinewidth=1, capsize=0)
            ax.annotate(en, (x, f["t600_med"]), xytext=(10, 6), textcoords="offset points", fontsize=9, color=INK)
        ax.axhline(900, color=INK2, lw=0.8, ls=(0, (4, 3)))
        ax.text(0.02, 915, "15 min", transform=ax.get_yaxis_transform(), fontsize=8, color=INK2)
        ax.set_xlabel(f"Peak mass-loss rate in {ATM[atm]} (% of total loss per °C)")
        ax.set_title(("(a) " if atm == "O" else "(b) ") + f"TGA in {ATM[atm]} vs cavity t600")
    axs[0].set_ylabel("Cavity t600 at 3000 mm (s, median; bar = repeats)")
    fig.text(0.01, 0.005, "Upper bar of EPS semi-noncomb. reaches 1800 s: one of three cavity tests never reached the criterion.",
             fontsize=8.5, color=INK2)
    axs[0].set_ylim(300, 1850)
    fig.tight_layout(rect=(0, 0.03, 1, 1))
    fig.savefig(path, dpi=150)
    plt.close(fig)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--data", type=Path, default=HERE / "data")
    args = ap.parse_args()
    (HERE / "figures").mkdir(exist_ok=True)
    (HERE / "results").mkdir(exist_ok=True)

    runs, rows = {}, []
    for atm in ("N", "O"):
        for f in sorted(glob.glob(os.path.join(args.data, atm, "*.csv"))):
            mat, rep = stem(f)
            run = process(f, atm)
            runs[(atm, mat, rep)] = run
            rows.append(dict(atmosphere=ATM[atm], material=MATERIALS[mat][0], label_ko=MATERIALS[mat][1], repeat=rep,
                             file=os.path.basename(f), **run["metrics"]))
    df = pd.DataFrame(rows)
    df.round(4).to_csv(HERE / "results" / "tga_metrics.csv", index=False, encoding="utf-8")

    keys = ["dm_tot_mg", "T_a05_C", "T_a10_C", "T_a50_C", "T_a90_C", "T_onset_C", "T_peak_C", "r_max_pct_C",
            "f_above400", "dsc_at_main_peak_mW", "dsc_exo_peak_C", "dsc_exo_mW"]
    summary, srows = {}, []
    for atm in ("N", "O"):
        for mat, (key, ko, _) in MATERIALS.items():
            sub = df[(df.atmosphere == ATM[atm]) & (df.material == key)]
            s = {}
            for k in keys:
                v = sub[k].to_numpy(dtype=float)
                s[f"{k}_mean"] = float(np.mean(v))
                s[f"{k}_halfrange"] = float(np.ptp(v) / 2)
            summary[(atm, mat)] = s
            srows.append(dict(atmosphere=ATM[atm], material=key, label_ko=ko, n=len(sub), **s))
    pd.DataFrame(srows).round(4).to_csv(HERE / "results" / "tga_summary.csv", index=False, encoding="utf-8")

    fire = fire_table()
    link = []
    for mat, (key, ko, _) in MATERIALS.items():
        if key not in fire:
            continue
        f = fire[key]
        link.append(dict(material=key, label_ko=ko, cavity_tests=f["n"], t600_median_s=f["t600_med"],
                         D600_median_min=f["D600_med"], H600_median_Kmin=f["H600_med"], R_max=f["R_max"],
                         **{f"{a}_{k}": summary[(a, mat)][f"{k}_mean"] for a in ("O", "N")
                            for k in ("T_onset_C", "T_peak_C", "r_max_pct_C", "T_a90_C", "f_above400")}))
    ldf = pd.DataFrame(link)
    ldf.round(4).to_csv(HERE / "results" / "fire_link.csv", index=False, encoding="utf-8")
    rho = []
    speed = -ldf["t600_median_s"].to_numpy()          # faster spread = larger
    sev = ldf["D600_median_min"].to_numpy()
    for col in [c for c in ldf.columns if c.startswith(("O_", "N_"))]:
        v = ldf[col].to_numpy(dtype=float)
        rho.append(dict(tga_metric=col, spearman_vs_speed=float(stats.spearmanr(v, speed).statistic),
                        spearman_vs_D600=float(stats.spearmanr(v, sev).statistic), n=len(v)))
    pd.DataFrame(rho).round(3).to_csv(HERE / "results" / "fire_link_rank.csv", index=False, encoding="utf-8")

    style()
    small_multiples(runs, "rem", "Mass remaining, % of the loss up to 790 °C",
                    "Mass-loss curves (normalised): solid = repeat 1, dashed = repeat 2",
                    HERE / "figures" / "fig1_mass_loss.png", ylim=(-3, 103),
                    note="Normalised to the mass lost between ramp start and 790 °C (absolute sample mass not in the export).")
    small_multiples(runs, "dtg", "DTG, % of total loss per °C", "Derivative mass loss (DTG)",
                    HERE / "figures" / "fig2_dtg.png")
    small_multiples(runs, "h", "Local heat flow (mW, endo up)",
                    "DSC, qualitative: heat flow minus its 100 °C running trend (no blank run available)",
                    HERE / "figures" / "fig3_dsc_local.png", xlim=(100, 700),
                    note="Positive = endothermic (EPS depolymerisation in N2), negative = exothermic. 100-700 °C only; not for enthalpy values.")
    fig_fire_link(summary, fire, HERE / "figures" / "fig4_fire_link.png")

    pd.set_option("display.width", 220)
    pd.set_option("display.max_columns", 40)
    print(pd.DataFrame(srows)[["atmosphere", "material", "dm_tot_mg_mean", "T_a05_C_mean", "T_onset_C_mean", "T_peak_C_mean",
                               "r_max_pct_C_mean", "r_max_pct_C_halfrange", "T_a50_C_mean", "T_a90_C_mean", "f_above400_mean",
                               "dsc_at_main_peak_mW_mean", "dsc_exo_peak_C_mean"]].round(2).to_string())
    print(ldf.round(2).to_string())
    print(pd.DataFrame(rho).round(2).to_string())


if __name__ == "__main__":
    main()
