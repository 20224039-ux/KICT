#!/usr/bin/env python3
"""Cavity fire tests (100 mm cavity, 30 kW gas burner for 20 min) coupled with TGA-DSC.

1. Ignition vs spread: t600 = t_ig (material starts adding heat) + t_spread (front climbs to 3000 mm).
2. Inverse surface heating: measured cavity temperature at 1000 mm as exposure, TGA decomposition temperature
   as the critical surface temperature, t_ig as the time it is reached -> effective thermal inertia (k rho c)_eff,
   TRP_eff and an apparent conductivity (illustration only: it lumps facings, melting, charring and the
   heat-transfer assumptions, so it is not the insulation's conductivity).
3. Burner off (1200 s): how much each material keeps burning on its own.

    python analysis.py          # needs ../cavity_fire_index/data/*.xlsx and ../tga_dsc/data/{N,O}/*.csv
"""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402
from scipy import stats  # noqa: E402

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent


def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


cfa = load_module("cavity_analysis", ROOT / "cavity_fire_index" / "analysis.py")
tga = load_module("tga_analysis", ROOT / "tga_dsc" / "analysis.py")

BURNER_KW, BURNER_OFF_S = 30.0, 1200          # test condition: 30 kW gas burner, 20 min
SIGMA = 5.670e-8
EXPOSURE_LEVEL = 1                              # index of 1000 mm: heating of the insulation near the burner flame
PLATEAU_S = (90, 180)                           # burner-only reference window
DEPART_K = (30, 50, 80)                         # rise above the burner-only level that marks material involvement
DEPART_MAIN = 50
HOLD_S = 30
TCRIT_RATE = (0.1, 0.2, 0.3)                    # %/°C of total loss: TGA (air) rate defining the critical temperature
TCRIT_MAIN = 0.2
HEAT_CASES = {"low": (10.0, 0.5), "central": (20.0, 0.7), "high": (30.0, 0.9)}   # h_c [W/m²K], emissivity
RHO_REF, C_REF = 30.0, 1400.0                   # kg/m³, J/kgK: only for the illustrative apparent conductivity
NOMINAL_KPC = (0.018 * 15 * 1300, 0.040 * 60 * 1500)   # catalogue k·ρ·c range of foam boards (k 0.018-0.040 W/mK,
NOMINAL_KPC_MID = float(np.sqrt(NOMINAL_KPC[0] * NOMINAL_KPC[1]))   # ρ 15-60 kg/m³, c 1300-1500 J/kgK)

# cavity material key -> (TGA file stem, Korean label, English label)
MATS = {
    "EPS_SC": ("eps_sc", "EPS 준불연", "EPS semi-noncomb."),
    "PIR_SC": ("pir_sc", "PIR(PUR) 준불연", "PIR(PUR) semi-noncomb."),
    "PF_out": ("pf_outter", "PF 외장용", "PF exterior"),
    "PF_in": ("pf_inner", "PF 내장용", "PF interior"),
}
SURFACE, INK, INK2, MUTED = "#fcfcfb", "#0b0b0b", "#52514e", "#898781"
GRID, AXIS = "#e1e0d9", "#c3c2b7"
MAT_COLOR = {"EPS_SC": "#e87ba4", "PIR_SC": "#008300", "PF_out": "#2a78d6", "PF_in": "#4a3aa7"}
MAT_MARK = {"EPS_SC": "s", "PIR_SC": "^", "PF_out": "D", "PF_in": "v"}
PHASE_COLOR = {"ignition": "#c3c2b7", "spread": "#2a78d6"}


# ----------------------------------------------------------------------------- TGA critical temperature

def tga_critical(data_dir):
    """First temperature above 150 °C where the air-atmosphere DTG reaches the given rate (mean of repeats)."""
    out = {}
    for key, (stem, _, _) in MATS.items():
        vals = {x: [] for x in TCRIT_RATE}
        for f in sorted((data_dir / "O").glob("*.csv")):
            if tga.stem(f.name)[0] != stem:
                continue
            run = tga.process(f, "O")
            T, r = run["T"], run["r"] * 100
            for x in TCRIT_RATE:
                idx = np.flatnonzero((T > 150) & (r >= x))
                vals[x].append(float(T[idx[0]]) if idx.size else np.nan)
        out[key] = {x: float(np.mean(v)) for x, v in vals.items()}
    return out


# ----------------------------------------------------------------------------- cavity events

def departure(v, thr):
    """Earliest time (s) any level from 1000 mm up stays > thr K above its burner-only level for HOLD_S."""
    t = v["t"]
    best = np.nan
    for i in range(1, 6):
        s = cfa.smooth(v["dT"][:, i], 31)
        plateau = np.nanmedian(s[PLATEAU_S[0]:PLATEAU_S[1] + 1])
        ok = np.nan_to_num(s - plateau, nan=-1e9) > thr
        run = 0
        for j in range(len(ok)):
            run = run + 1 if (ok[j] and t[j] > PLATEAU_S[1] - 60) else 0
            if run >= HOLD_S:
                tj = float(t[j - HOLD_S + 1])
                best = tj if np.isnan(best) else min(best, tj)
                break
    return best


def surface_temperature(Tg, kpc, hc, eps, n):
    """Semi-infinite solid, surface flux h_c(Tg-Ts) + eps*sigma(Tg^4-Ts^4), Duhamel superposition, 1-s steps (K).

    The flux of the current step is taken at the new surface temperature (implicit, solved by Newton), which
    keeps the scheme stable down to very small thermal inertia.
    """
    Ts = np.empty(n + 1)
    Ts[0] = Tg[0]
    q = np.empty(n)
    coef = 2.0 / np.sqrt(np.pi * kpc)
    w = np.sqrt(np.arange(1, n + 1)) - np.sqrt(np.arange(0, n))     # w[j-1] = sqrt(j) - sqrt(j-1)
    for m in range(1, n + 1):
        a = Ts[0] + coef * np.dot(q[:m - 1], w[m - 1:0:-1]) if m > 1 else Ts[0]
        g, T = Tg[m - 1], Ts[m - 1]
        for _ in range(4):
            f = hc * (g - T) + eps * SIGMA * (g ** 4 - T ** 4)
            df = -hc - 4 * eps * SIGMA * T ** 3
            T -= (T - a - coef * f) / (1 - coef * df)
        q[m - 1] = hc * (g - T) + eps * SIGMA * (g ** 4 - T ** 4)
        Ts[m] = T
    return Ts


KPC_LOG10 = (1.0, 9.0)


def invert_kpc(Tg, t_ig, T_crit, hc, eps):
    """(k rho c)_eff [W² s / m⁴ K²] for which the surface reaches T_crit exactly at t_ig.

    NaN when the cavity gas itself never gets hot enough (even a near-zero inertia surface stays below T_crit).
    """
    if not np.isfinite(t_ig):
        return np.nan
    n = int(round(t_ig))
    if n < 2 or T_crit <= Tg[0] or surface_temperature(Tg, 10 ** KPC_LOG10[0], hc, eps, n)[-1] < T_crit:
        return np.nan
    lo, hi = KPC_LOG10
    for _ in range(30):
        mid = 0.5 * (lo + hi)
        if surface_temperature(Tg, 10 ** mid, hc, eps, n)[-1] >= T_crit:
            lo = mid                       # surface too hot -> more inertia needed
        else:
            hi = mid
    return 10 ** (0.5 * (lo + hi))


# ----------------------------------------------------------------------------- figures

def style():
    tga.style()


def fig_phases(ph, path):
    fig, ax = plt.subplots(figsize=(11, 4.8))
    rows = ph.sort_values("t600_s")
    y = np.arange(len(rows))[::-1]
    for yy, (_, r) in zip(y, rows.iterrows()):
        ax.barh(yy, r["t_ig_s"], color=PHASE_COLOR["ignition"], height=0.6)
        if np.isfinite(r["t600_s"]):
            ax.barh(yy, r["t600_s"] - r["t_ig_s"], left=r["t_ig_s"], color=PHASE_COLOR["spread"], height=0.6)
            ax.text(r["t600_s"] + 15, yy, f"{r['t_ig_s']:.0f} + {r['t600_s'] - r['t_ig_s']:.0f} = {r['t600_s']:.0f} s",
                    va="center", fontsize=8.5, color=INK, bbox=dict(fc=SURFACE, ec="none", pad=1.5))
        else:
            ax.text(r["t_ig_s"] + 15, yy, "never reached dT 600 K at 3000 mm", va="center", fontsize=8.5, color=INK2,
                    bbox=dict(fc=SURFACE, ec="none", pad=1.5))
    ax.axvline(BURNER_OFF_S, color=INK2, lw=0.9, ls=(0, (4, 3)))
    ax.text(BURNER_OFF_S + 10, len(rows) - 0.4, "burner off (20 min)", fontsize=8.5, color=INK2)
    ax.set_yticks(y)
    ax.set_yticklabels([f"{MATS[r['material']][2]} #{int(r['repeat'])}" for _, r in rows.iterrows()])
    ax.set_xlim(0, 1900)
    ax.set_xlabel("Time from burner ignition (s)")
    ax.text(10, -0.72, "EPS semi-noncomb.: grey = melt-back period, blue starts at the delayed flash", ha="left",
            fontsize=8, color=INK2)
    ax.set_ylim(-1.0, len(rows) - 0.3)
    ax.grid(axis="y", visible=False)
    ax.set_title("t600 split into ignition (grey: first +50 K above the burner-only level) and spread (blue)")
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)


def fig_tga_links(mat, path):
    fig, axs = plt.subplots(1, 2, figsize=(14, 5.4))
    panels = (
        (axs[0], "tga_rmax_air", "t_spread_med", "t_spread_min", "t_spread_max",
         "TGA, air: peak mass-loss rate (% of total loss per °C)", "Cavity spread phase t600 − t_ig (s)",
         "(a) Spread phase vs oxidative reactivity"),
        (axs[1], "tga_f400_air", "D600_after_off_med", "D600_after_off_min", "D600_after_off_max",
         "TGA, air: share of mass loss above 400 °C (char oxidation)",
         "After burner off: Σ over heights of s with dT ≥ 600 K (1200–1800 s)",
         "(b) Burning without the burner vs char-oxidation share"),
    )
    for ax, xc, yc, ylo, yhi, xl, yl, title in panels:
        for _, r in mat.iterrows():
            k = r["material"]
            if not np.isfinite(r[yc]) or (ax is axs[0] and k == "EPS_SC"):
                continue
            ax.errorbar([r[xc]], [r[yc]], yerr=[[r[yc] - r[ylo]], [r[yhi] - r[yc]]], fmt=MAT_MARK[k],
                        color=MAT_COLOR[k], ms=10, mec=SURFACE, mew=1.2, elinewidth=1, capsize=0)
            left = ax is axs[1] and k == "PIR_SC"
            ax.annotate(MATS[k][2], (r[xc], r[yc]), xytext=(-9 if left else 9, 6), textcoords="offset points",
                        ha="right" if left else "left", fontsize=9, color=INK)
        ax.set_xlabel(xl)
        ax.set_ylabel(yl)
        ax.set_title(title)
    axs[0].set_xlim(0.3, 1.3)
    axs[0].text(0.98, 0.97, "EPS semi-noncomb. not shown:\nmelts back, then flashes", transform=axs[0].transAxes,
                ha="right", va="top", fontsize=8, color=INK2)
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)


def fig_inertia(mat, path):
    m = mat[mat["kpc_central"].notna()].reset_index(drop=True)
    fig, ax = plt.subplots(figsize=(11, 4.6))
    y = np.arange(len(m))[::-1]
    for yy, (_, r) in zip(y, m.iterrows()):
        k = r["material"]
        ax.plot([r["kpc_min"], r["kpc_max"]], [yy, yy], color=MAT_COLOR[k], lw=2.4, solid_capstyle="round")
        ax.plot([r["kpc_central"]], [yy], marker=MAT_MARK[k], color=MAT_COLOR[k], ms=11, mec=SURFACE, mew=1.2)
        ax.annotate(f"×{r['kpc_central'] / NOMINAL_KPC_MID:,.0f}", (r["kpc_central"], yy), xytext=(0, 10),
                    textcoords="offset points", ha="center", fontsize=8.5, color=INK)
    ax.axvspan(*NOMINAL_KPC, color=GRID, lw=0)
    ax.text(np.sqrt(NOMINAL_KPC[0] * NOMINAL_KPC[1]), y[-1] - 0.45, "bare foam\n(catalogue k, ρ, c)", ha="center",
            va="center", fontsize=8.5, color=INK2)
    ax.set_xscale("log")
    ax.set_xlim(1e2, 1e8)
    ax.set_ylim(y[-1] - 0.8, y[0] + 0.6)
    ax.set_yticks(y)
    ax.set_yticklabels([MATS[k][2] for k in m["material"]])
    ax.grid(axis="y", visible=False)
    ax.set_xlabel("Effective thermal inertia (k·ρ·c)_eff  (W² s m⁻⁴ K⁻²)   — dot: central case, bar: all 27 assumption sets")
    top = ax.secondary_xaxis("top", functions=(lambda x: x / (RHO_REF * C_REF), lambda k: k * RHO_REF * C_REF))
    top.set_xlabel(f"…read as conductivity with ρ = {RHO_REF:.0f} kg/m³, c = {C_REF:.0f} J/kgK  (W/mK)", color=INK2)
    ax.set_title("Inverse estimate from the cavity ignition time (× = central / mid catalogue value)", pad=34)
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)


def fig_burner_off(bo, path):
    tests = list(dict.fromkeys(bo["test"]))
    M = np.array([[bo[(bo.test == k) & (bo.height_mm == z)]["ssr"].iloc[0] for z in cfa.HEIGHTS] for k in tests])
    fig, ax = plt.subplots(figsize=(10, 5.6))
    norm = matplotlib.colors.TwoSlopeNorm(vmin=0.0, vcenter=1.0, vmax=2.0)
    im = ax.imshow(M, cmap="RdBu_r", norm=norm, aspect="auto")
    for i in range(M.shape[0]):
        for j in range(M.shape[1]):
            ax.text(j, i, f"{M[i, j]:.2f}", ha="center", va="center", fontsize=8.5,
                    color=INK if 0.4 < M[i, j] < 1.55 else SURFACE)
    ax.set_xticks(range(6))
    ax.set_xticklabels([f"{z} mm" for z in cfa.HEIGHTS])
    ax.set_yticks(range(len(tests)))
    ax.set_yticklabels([f"{MATS[bo[bo.test == k].material.iloc[0]][2]} #{k.split('-')[-1]}" for k in tests])
    ax.grid(False)
    cb = fig.colorbar(im, ax=ax, fraction=0.04)
    cb.set_label("mean dT 1200–1500 s / mean dT 900–1200 s  (1 = unchanged)")
    ax.set_title("After the burner stops: ratio of temperature rise (5 min after / 5 min before)")
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)


# ----------------------------------------------------------------------------- main

def main():
    D = ROOT / "cavity_fire_index" / "data"
    tests = cfa.read_tests([D / "EPS.xlsx", D / "PF.xlsx", D / "PIR_SC_2_max.xlsx"])
    cfa.prepare(tests)
    idx = {k: cfa.index_3000(k, v) for k, v in tests.items()}
    keys = [k for k in sorted(tests) if tests[k]["full"]]
    tcrit = tga_critical(ROOT / "tga_dsc" / "data")
    tsum = pd.read_csv(ROOT / "tga_dsc" / "results" / "tga_summary.csv")

    rows, inv_rows = [], []
    for k in keys:
        v = tests[k]
        mat = v["material"]
        dep = {thr: departure(v, thr) for thr in DEPART_K}
        t_ig = dep[DEPART_MAIN]
        t600 = idx[k]["t600_s"]
        Tg = cfa.smooth(v["H"][:, EXPOSURE_LEVEL], 11) + 273.15
        kpcs = {}
        for thr in DEPART_K:
            for rate in TCRIT_RATE:
                for case, (hc, eps) in HEAT_CASES.items():
                    kp = invert_kpc(Tg, dep[thr], tcrit[mat][rate] + 273.15, hc, eps) if mat != "EPS_SC" else np.nan
                    kpcs[(thr, rate, case)] = kp
                    inv_rows.append(dict(test=k, material=mat, depart_K=thr, tcrit_rate=rate, heat_case=case,
                                         t_ig_s=dep[thr], T_crit_C=tcrit[mat][rate], kpc_eff=kp))
        central = kpcs[(DEPART_MAIN, TCRIT_MAIN, "central")]
        allv = np.array([x for x in kpcs.values() if np.isfinite(x)])
        dTc = tcrit[mat][TCRIT_MAIN] - (Tg[0] - 273.15)
        rows.append(dict(test=k, material=mat, repeat=v["repeat"], t_ig_s=t_ig, t_ig_range_s=f"{min(dep.values()):.0f}-{max(dep.values()):.0f}",
                         t600_s=t600, t_spread_s=t600 - t_ig if np.isfinite(t600) else np.nan,
                         T_exposure_1000mm_mean_C=float(np.nanmean(Tg[:int(t_ig) + 1]) - 273.15) if np.isfinite(t_ig) else np.nan,
                         T_crit_C=tcrit[mat][TCRIT_MAIN], kpc_central=central,
                         kpc_min=float(allv.min()) if allv.size else np.nan, kpc_max=float(allv.max()) if allv.size else np.nan,
                         TRP_eff_kW_s05_m2=dTc * np.sqrt(np.pi * central / 4) / 1000 if np.isfinite(central) else np.nan,
                         k_app_W_mK_at_ref=central / (RHO_REF * C_REF) if np.isfinite(central) else np.nan))
    ph = pd.DataFrame(rows)
    ph.round(4).to_csv(HERE / "results" / "ignition_spread_inverse.csv", index=False, encoding="utf-8")
    pd.DataFrame(inv_rows).round(4).to_csv(HERE / "results" / "inverse_sensitivity.csv", index=False, encoding="utf-8")

    # burner-off behaviour
    bo_rows = []
    for k in keys + ["EPS_C-1", "EPS_C-2"]:
        v = tests[k]
        if not v["full"]:
            continue
        dT = v["dT"]
        for i, z in enumerate(cfa.HEIGHTS):
            on = float(np.nanmean(dT[BURNER_OFF_S - 300:BURNER_OFF_S, i]))
            off1 = float(np.nanmean(dT[BURNER_OFF_S:BURNER_OFF_S + 300, i]))
            off2 = float(np.nanmean(dT[BURNER_OFF_S + 300:1800, i]))
            bo_rows.append(dict(test=k, material=v["material"], height_mm=z, dT_on_900_1200=on, dT_off_1200_1500=off1,
                                dT_off_1500_1800=off2, ssr=off1 / on if on > 0 else np.nan,
                                D600_after_off_s=int(np.sum(np.nan_to_num(dT[BURNER_OFF_S:1801, i]) >= 600)),
                                D600_before_off_s=int(np.sum(np.nan_to_num(dT[:BURNER_OFF_S, i]) >= 600))))
    bo = pd.DataFrame(bo_rows)
    bo.round(3).to_csv(HERE / "results" / "burner_off.csv", index=False, encoding="utf-8")

    # material level + TGA
    mrows = []
    for mat, (stem, ko, en) in MATS.items():
        p = ph[ph.material == mat]
        b = bo[bo.material == mat]
        sp = p["t_spread_s"].to_numpy(dtype=float)
        sp_c = np.where(np.isfinite(sp), sp, 1800 - p["t_ig_s"].to_numpy(dtype=float))
        s_air = tsum[(tsum.atmosphere == "oxidative (O)") & (tsum.material == mat)].iloc[0]
        d6 = b.groupby("test")["D600_after_off_s"].sum()
        mrows.append(dict(
            material=mat, label_ko=ko, n=len(p), t_ig_med=float(np.median(p["t_ig_s"])), t_spread_med=float(np.median(sp_c)),
            t_spread_min=float(np.min(sp_c)), t_spread_max=float(np.max(sp_c)),
            kpc_central=float(np.nanmedian(p["kpc_central"])) if p["kpc_central"].notna().any() else np.nan,
            kpc_min=float(np.nanmin(p["kpc_min"])) if p["kpc_min"].notna().any() else np.nan,
            kpc_max=float(np.nanmax(p["kpc_max"])) if p["kpc_max"].notna().any() else np.nan,
            TRP_eff=float(np.nanmedian(p["TRP_eff_kW_s05_m2"])) if p["TRP_eff_kW_s05_m2"].notna().any() else np.nan,
            ssr_low_500_1500=float(b[b.height_mm <= 1500]["ssr"].mean()), ssr_top_2500_3000=float(b[b.height_mm >= 2500]["ssr"].mean()),
            D600_after_off_med=float(d6.median()), D600_after_off_min=float(d6.min()), D600_after_off_max=float(d6.max()),
            kpc_ratio_to_bare_foam=float(np.nanmedian(p["kpc_central"])) / NOMINAL_KPC_MID if p["kpc_central"].notna().any() else np.nan,
            tga_rmax_air=float(s_air["r_max_pct_C_mean"]), tga_f400_air=float(s_air["f_above400_mean"]),
            tga_Ta90_air=float(s_air["T_a90_C_mean"]), T_crit_C=tcrit[mat][TCRIT_MAIN]))
    mat = pd.DataFrame(mrows)
    mat.round(4).to_csv(HERE / "results" / "material_summary.csv", index=False, encoding="utf-8")
    sub = mat[mat.material != "EPS_SC"]
    rho_sp = stats.spearmanr(sub["tga_rmax_air"], -sub["t_spread_med"]).statistic
    rho_char = stats.spearmanr(mat["tga_f400_air"], mat["D600_after_off_med"])

    style()
    fig_phases(ph, HERE / "figures" / "fig1_ignition_spread.png")
    fig_tga_links(mat, HERE / "figures" / "fig2_tga_links.png")
    fig_inertia(mat, HERE / "figures" / "fig3_inverse_inertia.png")
    fig_burner_off(bo, HERE / "figures" / "fig4_burner_off.png")

    pd.set_option("display.width", 230)
    pd.set_option("display.max_columns", 40)
    print(ph.round(3).to_string())
    print(mat.round(3).to_string())
    print("spearman rmax_air vs spread speed (PF_in, PF_out, PIR):", rho_sp)
    print("spearman f400_air vs D600 after burner off (4 materials):", rho_char.statistic, rho_char.pvalue)
    print(bo.pivot_table(index="test", columns="height_mm", values="ssr").round(2).to_string())
    print(bo.pivot_table(index="test", columns="height_mm", values="D600_after_off_s").to_string())
    print("TGA critical temperatures:", {k: {r: round(x, 1) for r, x in v.items()} for k, v in tcrit.items()})


if __name__ == "__main__":
    (HERE / "figures").mkdir(parents=True, exist_ok=True)
    (HERE / "results").mkdir(parents=True, exist_ok=True)
    sys.exit(main())
