#!/usr/bin/env python3
"""Integrated cavity-fire hazard model: medium-scale cavity tests (30 kW burner, 20 min) + TGA in air.

Two regimes found in the data (see README.md):
    combustible grade (c)       PIR(PUR) c: whole cavity at dT 600 K within ~30 s, burnt out before the burner stops
    semi-noncombustible (sc)    ignition  t_ig ~ system constant (~4 min; facings, flame attachment)
                                spread    t_sp = t600 - t_ig,  t_sp * r_max,air ~ C_sys
                                          -> t600_hat = t_ig,sys + C_sys / r_max,air   (charring sc foams only)
                                afterburn burner off at 1200 s: temperature held or rising at >= 3 of 6 heights (B) ~ f400,air
                                melt-back non-charring foams (f400,air < 0.5): melt away, then flash (flag A)

The TGA screening only applies after the grade gate (c / sc). A blind check is kept: the prediction for
PIR(PUR) c made before its cavity data existed (fit on the five tests available then) against the measurement.

Reads only result tables already in the repository (no raw data needed):
    ../coupled_analysis/results/{ignition_spread_inverse,burner_off}.csv
    ../cavity_fire_index/results/risk_model.csv
    ../tga_dsc/results/tga_summary.csv

    python analysis.py
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent

T_REF_S, T_END_S = 900.0, 1800.0            # FRI reference time, evaluation window (as in risk_model.py)
F400_MELT = 0.50                            # f400,air below this: non-charring (melt-back) type
F400_SUSTAIN = 0.65                         # f400,air at or above this: self-sustained burning expected (provisional, sc)
SSR_SUSTAIN, SUSTAIN_LEVELS = 1.0, 3        # flag B: dT(1200-1500)/dT(900-1200) >= 1 at >= 3 of 6 heights
FRI_P = ((2.0, 4), (1.0, 3), (0.5, 2))      # FRI bound -> P (same ladder as risk_model.p_class)
BLIND_FIT_TESTS = ("PIR_SC-2", "PF_out-1", "PF_out-2", "PF_in-1", "PF_in-2")   # tests the published prediction used
BLIND_COMMIT = "81ea74e"

MATS = {   # key: (Korean label, English label, colour, marker, grade)
    "EPS_C": ("EPS 가연성", "EPS combustible", "#eda100", "o", "c"),
    "EPS_SC": ("EPS 준불연", "EPS semi-noncomb.", "#e87ba4", "s", "sc"),
    "PIR_C": ("PIR(PUR) 가연성", "PIR(PUR) combustible", "#8c564b", "P", "c"),
    "PIR_SC": ("PIR(PUR) 준불연", "PIR(PUR) semi-noncomb.", "#008300", "^", "sc"),
    "PF_out": ("PF 준불연 외부용", "PF exterior-use", "#2a78d6", "D", "sc"),
    "PF_in": ("PF 준불연 내부용", "PF interior-use", "#4a3aa7", "v", "sc"),
}
SURFACE, INK, INK2, MUTED = "#fcfcfb", "#0b0b0b", "#52514e", "#898781"
GRID, AXIS = "#e1e0d9", "#c3c2b7"
P_FILL = {1: "#eef3ec", 2: "#f6f1e1", 3: "#f7e6dc", 4: "#f3d9d9"}


def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def p_from_t600(t):
    if not np.isfinite(t) or t > T_END_S:
        return 1
    fri = T_REF_S / t
    return next((p for bound, p in FRI_P if fri >= bound), 1)


def r_bounds(t_ig, C):
    """r_max,air at which the predicted t600 crosses the P-class boundaries (FRI = 0.5, 1, 2)."""
    return {p: C / (T_REF_S / bound - t_ig) for bound, p in FRI_P}


def fit(tests):
    """System constants from charring sc-foam tests: t_ig,sys = median t_ig, C_sys = mean(t_sp * r_max)."""
    c = tests["t_spread_s"] * tests["r_max_air"]
    return float(np.median(tests["t_ig_s"])), float(c.mean()), float(c.std(ddof=1)) if len(c) > 1 else np.nan


def predict(r, t_ig, C, C_sd):
    t = t_ig + C / r
    return dict(t_sp_hat_s=C / r, t600_hat_s=t, t600_hat_lo_s=t_ig + (C - C_sd) / r, t600_hat_hi_s=t_ig + (C + C_sd) / r,
                alpha600_hat_K_s2=600 / t ** 2, FRI_hat=T_REF_S / t, P_hat=p_from_t600(t))


# ----------------------------------------------------------------------------- figures

def style():
    load_module("tga_analysis", ROOT / "tga_dsc" / "analysis.py").style()


def fig_spread_model(char, t_ig, C, C_sd, blind, loo, path):
    fig, axs = plt.subplots(1, 2, figsize=(14, 5.8))
    ax = axs[0]
    r = np.linspace(0.3, 1.4, 200)
    ax.fill_between(r, (C - C_sd) / r, (C + C_sd) / r, color=GRID, lw=0)
    ax.plot(r, C / r, color=INK2, lw=1.2)
    ax.text(0.98, 0.97, f"semi-noncombustible charring foams:\nt_sp = C_sys / r_max,air\n"
            f"C_sys = {C:.0f} ± {C_sd:.0f} s·%/°C (mean ± sd, {len(char)} tests)",
            transform=ax.transAxes, ha="right", va="top", fontsize=9, color=INK)
    for _, row in char.iterrows():
        k = row["material"]
        ax.plot(row["r_max_air"], row["t_spread_s"], MATS[k][3], color=MATS[k][2], ms=10, mec=SURFACE, mew=1.2)
    for k, dx, dy in (("PF_in", -12, 8), ("PIR_SC", 10, 4), ("PF_out", 10, -4)):
        sub = char[char.material == k]
        ax.annotate(MATS[k][1], (sub["r_max_air"].iloc[0], sub["t_spread_s"].max()), xytext=(dx, dy),
                    textcoords="offset points", fontsize=9, color=INK, ha="right" if dx < 0 else "left")
    b = blind.iloc[0]
    col = MATS["PIR_C"][2]
    ax.plot(b["r_max_air"], b["t_sp_hat_s"], MATS["PIR_C"][3], ms=11, mfc=SURFACE, mec=col, mew=1.8)
    ax.plot(b["r_max_air"], b["t600_measured_s"], MATS["PIR_C"][3], ms=11, color=col, mec=SURFACE, mew=1.2)
    ax.annotate("", xy=(b["r_max_air"], b["t600_measured_s"] + 30), xytext=(b["r_max_air"], b["t_sp_hat_s"] - 30),
                arrowprops=dict(arrowstyle="->", color=col, lw=1.2))
    ax.annotate(f"PIR(PUR) combustible\nblind prediction (hollow): t_sp {b['t_sp_hat_s']:.0f} s\n"
                f"measured (filled): whole cavity at 600 K\nin {b['t600_measured_s']:.0f} s, no spread phase",
                (b["r_max_air"], b["t_sp_hat_s"]), xytext=(-12, -40), textcoords="offset points", fontsize=8.5,
                color=INK, va="top", ha="right")
    ax.set_xlim(0.3, 1.4)
    ax.set_ylim(0, 1400)
    ax.set_xlabel("TGA in air: peak mass-loss rate r_max (% of total loss per °C)")
    ax.set_ylabel("Cavity spread phase t_sp = t600 − t_ig (s)")
    ax.set_title("(a) Spread phase × oxidative reactivity ≈ constant (sc charring foams)")

    ax = axs[1]
    for p, lo, hi in ((4, 0, 450), (3, 450, 900), (2, 900, 1800)):
        ax.axhspan(lo, hi, color=P_FILL[p], lw=0)
        ax.text(1880, (lo + hi) / 2, f"P = {p}", va="center", ha="right", fontsize=8.5, color=INK2)
    ax.plot([0, 1900], [0, 1900], color=INK2, lw=0.9, ls=(0, (4, 3)))
    for _, row in loo.iterrows():
        k = row["material"]
        ax.plot(row["t600_s"], row["t600_hat_loo_s"], MATS[k][3], color=MATS[k][2], ms=10, mec=SURFACE, mew=1.2)
    for k, dx, dy, ha in (("PF_in", -6, 14, "right"), ("PIR_SC", 10, -22, "left"), ("PF_out", 10, -22, "left")):
        sub = loo[loo.material == k]
        x = sub["t600_s"].min() if dx < 0 else sub["t600_s"].mean()
        ax.annotate(f"{MATS[k][1]}\n{sub['err_loo_pct'].iloc[0]:+.0f} %", (x, sub["t600_hat_loo_s"].iloc[0]),
                    xytext=(dx, dy), textcoords="offset points", fontsize=8.5, color=INK, ha=ha)
    ax.plot(b["t600_measured_s"], b["t600_hat_s"], MATS["PIR_C"][3], ms=11, mfc=SURFACE, mec=col, mew=1.8)
    ax.annotate(f"PIR(PUR) combustible, blind test:\npredicted {b['t600_hat_s']:.0f} s (P̂ = {int(b['P_hat'])}), "
                f"measured {b['t600_measured_s']:.0f} s (P = {int(b['P_measured'])})",
                (b["t600_measured_s"], b["t600_hat_s"]), xytext=(14, 120), textcoords="offset points", fontsize=8.5,
                color=INK, va="center", arrowprops=dict(arrowstyle="-", color=MATS["PIR_C"][2], lw=0.8))
    ax.set_xlim(0, 1900)
    ax.set_ylim(0, 1900)
    ax.set_aspect("equal")
    ax.set_xlabel("Measured t600 at 3000 mm (s)")
    ax.set_ylabel("Predicted t600 from TGA (s)")
    ax.set_title("(b) Filled: leave-one-material-out (sc); hollow: blind test (c)")
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)


def fig_screening(screen, bounds, path):
    fig, ax = plt.subplots(figsize=(12, 6.6))
    xmax, ymin, ymax = 2.7, 0.0, 0.9
    ax.axhspan(ymin, F400_MELT, color=GRID, lw=0)
    ax.text(xmax - 0.03, F400_MELT - 0.02, "non-charring (melt-back) type: spread formula not applicable,\n"
            "abrupt flash (flag A) to be checked in the cavity test", ha="right", va="top", fontsize=8.5, color=INK2)
    edges = [0.0, bounds[2], bounds[3], bounds[4], xmax]
    for p, lo, hi in zip((1, 2, 3, 4), edges[:-1], edges[1:]):
        ax.fill_between([lo, hi], F400_MELT, ymax, color=P_FILL[p], lw=0)
        ax.text((lo + min(hi, xmax)) / 2, ymax - 0.012, f"P̂ = {p}", ha="center", va="top", fontsize=9, color=INK2)
    for p in (2, 3, 4):
        ax.axvline(bounds[p], ymin=(F400_MELT - ymin) / (ymax - ymin), color=AXIS, lw=0.9)
        ax.text(bounds[p], F400_MELT + 0.008, f" {bounds[p]:.2f}", fontsize=8, color=INK2, va="bottom")
    ax.plot([0, xmax], [F400_SUSTAIN] * 2, color=INK2, lw=0.9, ls=(0, (4, 3)))
    ax.text(xmax - 0.03, F400_SUSTAIN + 0.008, "above: self-sustained burning\nafter the burner stops (B̂, provisional)",
            ha="right", va="bottom", fontsize=8.5, color=INK2)
    ax.text(xmax - 0.03, F400_SUSTAIN - 0.008, "f400 = 0.65", ha="right", va="top", fontsize=8.5, color=INK2)
    offs = {"EPS_C": (-10, 10, "right"), "EPS_SC": (10, 8, "left"), "PIR_C": (10, 26, "left"),
            "PIR_SC": (0, -34, "center"), "PF_out": (-12, 14, "right"), "PF_in": (10, -34, "left")}
    for _, s in screen.iterrows():
        k = s["material"]
        tested = s["cavity_tests"] > 0
        sc = MATS[k][4] == "sc"
        ax.errorbar(s["r_max_air"], s["f400_air"], xerr=s["r_max_air_hr"], yerr=s["f400_air_hr"], fmt="none",
                    ecolor=MATS[k][2], elinewidth=1)
        ax.plot(s["r_max_air"], s["f400_air"], MATS[k][3], ms=11, color=MATS[k][2],
                mfc=MATS[k][2] if sc else SURFACE, mec=SURFACE if sc else MATS[k][2], mew=1.2 if sc else 1.8)
        note = f"measured: {s['measured']}" if tested else "no 30-min cavity test"
        if not sc:
            note = "c grade: TGA map not applicable\n" + note
        dx, dy, ha = offs[k]
        ax.annotate(f"{MATS[k][1]}\n{note}", (s["r_max_air"], s["f400_air"]), xytext=(dx, dy), textcoords="offset points",
                    fontsize=8.5, color=INK, ha=ha)
    ax.set_xlim(0, xmax)
    ax.set_ylim(ymin, ymax)
    ax.grid(False)
    ax.set_xlabel("TGA in air: peak mass-loss rate r_max (% of total loss per °C)  — sets the predicted spread class P̂")
    ax.set_ylabel("TGA in air: share of mass loss above 400 °C, f400")
    ax.set_title("TGA screening map for semi-noncombustible (sc) products in the 100 mm cavity system "
                 "(filled: sc, hollow: c grade)")
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)


# ----------------------------------------------------------------------------- main

def main():
    inv = pd.read_csv(ROOT / "coupled_analysis" / "results" / "ignition_spread_inverse.csv")
    bo = pd.read_csv(ROOT / "coupled_analysis" / "results" / "burner_off.csv")
    rm = pd.read_csv(ROOT / "cavity_fire_index" / "results" / "risk_model.csv")
    ts = pd.read_csv(ROOT / "tga_dsc" / "results" / "tga_summary.csv")
    air = ts[ts.atmosphere == "oxidative (O)"].set_index("material")
    tga = pd.DataFrame({"r_max_air": air["r_max_pct_C_mean"], "r_max_air_hr": air["r_max_pct_C_halfrange"],
                        "f400_air": air["f_above400_mean"], "f400_air_hr": air["f_above400_halfrange"]})
    tga["type"] = np.where(tga["f400_air"] < F400_MELT, "melt-back", "charring")
    tga["grade"] = [MATS[m][4] for m in tga.index]

    # ---- per test: measured quantities, flags
    t = rm[["test", "material", "t600_s", "alpha600_K_s2", "FRI_measured", "P", "D600_min", "S", "R", "risk_level",
            "abrupt_flag"]].merge(inv[["test", "flash", "t_ig_s", "t_spread_s"]], on="test", how="left")
    nb = bo.assign(up=bo["ssr"] >= SSR_SUSTAIN).groupby("test")["up"].sum().rename("levels_ssr_ge1")
    t = t.merge(nb, on="test", how="left")
    t["flag_B"] = np.where(t["levels_ssr_ge1"] >= SUSTAIN_LEVELS, "B", "")
    t["abrupt_flag"] = t["abrupt_flag"].fillna("")
    t["persist_ratio_3000"] = np.where(np.isfinite(t["t600_s"]),
                                       t["D600_min"] * 60 / (T_END_S - t["t600_s"]), np.nan)  # share of the time left
    t = t.join(tga, on="material")
    t["C_test"] = t["t_spread_s"] * t["r_max_air"]

    char = t[(t["type"] == "charring") & (t["grade"] == "sc") & t["t600_s"].notna() & ~t["flash"].fillna(False)].copy()
    t_ig, C, C_sd = fit(char)
    cv_C = C_sd / C
    cv_tsp = char["t_spread_s"].std(ddof=1) / char["t_spread_s"].mean()

    # ---- blind test: the prediction published before the PIR(PUR) c cavity data (same fit, same tests)
    tig_b, C_b, sd_b = fit(char[char.test.isin(BLIND_FIT_TESTS)])
    pc = t[t.material == "PIR_C"]
    blind = pd.DataFrame([dict(material="PIR_C", fit_tests=" ".join(BLIND_FIT_TESTS), prediction_commit=BLIND_COMMIT,
                               t_ig_sys_s=tig_b, C_sys=C_b, r_max_air=tga.loc["PIR_C", "r_max_air"],
                               **predict(tga.loc["PIR_C", "r_max_air"], tig_b, C_b, sd_b),
                               B_hat="B" if tga.loc["PIR_C", "f400_air"] >= F400_SUSTAIN else "",
                               t600_measured_s=float(pc["t600_s"].median()), t600_measured_tests="/".join(
                                   f"{x:.0f}" for x in pc["t600_s"]), P_measured=int(pc["P"].max()),
                               R_measured="/".join(str(int(x)) for x in pc["R"]),
                               flags_measured="/".join(a + b for a, b in zip(pc["abrupt_flag"], pc["flag_B"])))])

    # ---- leave one material out (sc charring)
    loo_rows = []
    for m in char["material"].unique():
        tig_m, C_m, _ = fit(char[char.material != m])
        that = tig_m + C_m / tga.loc[m, "r_max_air"]
        for _, row in char[char.material == m].iterrows():
            loo_rows.append(dict(test=row["test"], material=m, t600_s=row["t600_s"], t_ig_sys_loo_s=tig_m, C_sys_loo=C_m,
                                 t600_hat_loo_s=that, P=row["P"], P_hat_loo=p_from_t600(that)))
    loo = pd.DataFrame(loo_rows)
    med = loo.groupby("material")["t600_s"].transform("median")
    loo["err_loo_pct"] = 100 * (loo["t600_hat_loo_s"] - med) / med

    # ---- screening for all six materials (grade gate, then TGA)
    bounds = r_bounds(t_ig, C)
    srows = []
    for m in MATS:
        r, f = tga.loc[m, "r_max_air"], tga.loc[m, "f400_air"]
        tested = t[t.material == m]
        row = dict(material=m, label_ko=MATS[m][0], grade=MATS[m][4], cavity_tests=len(tested), r_max_air=r,
                   r_max_air_hr=tga.loc[m, "r_max_air_hr"], f400_air=f, f400_air_hr=tga.loc[m, "f400_air_hr"],
                   type=tga.loc[m, "type"])
        if MATS[m][4] == "c":
            row["screen_result"] = "c grade: cavity test required (PIR(PUR) c flashed over the whole cavity in 30 s)"
        elif row["type"] == "melt-back":
            row["screen_result"] = "melt-back type: cavity test required (check flag A)"
        else:
            row.update(predict(r, t_ig, C, C_sd), B_hat="B" if f >= F400_SUSTAIN else "")
            row["P_hat_range"] = f"{p_from_t600(row['t600_hat_hi_s'])}-{p_from_t600(row['t600_hat_lo_s'])}"
            row["screen_result"] = f"P̂ = {row['P_hat']}" + (", B̂" if row["B_hat"] else "")
        if len(tested):
            ps = "/".join(str(int(x)) for x in tested["P"])
            rs = "/".join(str(int(x)) for x in tested["R"])
            fl = "/".join(sorted({a + b for a, b in zip(tested["abrupt_flag"], tested["flag_B"]) if a + b})) or "-"
            row.update(P_measured=ps, R_measured=rs, flags_measured=fl, measured=f"P = {ps}, R = {rs}, flags {fl}")
        srows.append(row)
    screen = pd.DataFrame(srows)

    # ---- final per-test table
    t = t.merge(screen[["material", "t600_hat_s", "P_hat", "B_hat"]], on="material", how="left")
    t = t.merge(loo[["test", "t600_hat_loo_s", "P_hat_loo"]], on="test", how="left")
    t["flags"] = t["abrupt_flag"] + t["flag_B"]
    in_b = (t["type"] == "charring") & (t["grade"] == "sc")
    t["B_consistent"] = np.where(in_b, (t["flag_B"] == "B") == (t["B_hat"] == "B"), np.nan)

    out = HERE / "results"
    t.round(4).to_csv(out / "final_by_test.csv", index=False, encoding="utf-8")
    screen.round(4).to_csv(out / "tga_screening.csv", index=False, encoding="utf-8")
    loo.round(2).to_csv(out / "loo_validation.csv", index=False, encoding="utf-8")
    blind.round(4).to_csv(out / "blind_test_pir_c.csv", index=False, encoding="utf-8")
    pd.DataFrame([dict(t_ig_sys_s=t_ig, C_sys=C, C_sys_sd=C_sd, C_sys_cv=cv_C, t_spread_cv=cv_tsp,
                       r_bound_P2=bounds[2], r_bound_P3=bounds[3], r_bound_P4=bounds[4], n_tests=len(char),
                       tests=" ".join(char["test"]), f400_melt=F400_MELT, f400_sustain=F400_SUSTAIN)]).round(4).to_csv(
        out / "system_constants.csv", index=False, encoding="utf-8")

    style()
    fig_spread_model(char, t_ig, C, C_sd, blind, loo, HERE / "figures" / "fig1_tga_spread_model.png")
    fig_screening(screen, bounds, HERE / "figures" / "fig2_tga_screening_map.png")

    pd.set_option("display.width", 250)
    pd.set_option("display.max_columns", 40)
    print(f"t_ig,sys = {t_ig:.0f} s, C_sys = {C:.1f} ± {C_sd:.1f} (CV {cv_C:.1%}); CV of t_sp itself {cv_tsp:.1%}")
    print("r_max bounds for P̂:", {k: round(v, 3) for k, v in bounds.items()})
    print(char[["test", "material", "t_ig_s", "t_spread_s", "r_max_air", "C_test"]].round(1).to_string())
    print(blind.round(1).to_string())
    print(loo.round(1).to_string())
    print(screen.drop(columns=["label_ko"]).round(3).to_string())
    print(t[["test", "material", "t600_s", "P", "S", "R", "risk_level", "flags", "levels_ssr_ge1", "persist_ratio_3000",
             "t600_hat_s", "P_hat", "t600_hat_loo_s", "P_hat_loo", "B_hat", "B_consistent"]].round(2).to_string())


if __name__ == "__main__":
    (HERE / "figures").mkdir(parents=True, exist_ok=True)
    (HERE / "results").mkdir(parents=True, exist_ok=True)
    main()
