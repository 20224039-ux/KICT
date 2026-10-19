#!/usr/bin/env python3
"""Reliability audit of the raw TGA-DSC exports (24 runs): what the data can and cannot support.

Per run (results/qa_runs.csv)
    mass signal   start / end reading, total loss, isothermal drift and noise, balance step, signal-to-noise,
                  apparent gain before decomposition (buoyancy / no tare), loss rate left at the ramp end
    temperature   actual heating rate, sample-minus-program lag, self-heating overshoot (sample runs ahead)
    gas, DSC      flow reading; heat-flow calibration factor and baseline columns; heat-flow drift
Per material x atmosphere (results/qa_replicates.csv)
    replicate difference of each metric next to the gap to the neighbouring material it is ranked against
Endpoint sensitivity (results/qa_endpoint.csv)
    metrics normalised to the loss up to 700 / 750 / 790 °C instead of the ramp end
Cross-checks (results/qa_crosscheck.csv) and verdicts (results/qa_verdict.csv)
    peak rate recomputed with a plain 10 °C central difference; rank correlation with the cavity t600 for each
    endpoint; per material x atmosphere verdict on each check (thresholds below are this analysis's own criteria)
    and an overall TGA verdict: reliable / conditional / endpoint (normalised metrics depend on the endpoint) /
    remeasure

    python qa_reliability.py            # needs data/N/*.csv and data/O/*.csv
"""
from __future__ import annotations

import glob
import importlib.util
import os
from pathlib import Path

import numpy as np
import pandas as pd
from scipy.ndimage import median_filter
from scipy.signal import savgol_filter

HERE = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location("tga_analysis", HERE / "analysis.py")
tga = importlib.util.module_from_spec(spec)
spec.loader.exec_module(tga)

NOMINAL_RATE = 5.0          # °C/min
ENDPOINTS_C = (700.0, 750.0, 790.0)
# verdict thresholds (good if better than the first value, poor if worse than the second) - own criteria, not a standard
SIGNAL_MG = (1.0, 0.5)      # total mass loss of the smaller replicate
GAIN_PCT = (2.0, 5.0)       # apparent mass gain below 200 °C, % of the total loss (buoyancy, no blank run)
END_PCT = (5.0, 20.0)       # loss rate left at the ramp end, % of the peak rate
TEMP_REP_C = (10.0, 25.0)   # replicate difference of T_onset, T_peak, T_a50 (largest of the three)
RATE_REP = (0.10, 0.20)     # replicate difference of r_max / mean
SEP = (3.0, 1.5)            # ranking: gap to the neighbouring material / replicate difference
DSC_EVENT_C = 10.0          # a DSC event counts as reproduced if its temperature agrees within this
DSC_EVENT_MW = 0.5          # ... and both replicates show at least this local heat flow (150-720 °C, away from the ramp ends)
CAVITY_T600 = {"EPS_SC": 975.0, "PIR_C": 30.0, "PIR_SC": 1079.5, "PF_out": 1350.0, "PF_in": 514.0}   # median s
SC_MATERIALS = ("EPS_SC", "PIR_SC", "PF_out", "PF_in")


def load(path):
    d = pd.read_csv(path)
    d.columns = [c.strip() for c in d.columns]
    return d.loc[:, ~d.columns.str.startswith("Unnamed")]


def audit(path, atm):
    d = load(path)
    t = d["Time"].to_numpy()                       # min
    Tp, Ts = d["Program Temperature"].to_numpy(), d["Sample Temperature"].to_numpy()
    w, hf = d["Unsubtracted Weight"].to_numpy(), d["Unsubtracted Heat Flow"].to_numpy()
    i0 = int(np.argmax(Tp > Tp[0] + 0.5))
    i1 = int(np.argmax(Tp >= Tp.max() - 0.01))
    dt_s = float(np.median(np.diff(t)) * 60)
    win = max(int(round(30 / dt_s)) | 1, 5)        # ~30 s smoothing window, odd
    # isothermal hold before the ramp: drift and noise of the balance
    iso = slice(int(0.1 * i0), i0)                 # skip the first 10 % (settling)
    wi = w[iso]
    trend = np.polyval(np.polyfit(t[iso], wi, 1), t[iso])
    iso_noise_ug = float(np.std(wi - trend) * 1e3)
    iso_drift_ug = float((trend[-1] - trend[0]) * 1e3)
    steps = np.diff(np.unique(np.round(w, 6)))
    step_ug = float(np.min(steps[steps > 0]) * 1e3) if (steps > 0).any() else np.nan
    # ramp
    sl = slice(i0, i1 + 1)
    ws = savgol_filter(w[sl], win * 8 + 1, 2)
    Tss = savgol_filter(Ts[sl], win * 8 + 1, 2)
    tr = t[sl]
    run = tga.process(path, atm)
    m = run["metrics"]
    T, a, r = run["T"], run["a"], run["r"]
    dm = m["dm_tot_mg"]
    w_ref = float(np.mean(ws[:200]))
    gain = float(np.max(ws[(Tss > Tss[0] + 5) & (Tss < 200)] - w_ref)) if ((Tss > Tss[0] + 5) & (Tss < 200)).any() else 0.0
    end_rate = float(np.mean(r[T > T[-1] - 20]) / np.max(r[(T >= 100) & (T <= tga.PEAK_TMAX_C)]))
    rate = np.polyfit(tr, Ts[sl], 1)[0]
    inst = np.gradient(Tss, tr)
    lag = Ts[sl] - Tp[sl]
    lag_med = float(np.median(lag))
    lag_s = savgol_filter(lag, win + (1 - win % 2), 1)
    # self-heating: sample running ahead of its own lag trend (~10 min running median), above 150 °C so the
    # ramp-start transient is not counted
    dec = max(int(round(1.0 / dt_s)), 1)           # ~1 Hz
    lag_d, T_d = lag_s[::dec], Tss[::dec]
    dev = lag_d - median_filter(lag_d, size=601, mode="nearest")
    hot = np.flatnonzero(T_d > 150)
    k_sh = int(hot[np.argmax(dev[hot])])
    overshoot, t_over = float(dev[k_sh]), float(T_d[k_sh])
    gas = d["Approx. Gas Flow"].to_numpy()
    cal = d["Heat Flow Calibration"].to_numpy()
    base_w, base_hf = d["Baseline Weight"].to_numpy(), d["Baseline Heat Flow"].to_numpy()
    hfs = savgol_filter(hf[sl], win * 8 + 1, 2)
    hf_at = lambda x: float(np.interp(x, Tss, hfs))   # noqa: E731
    mat, rep = tga.stem(os.path.basename(path))
    # independent check of the peak rate: plain central difference over +-5 °C on the same conversion curve
    dT = 10.0
    a_hi, a_lo = np.interp(T + dT / 2, T, a), np.interp(T - dT / 2, T, a)
    r_cd = (a_hi - a_lo) / dT
    band = (T >= 100 + dT) & (T <= tga.PEAK_TMAX_C)
    k = int(np.flatnonzero(band)[np.argmax(r_cd[band])])
    win_ev = (T > 150) & (T < 720)
    j = int(np.flatnonzero(win_ev)[np.argmax(np.abs(run["h"][win_ev]))])
    return dict(
        atmosphere=tga.ATM[atm], material=tga.MATERIALS[mat][0], repeat=rep, file=os.path.basename(path),
        sample_interval_s=round(dt_s, 3), iso_min=float(t[i0]), T_program_start_C=float(Tp[0]),
        T_sample_start_C=float(Ts[0]), w_start_reading_mg=float(w[0]), w_end_reading_mg=float(np.mean(w[i1 - 60:i1])),
        dm_tot_mg=dm, balance_step_ug=step_ug, iso_noise_ug=iso_noise_ug, iso_drift_ug=iso_drift_ug,
        snr_loss_to_noise=dm * 1e3 / iso_noise_ug if iso_noise_ug > 0 else np.nan,
        apparent_gain_below_200C_pct=100 * gain / dm, end_rate_pct_of_peak=100 * end_rate,
        heating_rate_C_min=float(rate), heating_rate_sd_C_min=float(np.std(inst[(tr > tr[0] + 2) & (tr < tr[-1] - 2)])),
        lag_median_C=lag_med, self_heating_C=overshoot, self_heating_at_C=t_over,
        gas_flow_mean=float(np.mean(gas)), gas_flow_min=float(np.min(gas)), gas_flow_max=float(np.max(gas)),
        hf_calibration_values=" ".join(f"{v:g}" for v in np.unique(np.round(cal, 6))[:3]),
        baseline_columns_zero=bool(np.all(base_w == 0) and np.all(base_hf == 0)),
        hf_at_100C_mW=hf_at(100), hf_at_700C_mW=hf_at(700), dsc_exo_mW=m["dsc_exo_mW"], dsc_exo_peak_C=m["dsc_exo_peak_C"],
        alpha_at_150C_pct=100 * float(np.interp(150.0, T, a)),
        r_max_check_pct_C=float(r_cd[k] * 100), T_peak_check_C=float(T[k]),
        dsc_event_C=float(T[j]), dsc_event_mW=float(run["h"][j]),
        **{k2: m[k2] for k2 in ("T_a05_C", "T_a50_C", "T_a90_C", "T_onset_C", "T_peak_C", "r_max_pct_C", "f_above400")},
    ), run


def endpoint_metrics(run, end_c):
    """Metrics with the conversion normalised to the loss up to end_c instead of the ramp end."""
    T, a = run["T"], run["a"]
    sel = T <= end_c
    T2, a2 = T[sel], a[sel] / np.interp(end_c, T, a)
    r2 = savgol_filter(np.gradient(a2, T2), tga.SMOOTH_DTG, 2)
    band = (T2 >= 100)
    at = lambda x: float(T2[np.argmax(a2 >= x)]) if (a2 >= x).any() else np.nan  # noqa: E731
    return dict(T_a50_C=at(0.5), T_a90_C=at(0.9), r_max_pct_C=float(np.max(r2[band]) * 100),
                f_above400=float(1 - np.interp(tga.HT_SPLIT_C, T2, a2)))


def main():
    rows, runs = [], {}
    for atm in ("N", "O"):
        for f in sorted(glob.glob(os.path.join(HERE / "data", atm, "*.csv"))):
            row, run = audit(f, atm)
            rows.append(row)
            runs[(row["atmosphere"], row["material"], row["repeat"])] = run
    q = pd.DataFrame(rows)
    out = HERE / "results"
    q.round(4).to_csv(out / "qa_runs.csv", index=False, encoding="utf-8")

    keys = ["T_a05_C", "T_a50_C", "T_a90_C", "T_onset_C", "T_peak_C", "r_max_pct_C", "f_above400", "dsc_exo_peak_C"]
    reps = []
    for (atm, mat), g in q.groupby(["atmosphere", "material"], sort=False):
        g = g.sort_values("repeat")
        reps.append(dict(atmosphere=atm, material=mat, dm_tot_mg_min=g["dm_tot_mg"].min(),
                         snr_min=g["snr_loss_to_noise"].min(),
                         **{f"d_{k}": float(abs(g[k].iloc[0] - g[k].iloc[1])) for k in keys},
                         **{f"mean_{k}": float(g[k].mean()) for k in keys}))
    rp = pd.DataFrame(reps)
    # separation of neighbouring materials in the ranking the conclusions use (air r_max and f400)
    for k in ("r_max_pct_C", "f_above400", "T_a50_C", "T_peak_C"):
        for atm in rp["atmosphere"].unique():
            sub = rp[rp.atmosphere == atm].sort_values(f"mean_{k}")
            vals, spread = sub[f"mean_{k}"].to_numpy(), sub[f"d_{k}"].to_numpy()
            gap = np.full(len(sub), np.nan)
            for i in range(len(sub)):
                near = [abs(vals[i] - vals[j]) for j in (i - 1, i + 1) if 0 <= j < len(sub)]
                gap[i] = min(near)
            rp.loc[sub.index, f"gap_{k}"] = gap
            rp.loc[sub.index, f"sep_{k}"] = gap / np.maximum(spread, 1e-9)
    rp.round(4).to_csv(out / "qa_replicates.csv", index=False, encoding="utf-8")

    ep = []
    for (atm, mat, rep), run in runs.items():
        for e in ENDPOINTS_C:
            ep.append(dict(atmosphere=atm, material=mat, repeat=rep, endpoint_C=e, **endpoint_metrics(run, e)))
    ep = pd.DataFrame(ep)
    ep.round(4).to_csv(out / "qa_endpoint.csv", index=False, encoding="utf-8")

    # cross-checks: summary table vs runs, peak rate by a second method, rank correlation with the cavity tests
    from scipy import stats
    summ = pd.read_csv(out / "tga_summary.csv")
    cc = []
    for _, g in q.groupby(["atmosphere", "material"], sort=False):
        s_row = summ[(summ.atmosphere == g.atmosphere.iloc[0]) & (summ.material == g.material.iloc[0])].iloc[0]
        for k2 in ("r_max_pct_C", "T_peak_C", "f_above400", "T_a50_C"):
            cc.append(dict(check=f"summary mean {k2}", atmosphere=g.atmosphere.iloc[0], material=g.material.iloc[0],
                           value=float(g[k2].mean()), reference=float(s_row[f"{k2}_mean"])))
        cc.append(dict(check="r_max: Savitzky-Golay vs 10 °C central difference", atmosphere=g.atmosphere.iloc[0],
                       material=g.material.iloc[0], value=float(g["r_max_check_pct_C"].mean()), reference=float(g["r_max_pct_C"].mean())))
        cc.append(dict(check="T_peak: Savitzky-Golay vs 10 °C central difference", atmosphere=g.atmosphere.iloc[0],
                       material=g.material.iloc[0], value=float(g["T_peak_check_C"].mean()), reference=float(g["T_peak_C"].mean())))
    cc = pd.DataFrame(cc)
    cc["diff"] = cc["value"] - cc["reference"]
    cc["diff_pct"] = 100 * cc["diff"] / cc["reference"].abs()
    rho = []
    air = "oxidative (O)"
    for e in ENDPOINTS_C:
        sub = ep[(ep.atmosphere == air) & (ep.endpoint_C == e)].groupby("material")["r_max_pct_C"].mean()
        for name, mats in (("sc 4 materials", SC_MATERIALS), ("5 materials incl. PIR_C", tuple(CAVITY_T600))):
            x = [sub[mt] for mt in mats]
            y = [-CAVITY_T600[mt] for mt in mats]
            rho.append(dict(check=f"Spearman r_max(air) vs spread speed, {name}", atmosphere=air, material=f"endpoint {e:.0f} °C",
                            value=float(stats.spearmanr(x, y).statistic), reference=np.nan))
    cc = pd.concat([cc, pd.DataFrame(rho)], ignore_index=True)
    cc.round(4).to_csv(out / "qa_crosscheck.csv", index=False, encoding="utf-8")

    # verdict per material x atmosphere
    def grade(v, lim, higher_is_better=False):
        if not np.isfinite(v):
            return "n/a"
        if higher_is_better:
            return "good" if v >= lim[0] else "fair" if v >= lim[1] else "poor"
        return "good" if v <= lim[0] else "fair" if v <= lim[1] else "poor"
    vd = []
    for (atm, mat), g in q.groupby(["atmosphere", "material"], sort=False):
        g = g.sort_values("repeat")
        r = rp[(rp.atmosphere == atm) & (rp.material == mat)].iloc[0]
        temp_rep = float(max(r["d_T_onset_C"], r["d_T_peak_C"], r["d_T_a50_C"]))
        rate_rep = float(r["d_r_max_pct_C"] / r["mean_r_max_pct_C"])
        ev_dT = float(abs(g["dsc_event_C"].iloc[0] - g["dsc_event_C"].iloc[1]))
        ev_same_sign = bool(np.sign(g["dsc_event_mW"].iloc[0]) == np.sign(g["dsc_event_mW"].iloc[1]))
        row = dict(atmosphere=atm, material=mat,
                   signal_mg=float(g["dm_tot_mg"].min()), gain_pct=float(g["apparent_gain_below_200C_pct"].max()),
                   end_rate_pct=float(g["end_rate_pct_of_peak"].max()), temp_rep_C=temp_rep, rate_rep=rate_rep,
                   sep_r_max=float(r["sep_r_max_pct_C"]), sep_f400=float(r["sep_f_above400"]),
                   dsc_event=f"{g['dsc_event_C'].iloc[0]:.0f}/{g['dsc_event_C'].iloc[1]:.0f} °C, "
                             f"{g['dsc_event_mW'].iloc[0]:+.1f}/{g['dsc_event_mW'].iloc[1]:+.1f} mW",
                   alpha150_pct=float(g["alpha_at_150C_pct"].mean()))
        row.update(v_signal=grade(row["signal_mg"], SIGNAL_MG, True), v_gain=grade(row["gain_pct"], GAIN_PCT),
                   v_end=grade(row["end_rate_pct"], END_PCT), v_temp_rep=grade(temp_rep, TEMP_REP_C),
                   v_rate_rep=grade(rate_rep, RATE_REP), v_rank=grade(row["sep_r_max"], SEP, True),
                   v_dsc="fair" if (ev_dT <= DSC_EVENT_C and ev_same_sign
                                    and g["dsc_event_mW"].abs().min() >= DSC_EVENT_MW) else "poor")
        tga_checks = [row[k] for k in ("v_signal", "v_gain", "v_end", "v_temp_rep", "v_rate_rep", "v_rank")]
        # overall (TGA only): remeasure = too little or distorted signal; endpoint = loss still running at the ramp end,
        # so normalised metrics depend on where the loss is cut; conditional = some check only fair
        row["overall"] = ("remeasure" if "poor" in (row["v_signal"], row["v_gain"], row["v_rate_rep"])
                          else "endpoint" if row["v_end"] == "poor"
                          else "conditional" if any(v != "good" for v in tga_checks) else "reliable")
        vd.append(row)
    vd = pd.DataFrame(vd)
    vd.round(4).to_csv(out / "qa_verdict.csv", index=False, encoding="utf-8")

    pd.set_option("display.width", 250)
    pd.set_option("display.max_columns", 60)
    print(q.drop(columns=["file"]).round(3).to_string())
    print(rp.round(3).to_string())
    piv = ep.groupby(["atmosphere", "material", "endpoint_C"])[["T_a90_C", "r_max_pct_C", "f_above400"]].mean().unstack()
    print(piv.round(3).to_string())
    print(cc.round(4).to_string())
    print(vd.round(3).to_string())


if __name__ == "__main__":
    main()
