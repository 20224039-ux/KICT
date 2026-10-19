#!/usr/bin/env python3
"""Collect everything the TGA-DSC analysis deck shows into deck_data.json (curves downsampled to 5 °C).

    python ../qa_reliability.py && python deck_data.py && node build_deck.js
"""
from __future__ import annotations

import glob
import importlib.util
import json
import os
from pathlib import Path

import numpy as np
import pandas as pd

HERE = Path(__file__).resolve().parent
TGA = HERE.parent
spec = importlib.util.spec_from_file_location("tga_analysis", TGA / "analysis.py")
tga = importlib.util.module_from_spec(spec)
spec.loader.exec_module(tga)

GRID = np.arange(50.0, 786.0, 5.0)
GRID_DSC = np.arange(150.0, 721.0, 5.0)
ORDER = ["EPS_C", "EPS_SC", "PIR_C", "PIR_SC", "PF_out", "PF_in"]
LABEL = {"EPS_C": "EPS 가연성", "EPS_SC": "EPS 준불연", "PIR_C": "PIR(PUR) 가연성", "PIR_SC": "PIR(PUR) 준불연",
         "PF_out": "PF 준불연 외부용", "PF_in": "PF 준불연 내부용"}
ATM = {"N2": "N₂", "oxidative (O)": "공기(O)"}


def rnd(x, n=4):
    if isinstance(x, (list, tuple, np.ndarray)):
        return [rnd(v, n) for v in x]
    if x is None or (isinstance(x, float) and not np.isfinite(x)):
        return None
    return round(float(x), n)


def main():
    curves = {m: {} for m in ORDER}
    dsc = {m: {} for m in ORDER}
    for atm in ("N", "O"):
        for f in sorted(glob.glob(os.path.join(TGA / "data", atm, "*.csv"))):
            stem, rep = tga.stem(os.path.basename(f))
            key = tga.MATERIALS[stem][0]
            run = tga.process(f, atm)
            name = f"{'N2' if atm == 'N' else 'O'}#{rep}"
            curves[key][name] = dict(rem=rnd(np.interp(GRID, run["T"], run["rem"]), 2),
                                     dtg=rnd(np.interp(GRID, run["T"], run["dtg"]), 4))
            dsc[key][name] = rnd(np.interp(GRID_DSC, run["T"], run["h"]), 3)
    res = TGA / "results"
    summ = pd.read_csv(res / "tga_summary.csv")
    qa = pd.read_csv(res / "qa_runs.csv")
    reps = pd.read_csv(res / "qa_replicates.csv")
    ep = pd.read_csv(res / "qa_endpoint.csv")
    vd = pd.read_csv(res / "qa_verdict.csv")
    cc = pd.read_csv(res / "qa_crosscheck.csv")
    link = pd.read_csv(res / "fire_link.csv")

    table = []
    for m in ORDER:
        for atm in ("N2", "oxidative (O)"):
            s = summ[(summ.material == m) & (summ.atmosphere == atm)].iloc[0]
            table.append(dict(material=m, label=LABEL[m], atm=ATM[atm],
                              **{k: rnd(s[f"{k}_mean"], 4) for k in ("T_a05_C", "T_onset_C", "T_peak_C", "r_max_pct_C",
                                                                     "T_a50_C", "T_a90_C", "f_above400")},
                              **{f"{k}_hr": rnd(s[f"{k}_halfrange"], 4) for k in ("T_onset_C", "T_peak_C", "r_max_pct_C",
                                                                                  "T_a50_C", "f_above400")}))
    runs = []
    for _, r in qa.iterrows():
        runs.append(dict(material=r["material"], atm=ATM[r["atmosphere"]], rep=int(r["repeat"]),
                         **{k: rnd(r[k], 4) for k in ("w_start_reading_mg", "w_end_reading_mg", "dm_tot_mg", "iso_noise_ug",
                                                     "iso_drift_ug", "snr_loss_to_noise", "apparent_gain_below_200C_pct",
                                                     "end_rate_pct_of_peak", "heating_rate_C_min", "lag_median_C",
                                                     "self_heating_C", "gas_flow_min", "gas_flow_max", "hf_at_100C_mW",
                                                     "hf_at_700C_mW", "dsc_event_C", "dsc_event_mW", "alpha_at_150C_pct",
                                                     "r_max_pct_C", "r_max_check_pct_C", "T_peak_C", "T_peak_check_C")}))
    endpoint = []
    for (m, atm, e), g in ep.groupby(["material", "atmosphere", "endpoint_C"]):
        endpoint.append(dict(material=m, atm=ATM[atm], endpoint=float(e),
                             **{k: rnd(g[k].mean(), 4) for k in ("T_a90_C", "r_max_pct_C", "f_above400")}))
    verdict = [dict(material=r["material"], label=LABEL[r["material"]], atm=ATM[r["atmosphere"]],
                    **{k: (rnd(r[k], 4) if isinstance(r[k], (int, float, np.floating)) else r[k]) for k in vd.columns
                       if k not in ("material", "atmosphere")}) for _, r in vd.iterrows()]
    separation = [dict(material=r["material"], atm=ATM[r["atmosphere"]], sep_r_max=rnd(r["sep_r_max_pct_C"], 3),
                       sep_f400=rnd(r["sep_f_above400"], 3), gap_r_max=rnd(r["gap_r_max_pct_C"], 4),
                       d_r_max=rnd(r["d_r_max_pct_C"], 4)) for _, r in reps.iterrows()]
    cav = [dict(material=r["material"], t600=rnd(r["t600_median_s"], 1),
                r_max_air=rnd(r["O_r_max_pct_C"], 4)) for _, r in link.iterrows()]
    spear = [dict(check=r["check"], endpoint=r["material"], rho=rnd(r["value"], 3))
             for _, r in cc[cc.check.str.startswith("Spearman")].iterrows()]
    xcheck = cc[~cc.check.str.startswith("Spearman")]
    xsum = []
    for chk, g in xcheck.groupby("check", sort=False):
        xsum.append(dict(check=chk, max_abs_diff=rnd(g["diff"].abs().max(), 4),
                         min_pct=rnd(g["diff_pct"].min(), 2), max_pct=rnd(g["diff_pct"].max(), 2)))
    hf_drift = (qa["hf_at_700C_mW"] - qa["hf_at_100C_mW"])
    stats = dict(
        w_start_min=rnd(qa.w_start_reading_mg.min(), 3), w_start_max=rnd(qa.w_start_reading_mg.max(), 3),
        w_end_min=rnd(qa.w_end_reading_mg.min(), 3), w_end_max=rnd(qa.w_end_reading_mg.max(), 3),
        n_negative_start=int((qa.w_start_reading_mg < 0).sum()), n_negative_end=int((qa.w_end_reading_mg < 0).sum()),
        dm_min=rnd(qa.dm_tot_mg.min(), 3), dm_max=rnd(qa.dm_tot_mg.max(), 3),
        noise_min=rnd(qa.iso_noise_ug.min(), 2), noise_max=rnd(qa.iso_noise_ug.max(), 2),
        snr_min=rnd(qa.snr_loss_to_noise.min(), 0), snr_max=rnd(qa.snr_loss_to_noise.max(), 0),
        rate_min=rnd(qa.heating_rate_C_min.min(), 3), rate_max=rnd(qa.heating_rate_C_min.max(), 3),
        lag_min=rnd(qa.lag_median_C.min(), 2), lag_max=rnd(qa.lag_median_C.max(), 2),
        hf_drift_min=rnd(hf_drift.min(), 1), hf_drift_max=rnd(hf_drift.max(), 1),
        balance_step_ug=rnd(qa.balance_step_ug.median(), 3),
        flow_mean_min=rnd(qa.gas_flow_mean.min(), 2), flow_mean_max=rnd(qa.gas_flow_mean.max(), 2),
    )
    out = dict(grid=GRID.tolist(), grid_dsc=GRID_DSC.tolist(), order=ORDER, label=LABEL, curves=curves, dsc=dsc,
               table=table, runs=runs, endpoint=endpoint, verdict=verdict, separation=separation, cavity=cav,
               spearman=spear, crosscheck=xsum, stats=stats)
    (HERE / "deck_data.json").write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(stats, ensure_ascii=False, indent=1))
    print(json.dumps(xsum, ensure_ascii=False, indent=1))
    print(json.dumps(spear, ensure_ascii=False))


if __name__ == "__main__":
    main()
