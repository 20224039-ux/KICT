#!/usr/bin/env python3
"""TGA -> PyroSim (FDS) material inputs: what the TGA-DSC runs can give, stage by stage.

FDS derives A and E of a first-order reaction from REFERENCE_TEMPERATURE (peak of the normalised mass loss
rate in a TGA run) and REFERENCE_RATE (that peak rate per unit mass of the reacting component, 1/s) at the
TGA HEATING_RATE (FDS read.f90: E = e r_p R T_p^2 / Tdot, A = e r_p exp(E / R T_p); User's Guide,
"Estimating Kinetic Parameters"). The conversion used here is normalised to the loss up to 790 °C, so a stage
carrying a share f of that loss has its own peak conversion rate r_p / f, and REFERENCE_RATE = Tdot r_p / f.
That needs no initial sample mass. The residue yield (NU_MATL) does, and the density, the specific heat and
the heat of reaction cannot come from these runs at all (no sample mass, no blank, no heat-flow calibration).

Writes
    results/pyrosim_reactions.csv   one row per material x atmosphere x stage (mean of the two runs and their range)
    results/pyrosim_tga_curves.csv  mean measured curves (1 - alpha, d alpha/dt) to compare with an FDS
                                    TGA_ANALYSIS run (normalise the FDS mass the same way, see the note in the file)
    report/pyrosim_data.json        the same numbers for the PyroSim deck

    python pyrosim_inputs.py
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
spec = importlib.util.spec_from_file_location("tga_analysis", HERE / "analysis.py")
tga = importlib.util.module_from_spec(spec)
spec.loader.exec_module(tga)

BETA_K_MIN = 5.0                    # TGA heating rate = FDS HEATING_RATE (K/min)
BETA = BETA_K_MIN / 60.0            # K/s
R = 8.314462618                     # J/(mol K)
X_O2_AIR = 0.209                    # oxygen volume fraction if the 'O' gas was air (unconfirmed)
GRID = np.arange(60.0, 786.0, 1.0)  # °C
T_START = 150.0                     # loss below this (moisture, volatiles) is reported apart from the stages
COMPLETE = 0.2                      # a stage counts as finished if its rate at the stage end is below this share of its peak
ORDER = ["EPS_C", "EPS_SC", "PIR_C", "PIR_SC", "PF_out", "PF_in"]
LABEL = {key: ko for key, ko, _ in tga.MATERIALS.values()}
ATM_KO = {"N": "N₂", "O": "공기(O)"}
# DTG valley that ends each stage: (lower, upper) search window, °C. Chosen from the mean DTG curves; stages run
# from T_START to the first valley, and from the last valley to the ramp end.
VALLEYS = {
    ("EPS_C", "N"): [], ("EPS_SC", "N"): [],
    ("PIR_C", "N"): [(330, 700)], ("PIR_SC", "N"): [(340, 700)],
    ("PF_out", "N"): [(310, 600)], ("PF_in", "N"): [(310, 580)],
    ("EPS_C", "O"): [], ("EPS_SC", "O"): [(360, 480)],
    ("PIR_C", "O"): [(320, 480)], ("PIR_SC", "O"): [(320, 480)],
    ("PF_out", "O"): [(220, 420)], ("PF_in", "O"): [(220, 420)],
}
# how each stage maps onto a PyroSim / FDS reaction (Korean, shown in the deck and the CSV)
ROLE = {
    ("EPS_C", "N", 1): "반응 1: 열분해 (단일 단계)", ("EPS_SC", "N", 1): "반응 1: 열분해 (단일 단계)",
    ("PIR_C", "N", 1): "반응 1: 열분해 → 탄화층", ("PIR_C", "N", 2): "탄화층 열분해 (790 °C에서 진행 중)",
    ("PIR_SC", "N", 1): "반응 1: 열분해 → 탄화층", ("PIR_SC", "N", 2): "탄화층 열분해 (790 °C에서 진행 중)",
    ("PF_out", "N", 1): "반응 1: 1단계 열분해 → 탄화층", ("PF_out", "N", 2): "탄화 진행 (790 °C에서 가속 중)",
    ("PF_in", "N", 1): "반응 1: 1단계 열분해 → 탄화층", ("PF_in", "N", 2): "탄화 진행 (790 °C에서 진행 중)",
    ("EPS_C", "O", 1): "공기 중 분해 (재측정 대상)",
    ("EPS_SC", "O", 1): "공기 중 주분해", ("EPS_SC", "O", 2): "잔류물 산화 (790 °C에서 진행 중)",
    ("PIR_C", "O", 1): "공기 중 1단계 분해", ("PIR_C", "O", 2): "반응 2: 탄화층 산화 (N_O2 = 1)",
    ("PIR_SC", "O", 1): "공기 중 1단계 분해", ("PIR_SC", "O", 2): "반응 2: 탄화층 산화 (N_O2 = 1)",
    ("PF_out", "O", 1): "초기 감소 (휘발분)", ("PF_out", "O", 2): "반응 2: 탄화층 산화 (N_O2 = 1)",
    ("PF_in", "O", 1): "초기 감소 (휘발분)", ("PF_in", "O", 2): "반응 2: 탄화층 산화 (N_O2 = 1)",
}


def load_runs():
    runs = {}
    for atm in ("N", "O"):
        for f in sorted(glob.glob(os.path.join(HERE / "data", atm, "*.csv"))):
            stem, rep = tga.stem(os.path.basename(f))
            r = tga.process(f, atm)
            runs[(tga.MATERIALS[stem][0], atm, rep)] = dict(a=np.interp(GRID, r["T"], r["a"]),
                                                            r=np.interp(GRID, r["T"], r["r"]))
    return runs


def bounds(r_mean, key):
    """Stage limits: T_START, the DTG minimum inside each valley window, the ramp end."""
    cuts = [T_START]
    for lo, hi in VALLEYS[key]:
        sel = (GRID >= lo) & (GRID <= hi)
        cuts.append(float(GRID[sel][np.argmin(r_mean[sel])]))
    return cuts + [float(GRID[-1])]


def stage_numbers(a, r, lo, hi):
    """Peak temperature (°C), peak rate (1/°C), stage share f of the loss to 790 °C, rate at the stage end / peak."""
    sel = (GRID >= lo) & (GRID <= hi)
    i = int(np.flatnonzero(sel)[np.argmax(r[sel])])
    f = float(np.interp(hi, GRID, a) - np.interp(lo, GRID, a))
    end = float(np.mean(r[(GRID > hi - 10) & (GRID <= hi)]))
    return float(GRID[i]), float(r[i]), f, end / float(r[i])


def fds_constants(ref_rate, tp_c):
    """A (1/s) and E (J/mol) that FDS computes from REFERENCE_TEMPERATURE / REFERENCE_RATE / HEATING_RATE."""
    tp = tp_c + 273.15
    e = np.e * ref_rate * R * tp ** 2 / BETA
    return e, np.e * ref_rate * np.exp(e / (R * tp))


def main():
    runs = load_runs()
    verdict = pd.read_csv(HERE / "results" / "qa_verdict.csv")
    overall = {(r.material, "N" if r.atmosphere == "N2" else "O"): r.overall for r in verdict.itertuples()}
    rows, curves = [], {"T_C": GRID}
    for m in ORDER:
        for atm in ("N", "O"):
            a_runs = [runs[(m, atm, k)]["a"] for k in (1, 2)]
            r_runs = [runs[(m, atm, k)]["r"] for k in (1, 2)]
            a_mean, r_mean = np.mean(a_runs, axis=0), np.mean(r_runs, axis=0)
            curves[f"{m}_{atm}_mass_norm"] = 1 - a_mean
            curves[f"{m}_{atm}_rate_per_s"] = r_mean * BETA
            cuts = bounds(r_mean, (m, atm))
            low_loss = float(np.interp(T_START, GRID, a_mean))
            for k, (lo, hi) in enumerate(zip(cuts[:-1], cuts[1:]), start=1):
                tp, rp, f, end_share = stage_numbers(a_mean, r_mean, lo, hi)
                per_run = [stage_numbers(a, r, lo, hi) for a, r in zip(a_runs, r_runs)]
                ref = BETA * rp / f
                ref_runs = [BETA * p[1] / p[2] for p in per_run]
                e_j, a_s = fds_constants(ref, tp)
                last = hi >= GRID[-1]
                rows.append(dict(
                    material=m, label=LABEL[m], atmosphere=ATM_KO[atm], stage=k, role=ROLE[(m, atm, k)],
                    T_from_C=lo, T_to_C=hi, loss_share=round(f, 6), finished=bool(end_share < COMPLETE or not last),
                    end_rate_share=round(end_share, 3) if last else None,
                    REFERENCE_TEMPERATURE_C=round(tp, 1), T_peak_runs_C=" / ".join(f"{p[0]:.0f}" for p in per_run),
                    peak_rate_pct_per_C=round(100 * rp, 6),
                    REFERENCE_RATE_per_s=float(f"{ref:.3g}"), REFERENCE_RATE_runs=" / ".join(f"{v:.3g}" for v in ref_runs),
                    HEATING_RATE_K_min=BETA_K_MIN, PYROLYSIS_RANGE_equiv_C=round(2 * f / rp, 0),
                    E_kJ_mol=round(e_j / 1000, 0), A_per_s=float(f"{a_s:.3g}"),
                    A_per_s_if_N_O2_1_air=float(f"{a_s / X_O2_AIR:.3g}") if atm == "O" else None,
                    loss_below_150C_share=round(low_loss, 3), tga_verdict=overall[(m, atm)]))
    st = pd.DataFrame(rows)
    out = HERE / "results"
    st.to_csv(out / "pyrosim_reactions.csv", index=False, encoding="utf-8")
    cv = pd.DataFrame(curves)
    note = ("# mass_norm = 1 - alpha, alpha = loss since the ramp start / loss up to 790 °C (mean of two runs); "
            "rate_per_s = d alpha / dt at 5 K/min. To compare with FDS TGA_ANALYSIS (CHID_tga.csv, mass / initial mass Y): "
            "alpha_FDS = (1 - Y) / (1 - Y at 790 °C)\n")
    with open(out / "pyrosim_tga_curves.csv", "w", encoding="utf-8") as fh:
        fh.write(note)
        cv.round(5).to_csv(fh, index=False)
    js = dict(beta_K_min=BETA_K_MIN, x_o2_air=X_O2_AIR, stages=json.loads(st.to_json(orient="records", force_ascii=False)))
    (HERE / "report" / "pyrosim_data.json").write_text(json.dumps(js, ensure_ascii=False, indent=1), encoding="utf-8")
    pd.set_option("display.width", 250)
    pd.set_option("display.max_columns", 30)
    print(st.drop(columns=["label", "role"]).to_string())


if __name__ == "__main__":
    main()
