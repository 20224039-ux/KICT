#!/usr/bin/env python3
"""외벽(중공층) 화재위험 평가식 – 재료 열물성(k, ρ, c, T_ig)·발열량으로 α와 위험등급을 예측하고 3 m 시험으로 보정.

Physics chain (thermally thick insulation lining the cavity, see RISK_MODEL.md):
    TRP   = ΔT_ig · sqrt(π k ρ c / 4)                    thermal response parameter   [kW s^0.5 / m²]
    t_ig  = (TRP / q″)²                                   ignition time under flux q″
    v_up  ∝ Q″^(2/3) / TRP²                               upward spread velocity (FPI-type scaling)
    M     = Q″^(1/3) / TRP                                material index (FPI-type)
    t600  = 1 / (K · M)²        α600 = 600 / t600² = 600 · K⁴ · M⁴      K: cavity-system constant (fitted)
    FRI   = t_ref / t600 ≈ floors climbed within t_ref (h_floor ≈ 3 m = test height)

    python risk_model.py                       # uses results/index_3000mm.csv + material_properties.csv

Without material data it still writes the back-calculated quantities from the 3 m tests
(FRI and the material-index ratio the model requires). Fill material_properties.csv to calibrate K.
"""
from __future__ import annotations

import argparse
import csv
from pathlib import Path

import numpy as np

HERE = Path(__file__).resolve().parent
THETA_REF = 600.0          # K
T_REF_S = 900.0            # s, reference time for FRI (BS 8414-style 15 min window)
FRI_GRADES = ((2.0, "IV"), (1.0, "III"), (0.5, "II"))   # FRI >= bound -> grade; else I
K_F = 0.01                 # m²/kW, flame-length coefficient of the linearised upward-spread model (verify / calibrate)
REFERENCE = "PIR_SC"       # material the required index ratios are expressed against


def fnum(s):
    try:
        return float(s)
    except (TypeError, ValueError):
        return np.nan


def read_csv(path):
    with open(path, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def grade_from_fri(fri):
    if not np.isfinite(fri):
        return "n/a"
    for bound, g in FRI_GRADES:
        if fri >= bound:
            return g
    return "I"


def material_terms(p):
    """TRP (nominal from k, ρ, c, ΔT_ig and/or from cone ignition times), index M and spread gate b."""
    k, rho, c = fnum(p.get("k_W_mK")), fnum(p.get("rho_kg_m3")), fnum(p.get("c_J_kgK"))
    dT = fnum(p.get("T_ig_C")) - fnum(p.get("T0_C") or 20)
    krc = (k / 1e3) * rho * (c / 1e3)                                   # kW² s / (m⁴ K²)
    trp_nom = dT * np.sqrt(np.pi * krc / 4) if np.isfinite(krc) and np.isfinite(dT) else np.nan
    q, tig = fnum(p.get("q_cone_kW_m2")), fnum(p.get("t_ig_cone_s"))
    qcr = fnum(p.get("q_cr_kW_m2"))
    qcr = 0.0 if not np.isfinite(qcr) else qcr
    trp_cone = (q - qcr) * np.sqrt(tig) if np.isfinite(q) and np.isfinite(tig) else np.nan
    trp_given = fnum(p.get("TRP_kW_s05_m2"))
    trp = next((x for x in (trp_given, trp_cone, trp_nom) if np.isfinite(x)), np.nan)
    src = "given" if np.isfinite(trp_given) else "cone" if np.isfinite(trp_cone) else "k-rho-c" if np.isfinite(trp_nom) else ""
    qpk, thr = fnum(p.get("pHRR_kW_m2")), fnum(p.get("THR_MJ_m2"))
    M = qpk ** (1 / 3) / trp if np.isfinite(qpk) and np.isfinite(trp) and trp > 0 else np.nan
    t_b = thr * 1e3 / qpk if np.isfinite(thr) and np.isfinite(qpk) and qpk > 0 else np.nan    # burning time ≈ THR / pHRR
    b = K_F * qpk - 1 - tig / t_b if all(np.isfinite([qpk, tig, t_b])) else np.nan
    return dict(kRhoC=krc, TRP_nominal=trp_nom, TRP_cone=trp_cone, TRP_used=trp, TRP_source=src, M=M,
                t_burn_s=t_b, b_gate=b)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--tests", type=Path, default=HERE / "results" / "index_3000mm.csv")
    ap.add_argument("--props", type=Path, default=HERE / "material_properties.csv")
    ap.add_argument("--out", type=Path, default=HERE / "results" / "risk_model.csv")
    ap.add_argument("--t-ref", type=float, default=T_REF_S)
    args = ap.parse_args()

    tests = [r for r in read_csv(args.tests) if r["in_index"] == "True"]
    props = {r["material"]: r for r in read_csv(args.props)} if args.props.exists() else {}
    terms = {m: material_terms(p) for m, p in props.items()}

    # back-calculation from the 3 m tests (no material data needed)
    t600 = {r["test"]: fnum(r["t600_s"]) for r in tests}
    mat = {r["test"]: r["material"] for r in tests}
    ref = [t600[k] for k in t600 if mat[k] == REFERENCE and np.isfinite(t600[k])]
    t_ref_mat = float(np.exp(np.mean(np.log(ref)))) if ref else np.nan

    # calibration of the system constant K (slope fixed to the theoretical value 1 on log-log)
    pairs = [(k, terms[mat[k]]["M"]) for k in t600
             if mat[k] in terms and np.isfinite(terms[mat[k]]["M"]) and np.isfinite(t600[k])]
    K = np.nan
    if pairs:
        y = np.array([np.log(1 / np.sqrt(t600[k])) for k, _ in pairs])
        x = np.array([np.log(m) for _, m in pairs])
        K = float(np.exp(np.mean(y - x)))
        if len({mat[k] for k, _ in pairs}) >= 3:
            slope = np.polyfit(x, y, 1)[0]
            print(f"free log-log slope (theory = 1): {slope:.2f}")
        print(f"K = {K:.4g} (from {len(pairs)} tests)")

    rows = []
    for r in tests:
        k = r["test"]
        t = t600[k]
        fri = args.t_ref / t if np.isfinite(t) else np.nan
        row = dict(test=k, material=mat[k], t600_s=t, FRI_measured=fri,
                   grade_measured=grade_from_fri(fri) if np.isfinite(t) else "I",
                   required_M_ratio_vs_ref=np.sqrt(t_ref_mat / t) if np.isfinite(t) else np.nan)
        tm = terms.get(mat[k], {})
        row.update({f"{n}": tm.get(n, np.nan) for n in ("kRhoC", "TRP_nominal", "TRP_cone", "TRP_used", "M", "b_gate")})
        row["TRP_source"] = tm.get("TRP_source", "")
        if np.isfinite(K) and np.isfinite(tm.get("M", np.nan)):
            tp = 1 / (K * tm["M"]) ** 2
            row.update(t600_pred_s=tp, alpha600_pred_K_s2=THETA_REF / tp ** 2, FRI_pred=args.t_ref / tp,
                       grade_pred=grade_from_fri(args.t_ref / tp))
        rows.append(row)
    # materials with properties but no cavity test: prediction only
    for m, tm in terms.items():
        if m in mat.values() or not (np.isfinite(K) and np.isfinite(tm["M"])):
            continue
        tp = 1 / (K * tm["M"]) ** 2
        rows.append(dict(test="(prediction)", material=m, TRP_used=tm["TRP_used"], TRP_source=tm["TRP_source"],
                         M=tm["M"], b_gate=tm["b_gate"], t600_pred_s=tp, alpha600_pred_K_s2=THETA_REF / tp ** 2,
                         FRI_pred=args.t_ref / tp, grade_pred=grade_from_fri(args.t_ref / tp)))

    keys = list(dict.fromkeys(k for r in rows for k in r))
    with open(args.out, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=keys)
        w.writeheader()
        for r in rows:
            w.writerow({k: ("" if (isinstance(v, float) and not np.isfinite(v)) else
                            round(v, 6) if isinstance(v, float) else v) for k, v in ((k, r.get(k, np.nan)) for k in keys)})
    for r in rows:
        print(f"{r['test']:12s} {r['material']:7s} t600={r.get('t600_s', np.nan):7.0f} FRI={r.get('FRI_measured', np.nan):5.2f} "
              f"grade={r.get('grade_measured', '-'):3s} M/M_ref(required)={r.get('required_M_ratio_vs_ref', np.nan):5.2f} "
              f"pred t600={r.get('t600_pred_s', np.nan):7.0f} grade_pred={r.get('grade_pred', '-')}")


if __name__ == "__main__":
    main()
