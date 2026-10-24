#!/usr/bin/env python3
"""기축 건축물 외벽 화재위험성 평가 방법(안) — 참조 구현 (existing-building façade fire risk, Korea).

A building of any form is cut into façade zones z: a vertical run of one wall system with one exposure, from storey b
to t, in a vertical column col (zones stacked in one column share the storeys above). Three ways a fire meets a
zone are counted, each as frequency × probability of reaching each storey above × people exposed there:

  IE1  fire in a unit/room behind the zone at storey i, flashover, flames out of the window
       R1 = Σ_i f_i Σ_j c(i+j) [ p_e a (b q_min + (1 − b) q_a) + (1 − p_e a) b q_b ]
  IE2  fire at the foot of the zone (piloti car, refuse, equipment):  R2 = Σ_s F_s Σ_j c(o+j) p_e a_s(j) q_s(j)
  IE3  flames out of the top storey of the zone below in the same column, hitting this zone's lowest edge
       R3 = f_top Σ_j c(o+j) p_e a(j) [ b q_min + (1 − b) q_a − b q_b ]        (leapfrog part already in R1 below)

  a(j)  façade path reaches the j-th storey above: (1 − p_cb)^j · S_FS(jτ) · [p_B + (1 − p_B) S_exp(jτ)],
        τ = s · t600 · h / 3 m — the per-storey time of the 3 m cavity test, scaled to the building
  b(j)  window-to-window leapfrog: (p_L (1 − p_s u_s))^j · S_FS(j τ_L)      (stays with a non-combustible façade)
  q(t)  people on that storey not yet out when the fire arrives: S_RSET(t); stair smoke from a piloti fire adds p_st
  c(k)  people exposed on storey k of the column (spread band)

R [persons exposed / year], reference R_ref = the same building with every façade non-combustible (leapfrog only).
Grade from the individual exposure risk IR = R / occupants, as a ratio to the home-fire risk λ_res:
r = IR / λ_res; 1 < 0.01 ≤ 2 < 0.1 ≤ 3 < 1 ≤ 4 < 10 ≤ 5; flag M when one façade exposure exposes ≥ 10 people.

Material inputs (t600, p_B, A) come from the evidence level: E0 desk (type + permit-date rules → conservative default),
E1 site survey, E2 samples (FT-IR identifies, cone gates grade, TGA predicts t600 and p_B for charring
semi-noncombustible products), E3 system test / simulation.

Writes results/*.csv, figures/*.png, report/ebfra_data.json
    python ebfra.py
"""
from __future__ import annotations

import glob
import json
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib import font_manager
from scipy.stats import norm

HERE = Path(__file__).resolve().parent
SEED = 551
N_MC = 4000
GRADE_KO = {1: "낮음", 2: "보통", 3: "높음", 4: "매우 높음", 5: "심각"}
OPTIONS = [  # remediation options (applied where they make sense)
    ("현재", ()),
    ("개구부 차단재", ("closer",)),
    ("화재확산방지구조", ("barrier",)),
    ("하단 보호·필로티 개선", ("base",)),
    ("차단재 + 확산방지구조 + 하단 보호", ("closer", "barrier", "base")),
    ("스프링클러 전층 설치", ("spr",)),
    ("외장재 불연 교체", ("replace",)),
]
# evidence level → spread of ln t600 around the value used (multiscale_fra): between semi-noncombustible products,
# TGA prediction error, two repeat tests
EV_SIGMA = {"E0": 0.506, "E1": 0.506, "E2": 0.166, "E3": 0.0927 / np.sqrt(2)}


# ------------------------------------------------------------------------------------------------ inputs
def load():
    par = pd.read_csv(HERE / "inputs" / "parameters.csv")
    types = pd.read_csv(HERE / "inputs" / "facade_types.csv").set_index("key")
    hist = pd.read_csv(HERE / "inputs" / "code_history.csv", dtype=str, keep_default_na=False)
    blds = json.loads((HERE / "inputs" / "buildings.json").read_text(encoding="utf-8"))["buildings"]
    P = {r.key: float(r.best) for r in par.itertuples()}
    return par, P, types, hist, blds


def S_ln(t, m, sig):
    """P(X > t) for a lognormal X with median m"""
    return norm.sf((np.log(t) - np.log(m)) / sig)


def S_uni(t, lo, hi):
    return np.clip((hi - t) / (hi - lo), 0.0, 1.0)


def material(z, P, T):
    """t600 (s), p_B, A, façade path, source of the values — by evidence level"""
    row = T.loc[z["type"]]
    if not int(row.path):
        return dict(t600=None, pB=0.0, A=0, path=False, cb_ok=False, pe_ok=False, src=f"{row.label}: 외벽 경로 없음")
    t600, pB, A, src = float(row.t600_s), float(row.pB), int(row.A), f"{row.label} 기본값({row.evidence})"
    # fire stops and opening closers get credit only where the combustible layer is enclosed (cavity, ETICS),
    # not where it is the outer face or a continuous panel core (PE-core ACP, sandwich panel, unknown)
    ok = dict(cb_ok=bool(int(row.cb_ok)), pe_ok=bool(int(row.pe_ok)))
    tga = z.get("tga")
    if tga:
        if tga["cone"] == "c":
            c = T.loc["C_UNKNOWN"]
            t600, pB, A, src = float(c.t600_s), float(c.pB), 1, "콘: 준불연 기준 미달 → 가연 경계값(E2)"
        elif tga["f400_air"] < P["f_melt"]:
            A, src = 1, "TGA: 용융형 → 확산식 미적용, 급발 A(E2)"
        else:
            t600 = P["tga_c0"] + P["tga_c1"] / tga["r_max_air"]
            pB = float(np.clip((tga["f400_air"] - P["f_lo"]) / (P["f_hi"] - P["f_lo"]), 0.05, 0.95))
            A, src = 0, "TGA 예측식(E2): t̂₆₀₀ = 247.5 + 422.97 / r_max,air"
    if "t600" in z:
        t600, src = z["t600"], "시스템 시험값(E3)"
    if "pB" in z:
        pB = z["pB"]
    return dict(t600=t600, pB=pB, A=A, path=True, src=src, **ok)


# ------------------------------------------------------------------------------------------------ model
def assess(b, P, T, opt=(), over=None):
    """risk of one building; P values may be arrays (Monte Carlo) — results broadcast to their shape"""
    opt = set(opt)
    N, h = b["N"], b["h"]
    sf = b["sprinkler"]                       # true = every storey, {"from": k} = storey k and above, false = none
    spr_from = 1 if (sf is True or "spr" in opt) else (int(sf["from"]) if isinstance(sf, dict) else 0)
    gs = lambda k: 1.0 - P["p_s"] * (1.0 if spr_from and k >= spr_from else 0.0)
    n_spr = lambda i, j: max(0, i + j - max(i + 1, spr_from) + 1) if spr_from else 0   # sprinklered storeys in (i, i+j]
    pil = dict(b["piloti"]) if b.get("piloti") else None
    if pil and "base" in opt:                 # piloti ceiling non-combustible + ground storey compartmented
        pil = dict(soffit_combustible=False, ground_compartment=True)
    H = b.get("H_reach", P["h_reach"])
    sig, sigr = P["sig_fs"], P["sig_r"]

    zones = []
    for z0 in b["zones"]:
        z = {**z0, **((over or {}).get(z0["id"], {}))}
        m = material(z, P, T)
        if "replace" in opt and m["path"]:
            z = {k: v for k, v in z.items() if k not in ("tga", "t600", "pB")}
            z["type"] = "NC"
            m = material(z, P, T)
        z["m"] = m
        z["pe"] = P["p_e1"] if ((z.get("closer") or "closer" in opt) and m["pe_ok"]) else P["p_e0"]      # window plumes
        z["pe2"] = P["p_e1"] if ((z.get("base_closed") or "base" in opt) and m["pe_ok"]) else P["p_e0"]  # fire at the foot
        z["pcb"] = P["p_cb"] if ((z.get("barrier") or "barrier" in opt) and m["cb_ok"]) else 0.0
        z["pL"] = np.minimum(0.9, P["p_L"] * (P["f_gap"] if z.get("gap") else 1.0))
        z["tau"] = P["s"] * m["t600"] / 60.0 * h / 3.0 if m["path"] else None
        base = z.get("n_u", 0) * P["lam_res"] if z["use"] == "res" else z.get("area", 0) * P["lam_com"]
        z["fb"] = base * P["p_fo"] * P["p_v"] * z.get("win", 1.0)          # per origin storey, before sprinklers
        z["ie1"] = z.get("win", 1.0) > 0 and (z.get("n_u", 0) > 0 or z.get("area", 0) > 0)
        zones.append(z)

    def band(col, k):
        for w in zones:
            if w["col"] == col and w["b"] <= k <= w["t"]:
                return float(w["occ_band"]), w["use"]
        return 0.0, None

    mfs = lambda k: P["t_fs_lo"] if (k - 1) * h <= H else P["t_fs_hi"]
    q = lambda t, use: S_ln(t, P["rset_res"] if use == "res" else P["rset_com"], sigr)

    leap = lambda z, i, j: z["pL"] ** j * (1 - P["p_s"]) ** n_spr(i, j)    # sprinklered storeys stop the leapfrog

    def a_path(z, j, k, lo, hi):
        ta = j * z["tau"]
        return ta, (1 - z["pcb"]) ** j * S_ln(ta, mfs(k), sig) * (z["m"]["pB"] + (1 - z["m"]["pB"]) * S_uni(ta, lo, hi))

    out_z, F_ign, F2 = [], 0.0, 0.0
    for z in zones:
        col, m = z["col"], z["m"]
        r1 = r1ref = r2 = r3 = 0.0
        # IE1 — fires behind this zone
        if z["ie1"]:
            for i in range(z["b"], z["t"] + 1):
                fi = z["fb"] * gs(i)
                F_ign = F_ign + fi
                for j in range(1, N - i + 1):
                    k = i + j
                    c, use = band(col, k)
                    if c == 0:
                        continue
                    tl = j * P["tau_L"]
                    bj = leap(z, i, j) * S_ln(tl, mfs(k), sig)
                    qb = q(tl, use)
                    ref = c * bj * qb
                    if m["path"] and k <= z["t"]:
                        ta, aj = a_path(z, j, k, P["te1_lo"], P["te1_hi"])
                        pa = z["pe"] * aj
                        val = c * (pa * (bj * q(np.minimum(ta, tl), use) + (1 - bj) * q(ta, use)) + (1 - pa) * bj * qb)
                        p2 = pa + (1 - pa) * bj
                    else:
                        val, p2 = ref, bj
                    r1 = r1 + fi * val
                    r1ref = r1ref + fi * ref
                    if j == 2:
                        F2 = F2 + fi * p2
        # IE2 — fires at the foot of this zone
        o = z["b"] - 1
        for e in z.get("ext", []):
            if e["src"] == "piloti" and not pil:
                continue
            if e["src"] == "piloti":
                lam, px = P["lam_car"] * e["count"], (P["p_x_pc"] if pil["soffit_combustible"] else P["p_x_pn"])
                lo, hi = (P["tep_lo"], P["tep_hi"]) if pil["soffit_combustible"] else (P["te2_lo"], P["te2_hi"])
                stair = not pil["ground_compartment"]
            else:
                lam, px, lo, hi, stair = P["lam_adj"] * e["count"], P["p_x_adj"], P["te2_lo"], P["te2_hi"], False
            Fs = lam * px
            F_ign = F_ign + Fs
            if not m["path"]:
                continue
            for j in range(1, z["t"] - o + 1):
                k = o + j
                c, use = band(col, k)
                if c == 0:
                    continue
                ta, aj = a_path(z, j, k, lo, hi)
                qa = q(ta, use)
                if stair:
                    qa = qa + (1 - qa) * P["p_stair"]
                r2 = r2 + Fs * c * z["pe2"] * aj * qa
                if j == 2:
                    F2 = F2 + Fs * z["pe2"] * aj
        # IE3 — flames from the top storey of the zone below (same column) hit this zone's lowest edge
        y = next((w for w in zones if w["col"] == col and w["t"] == z["b"] - 1 and w["ie1"]), None)
        if y is not None and m["path"]:
            fy = y["fb"] * gs(y["t"])
            for j in range(1, z["t"] - o + 1):
                k = o + j
                c, use = band(col, k)
                if c == 0:
                    continue
                tl = j * P["tau_L"]
                bj = leap(y, y["t"], j) * S_ln(tl, mfs(k), sig)
                ta, aj = a_path(z, j, k, P["te1_lo"], P["te1_hi"])
                pa = z["pe2"] * aj
                r3 = r3 + fy * c * pa * (bj * q(np.minimum(ta, tl), use) + (1 - bj) * q(ta, use) - bj * q(tl, use))
                if j == 2:
                    F2 = F2 + fy * pa * (1 - bj)
        occ = (z.get("n_u", 0) * P["occ_unit"] if z["use"] == "res" else z.get("area", 0) * P["dens_com"]) * (z["t"] - z["b"] + 1)
        out_z.append(dict(zone=z["id"], name=z["name"], type=z["type"], material_src=m["src"], t600=m["t600"], pB=m["pB"],
                          A=m["A"], tau_min=z["tau"], r1=r1, r1ref=r1ref, r2=r2, r3=r3, occ=occ))

    R = sum(o["r1"] + o["r2"] + o["r3"] for o in out_z)
    Rref = sum(o["r1ref"] for o in out_z)
    Nocc = sum(o["occ"] for o in out_z)
    IR = R / Nocc
    r = IR / P["lam_res"]
    grade = 1 + sum((r >= P[g]).astype(int) if isinstance(r, np.ndarray) else int(r >= P[g]) for g in ("g1", "g2", "g3", "g4"))
    r_ref = Rref / Nocc / P["lam_res"]
    grade_ref = 1 + sum((r_ref >= P[g]).astype(int) if isinstance(r_ref, np.ndarray) else int(r_ref >= P[g])
                        for g in ("g1", "g2", "g3", "g4"))
    Cbar = R / F_ign
    return dict(R=R, Rref=Rref, Nocc=Nocc, IR=IR, r=r, grade=grade, Cbar=Cbar, M=Cbar >= P["m_flag"], RR=R / Rref,
                AF=1 - Rref / R, dR=R - Rref, F_ign=F_ign, F2=F2, r_ref=r_ref, grade_ref=grade_ref,
                A=max(o["A"] for o in out_z), zones=out_z)


def label(res):
    g = int(res["grade"])
    flags = ("·M" if res["M"] else "") + ("·A" if res["A"] else "")
    return f"{g}등급({GRADE_KO[g]}){flags}"


# ------------------------------------------------------------------------------------------------ analyses
def options_for(b):
    out = []
    for name, opt in OPTIONS:
        if "spr" in opt and b["sprinkler"] is True:
            continue
        if "base" in opt and not any(z.get("ext") for z in b["zones"]):
            continue
        out.append((name, opt))
    return out


def sample(par, n, rng):
    """epistemic Monte Carlo: every parameter with a range, others at best"""
    S = {}
    for r in par.itertuples():
        if r.distribution == "logu":
            S[r.key] = np.exp(rng.uniform(np.log(r.low), np.log(r.high), n))
        elif r.distribution == "tri":
            S[r.key] = rng.triangular(r.low, r.best, r.high, n)
        else:
            S[r.key] = float(r.best)          # fixed values stay scalars
    return S


def flat(res, **extra):
    keys = ("R", "Rref", "Nocc", "IR", "r", "grade", "Cbar", "M", "RR", "AF", "dR", "F_ign", "F2", "r_ref", "grade_ref", "A")
    d = {k: (float(res[k]) if not isinstance(res[k], (bool, np.bool_)) else bool(res[k])) for k in keys}
    d["grade"] = int(res["grade"])
    d["grade_ref"] = int(res["grade_ref"])
    d["label"] = label(res)
    return {**extra, **d}


def reach_curves(P, T, keys=("ETICS_EPS_C", "CAV_PF_IN", "CAV_PIR_SC", "CAV_PF_OUT", "NC"), h=3.0, jmax=10):
    """probability that a window plume (IE1) reaches the j-th storey above, by wall type — reference storey within
    ladder reach, no sprinklers, no barrier or closer"""
    rows = []
    for key in keys:
        m = material({"type": key}, P, T)
        for j in range(1, jmax + 1):
            tl = j * P["tau_L"]
            bj = P["p_L"] ** j * S_ln(tl, P["t_fs_lo"], P["sig_fs"])
            if m["path"]:
                ta = j * P["s"] * m["t600"] / 60 * h / 3
                aj = S_ln(ta, P["t_fs_lo"], P["sig_fs"]) * (m["pB"] + (1 - m["pB"]) * S_uni(ta, P["te1_lo"], P["te1_hi"]))
                pa = P["p_e0"] * aj
                pj = pa + (1 - pa) * bj
            else:
                pj = bj
            rows.append(dict(type=key, label=T.loc[key, "label"], j=j, p=float(pj)))
    return rows


def evidence_mc(b, S, P, T, rng):
    """grade spread at each evidence level: t600 drawn around the value used (EV_SIGMA), p_B from its beta law
    (E0/E1: the semi-noncombustible population, E2: TGA prior with weight 2 tests, E3: two tests)"""
    sc = T[(T.grade == "준불연") & (T.evidence == "E3")].t600_s.astype(float)
    pop = float(np.exp(np.log(sc).mean()))
    n = len(S["s"])
    Z = rng.standard_normal(n)
    rows = []
    for lev, spec in b.get("evidence_levels", {}).items():
        over = {}
        for z0 in b["zones"]:
            z = {**z0, **spec["zones"].get(z0["id"], {})}
            m = material(z, P, T)
            if not m["path"]:
                continue
            if lev in ("E0", "E1"):
                t600, pB = pop * np.exp(EV_SIGMA[lev] * Z), rng.beta(0.5, 0.5, n)
            else:
                pb = float(m["pB"])
                t600, pB = m["t600"] * np.exp(EV_SIGMA[lev] * Z), rng.beta(3 * pb, 3 * (1 - pb), n) if lev == "E3" \
                    else rng.beta(2 * pb, 2 * (1 - pb), n)
            over[z0["id"]] = {**spec["zones"].get(z0["id"], {}), "t600": t600, "pB": pB}
        res = assess(b, S, T, over=over)
        g = res["grade"]
        rows.append(dict(building=b["id"], level=lev, sigma_ln_t600=EV_SIGMA[lev], t600_median=float(np.median(over[b["zones"][0]["id"]]["t600"])),
                         r_p05=float(np.percentile(res["r"], 5)), r_p50=float(np.percentile(res["r"], 50)),
                         r_p95=float(np.percentile(res["r"], 95)), **{f"P_g{k}": float(np.mean(g == k)) for k in range(1, 6)}))
    return rows


def korean_font():
    for pattern in ("C:/Windows/Fonts/malgun*.ttf", "/usr/share/fonts/truetype/nanum/NanumGothic*.ttf",
                    "/Library/Fonts/AppleGothic.ttf", "/System/Library/Fonts/Supplemental/AppleGothic.ttf",
                    "/usr/share/fonts/opentype/noto/NotoSansCJK*.ttc"):
        for f in glob.glob(pattern):
            try:
                font_manager.fontManager.addfont(f)
            except (OSError, RuntimeError):
                pass
    names = {f.name for f in font_manager.fontManager.ttflist}
    for name in ("Malgun Gothic", "NanumGothic", "AppleGothic", "Noto Sans CJK KR", "Noto Sans KR"):
        if name in names:
            return name
    raise SystemExit("no Hangul font found (Malgun Gothic, NanumGothic, AppleGothic or Noto Sans CJK KR)")


def main():
    par, P, T, hist, blds = load()
    rng = np.random.default_rng(SEED)
    rows, zrows, rem, ev = [], [], [], []
    for b in blds:
        base = assess(b, P, T)
        rows.append(flat(base, building=b["id"], name=b["name"], form=b["form"]))
        for o in base["zones"]:
            zrows.append({"building": b["id"], **{k: (float(v) if isinstance(v, (np.floating, float)) else v) for k, v in o.items()}})
        for name, opt in options_for(b):
            rem.append(flat(assess(b, P, T, opt), building=b["id"], option=name))
        for lev, spec in b.get("evidence_levels", {}).items():
            res = assess(b, P, T, over=spec["zones"])
            z1 = res["zones"][0]
            ev.append(flat(res, building=b["id"], level=lev, note=spec["note"], t600=z1["t600"], pB=z1["pB"], src=z1["material_src"]))
    summ, zone_df, rem_df, ev_df = pd.DataFrame(rows), pd.DataFrame(zrows), pd.DataFrame(rem), pd.DataFrame(ev)

    # epistemic Monte Carlo (same draws for every building and option)
    S = sample(par, N_MC, rng)
    mc, mc_rows = {}, []
    for b in blds:
        res = assess(b, S, T)
        repl = assess(b, S, T, ("replace",))
        mc[b["id"]] = res
        g = res["grade"]
        pct = lambda v, p: float(np.percentile(v, p))
        mc_rows.append(dict(building=b["id"], IR_p05=pct(res["IR"], 5), IR_p50=pct(res["IR"], 50), IR_p95=pct(res["IR"], 95),
                            r_p05=pct(res["r"], 5), r_p50=pct(res["r"], 50), r_p95=pct(res["r"], 95),
                            RR_p05=pct(res["RR"], 5), RR_p50=pct(res["RR"], 50), RR_p95=pct(res["RR"], 95),
                            R_ratio90=pct(res["R"], 95) / pct(res["R"], 5), r_ratio90=pct(res["r"], 95) / pct(res["r"], 5),
                            RR_ratio90=pct(res["RR"], 95) / pct(res["RR"], 5),
                            **{f"P_g{k}": float(np.mean(g == k)) for k in range(1, 6)},
                            P_M=float(np.mean(res["M"])), P_lower_after_replace=float(np.mean(repl["grade"] < g))))
    mc_df = pd.DataFrame(mc_rows)
    evmc_df = pd.DataFrame([r for b in blds for r in evidence_mc(b, S, P, T, np.random.default_rng(SEED + 1))])

    # one-at-a-time sensitivity of r (= IR / λ_res) for each building
    torn = []
    for b in blds:
        r0 = float(assess(b, P, T)["r"])
        for p in par.itertuples():
            if p.distribution == "fixed":
                continue
            lo_, hi_ = (assess(b, {**P, p.key: v}, T)["r"] for v in (p.low, p.high))
            torn.append(dict(building=b["id"], key=p.key, symbol=p.symbol, item=p.item, low=p.low, high=p.high,
                             r_low=float(lo_), r_high=float(hi_), swing=float(abs(np.log10(hi_) - np.log10(lo_))), r_base=r0))
    torn_df = pd.DataFrame(torn).sort_values(["building", "swing"], ascending=[True, False])

    # value of the simulation step: grade against the scale factor s
    sg = np.round(np.exp(np.linspace(np.log(0.25), np.log(1.0), 9)), 3)
    s_rows = []
    for b in blds:
        for sv in sg:
            res = assess(b, {**P, "s": float(sv)}, T)
            s_rows.append(dict(building=b["id"], s=float(sv), r=float(res["r"]), grade=int(res["grade"])))
    s_df = pd.DataFrame(s_rows)

    out = HERE / "results"
    for df, name in ((summ, "example_summary"), (zone_df, "zone_contributions"), (rem_df, "remediation"),
                     (ev_df, "evidence_levels"), (evmc_df, "evidence_uncertainty"), (mc_df, "uncertainty"), (torn_df, "sensitivity"), (s_df, "scale_factor")):
        df.to_csv(out / f"{name}.csv", index=False, encoding="utf-8")
    figures(blds, summ, rem_df, mc, P, torn_df)

    data = dict(params=json.loads(par.to_json(orient="records", force_ascii=False, double_precision=15)),
                types=json.loads(T.reset_index().to_json(orient="records", force_ascii=False, double_precision=15)),
                history=json.loads(hist.to_json(orient="records", force_ascii=False, double_precision=15)), buildings=blds,
                summary=json.loads(summ.to_json(orient="records", force_ascii=False, double_precision=15)),
                zones=json.loads(zone_df.to_json(orient="records", force_ascii=False, double_precision=15)),
                remediation=json.loads(rem_df.to_json(orient="records", force_ascii=False, double_precision=15)),
                evidence=json.loads(ev_df.to_json(orient="records", force_ascii=False, double_precision=15)),
                evidence_mc=json.loads(evmc_df.to_json(orient="records", force_ascii=False, double_precision=15)),
                uncertainty=json.loads(mc_df.to_json(orient="records", force_ascii=False, double_precision=15)),
                sensitivity=json.loads(torn_df.to_json(orient="records", force_ascii=False, double_precision=15)),
                scale=json.loads(s_df.to_json(orient="records", force_ascii=False, double_precision=15)),
                reach=reach_curves(P, T),
                n_mc=N_MC, seed=SEED, grade_ko=GRADE_KO)
    (HERE / "report" / "ebfra_data.json").write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")

    pd.set_option("display.width", 220)
    print(summ[["building", "name", "R", "Rref", "IR", "r", "label", "Cbar", "RR", "AF"]].to_string(index=False))
    print(rem_df[["building", "option", "r", "label", "dR"]].to_string(index=False))
    print(ev_df[["building", "level", "t600", "pB", "r", "label"]].to_string(index=False))
    print(mc_df.round(3).to_string(index=False))
    print(evmc_df.round(3).to_string(index=False))
    print(zone_df[["building", "zone", "r1", "r1ref", "r2", "r3", "occ", "tau_min"]].to_string(index=False))


def figures(blds, summ, rem, mc, P, torn):
    plt.rcParams.update({"font.family": korean_font(), "font.size": 10, "axes.unicode_minus": False,
                         "mathtext.fontset": "dejavusans", "axes.spines.top": False, "axes.spines.right": False})
    from matplotlib.ticker import FixedLocator, FuncFormatter, NullLocator
    pow10 = FuncFormatter(lambda v, _: f"$10^{{{int(round(np.log10(v)))}}}$" if v > 0 else "")
    cols = {"A": "#E34948", "B": "#2A78D6", "C": "#6250D6"}
    bounds = [P["g1"], P["g2"], P["g3"], P["g4"]]
    shades = ["#E2F3E2", "#F1F5E1", "#FEF2D3", "#FCE6DC", "#F8DEDE"]

    # 1 grade map: r for each remediation option, grade bands
    fig, axes = plt.subplots(1, len(blds), figsize=(12, 4.6), sharey=True)
    for ax, b in zip(axes, blds):
        d = rem[rem.building == b["id"]].reset_index(drop=True)
        edges = [1e-5] + bounds + [1e3]
        for g in range(5):
            ax.axhspan(edges[g], edges[g + 1], color=shades[g], lw=0, zorder=0)
        ax.scatter(range(len(d)), d.r, color=cols[b["id"]], s=36, zorder=3)
        ax.set_yscale("log")
        ax.set_ylim(1e-5, 1e3)
        ax.yaxis.set_major_locator(FixedLocator([10.0 ** e for e in range(-5, 4)]))
        ax.yaxis.set_minor_locator(NullLocator())
        ax.yaxis.set_major_formatter(pow10)
        ax.set_xticks(range(len(d)))
        ax.set_xticklabels(d.option, rotation=60, ha="right", fontsize=8.5)
        ax.set_title(f"{b['id']}  {b['name']}", fontsize=10.5, loc="left")
        for g in range(5):
            ax.text(len(d) - 0.4, np.sqrt(edges[g] * edges[g + 1]), f"{g + 1}", fontsize=9, color="#6B7280", va="center")
    axes[0].set_ylabel("외벽 노출 위험비 r = IR / λ_res (로그)")
    fig.tight_layout()
    fig.savefig(HERE / "figures" / "fig1_grade_by_option.png", dpi=170)
    plt.close(fig)

    # 2 grade probabilities from the epistemic Monte Carlo
    fig, ax = plt.subplots(figsize=(6.2, 3.6))
    gcol = ["#86C38B", "#B9D98C", "#F2C14E", "#EB8A3A", "#D64545"]
    for i, b in enumerate(blds):
        g = mc[b["id"]]["grade"]
        left = 0.0
        for k in range(1, 6):
            p = float(np.mean(g == k))
            ax.barh(i, p, left=left, color=gcol[k - 1], edgecolor="white", height=0.6)
            if p >= 0.06:
                ax.text(left + p / 2, i, f"{k}등급\n{p:.0%}", ha="center", va="center", fontsize=8.5)
            left += p
    ax.set_yticks(range(len(blds)))
    ax.set_yticklabels([f"{b['id']} {b['name']}" for b in blds])
    ax.invert_yaxis()
    ax.set_xlim(0, 1)
    ax.set_xlabel("등급 확률(인식론적 불확실성, 몬테카를로)")
    fig.tight_layout()
    fig.savefig(HERE / "figures" / "fig2_grade_uncertainty.png", dpi=170)
    plt.close(fig)

    # 3 tornado for each building (top 8)
    fig, axes = plt.subplots(1, len(blds), figsize=(12, 4.2))
    for ax, b in zip(axes, blds):
        d = torn[torn.building == b["id"]].head(8).iloc[::-1]
        base = d.r_base.iloc[0]
        for i, rr in enumerate(d.itertuples()):
            lo_, hi_ = sorted([rr.r_low, rr.r_high])
            ax.barh(i, np.log10(hi_ / base), left=0, color="#EB6834", height=0.6)
            ax.barh(i, np.log10(lo_ / base), left=0, color="#2A78D6", height=0.6)
        ax.axvline(0, color="#4B5563", lw=0.8)
        ax.set_yticks(range(len(d)))
        ax.set_yticklabels([s if isinstance(s, str) else "" for s in d.symbol], fontsize=9)
        ax.set_title(f"{b['id']} {b['name']}", fontsize=10.5, loc="left")
        ax.set_xlabel("$\\log_{10}$(r / 기준값)")
    fig.tight_layout()
    fig.savefig(HERE / "figures" / "fig3_sensitivity.png", dpi=170)
    plt.close(fig)


if __name__ == "__main__":
    main()
