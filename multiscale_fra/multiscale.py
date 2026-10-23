#!/usr/bin/env python3
"""Multi-scale façade fire risk assessment (MS-FRA): specimen tests, the mid-scale cavity test and simulation, each
stating the material in the same two numbers, combined into one building risk.

The building model (facade_fra/qfra.py, NFPA 551 5.5: R = sum F C) takes two material numbers: the time the cavity fire
needs to climb one storey (t600 at 3000 mm in the 3 m cavity test, scaled by s) and the probability p_B that the cavity
keeps burning once the exposing fire is out (flag B). Every scale is read as an estimate of these two numbers, the
"mid-scale equivalent", so each scale gives a building risk of its own, and the scales can be combined and checked
against each other.

  L1 specimen    cone calorimeter, TGA-DSC, FT-IR. Only TGA-DSC exists so far:
                   t600 = t_ig,sys + C_sys / r_max,air   (semi-noncombustible charring products; constants refitted
                                                          without the product, integrated_model leave-one-out)
                   p_B ~ Beta(n0 m, n0 (1 - m)),  m = ramp of f400,air from 0.60 to 0.70 (B-hat boundary 0.65 in the middle)
                 gates: grade c -> the TGA relations do not apply (PIR(PUR) c: 812 s predicted, 30 s measured);
                        f400,air < 0.5 (melting, EPS) -> the spread relation does not apply
  L2 mid-scale   t600 of each repeat test and flag B (Jeffreys): exactly facade_fra/qfra.py, same random numbers
  L3 simulation  not run: it would replace the assumed scale factor s; reported as the value of knowing s
  L1+L2          TGA prior updated by the tests: normal-normal on ln t600 (between-test SD pooled from the repeats),
                 beta-binomial on p_B

Evidence modes: L0 grade only (certificate), L1 grade + TGA, L1u TGA without the gates, L2 tests only, L1+L2.

Writes results/*.csv, figures/*.png, report/ms_data.json
    python multiscale.py
"""
from __future__ import annotations

import importlib.util
import json
import os
import pickle
from pathlib import Path
from statistics import NormalDist

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.ticker import FixedLocator, FuncFormatter, NullFormatter, NullLocator
from scipy.special import betaln, gammaln
from scipy.stats import beta as beta_dist

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
_spec = importlib.util.spec_from_file_location("qfra", ROOT / "facade_fra" / "qfra.py")
qf = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(qf)

MATS = qf.MATERIALS                     # products with cavity tests
CHAR = ["PIR_SC", "PF_out", "PF_in"]    # semi-noncombustible charring products: the TGA spread relation was fitted on them
ALL = ["EPS_C"] + MATS                  # EPS_C: TGA only (cavity records 4 min, no t600)
LABEL = qf.LABEL
F_RAMP = (0.60, 0.70)                   # f400,air ramp for the p_B prior
N0 = 2.0                                # weight of the TGA p_B prior, in tests
F_MELT = 0.5                            # f400,air below this: melting type (EPS)
TAU_MIN = 0.05
MODES = ["L0", "L1", "L1u", "L2", "L1+L2"]
MODE_KO = {"L0": "등급만(인증서)", "L1": "소형시료(등급+TGA)", "L1u": "TGA 단독(게이트 없음)",
           "L2": "중규모 시험", "L1+L2": "통합(TGA+중규모)"}
CASES = [(m, "B1", g) for g in ("M0", "M1") for m in ["NC"] + MATS]
S_GRID = [0.25, 0.30, 0.35, 0.42, 0.50, 0.59, 0.71, 0.84, 1.00]
VOI_OUTER, VOI_INNER = 1000, 200
COLOR = qf.COLOR


_CACHE = {}


def cached(key, fn):
    """Monte Carlo results kept in the pickle named by MS_CACHE (if set); delete the file after changing any input"""
    path = os.environ.get("MS_CACHE")
    if path and not _CACHE and os.path.exists(path):
        with open(path, "rb") as fh:
            _CACHE.update(pickle.load(fh))
    if key not in _CACHE:
        _CACHE[key] = fn()
        if path:
            with open(path, "wb") as fh:
                pickle.dump(_CACHE, fh)
    return _CACHE[key]


def z_band(p):
    return NormalDist().inv_cdf(p)


# ------------------------------------------------------------------------------------------------ data
def load():
    tga = pd.read_csv(ROOT / "tga_dsc" / "results" / "tga_summary.csv")
    air = tga[tga.atmosphere.str.startswith("oxidative")].set_index("material")
    n2 = tga[tga.atmosphere == "N2"].set_index("material")
    ver = pd.read_csv(ROOT / "tga_dsc" / "results" / "qa_verdict.csv")
    v_air = ver[ver.atmosphere.str.startswith("oxidative")].set_index("material").overall
    const = pd.read_csv(ROOT / "integrated_model" / "results" / "system_constants.csv").iloc[0]
    loo = pd.read_csv(ROOT / "integrated_model" / "results" / "loo_validation.csv").groupby("material").first()
    tests = pd.read_csv(ROOT / "facade_fra" / "inputs" / "test_inputs.csv")
    idx = pd.read_csv(ROOT / "cavity_fire_index" / "results" / "index_3000mm.csv").set_index("test")
    rm = pd.read_csv(ROOT / "cavity_fire_index" / "results" / "risk_model.csv").set_index("test")
    ig = pd.read_csv(ROOT / "coupled_analysis" / "results" / "ignition_spread_inverse.csv").set_index("test")
    return dict(air=air, n2=n2, v_air=v_air, t_ig_sys=float(const.t_ig_sys_s), c_sys=float(const.C_sys),
                c_sys_cv=float(const.C_sys_cv), loo=loo.t600_hat_loo_s.to_dict(), tests=tests, idx=idx, rm=rm, ig=ig)


def observed(ctx, m):
    t = ctx["tests"]
    obs = t[(t.material == m) & t.used]
    return obs, [None if pd.isna(v) else float(v) for v in obs.t600_s], int(obs.flag_B.sum()), len(obs)


def pb_prior(f400):
    m = float(np.clip((f400 - F_RAMP[0]) / (F_RAMP[1] - F_RAMP[0]), 0.05, 0.95))
    return N0 * m, N0 * (1.0 - m)


def tga_t600(ctx, m, loo=True):
    """TGA spread relation t600 = t_ig,sys + C_sys / r_max,air; leave-one-out constants for the products it was fitted on"""
    if loo and m in ctx["loo"]:
        return float(ctx["loo"][m])
    return ctx["t_ig_sys"] + ctx["c_sys"] / float(ctx["air"].loc[m, "r_max_pct_C_mean"])


def fit_stats(ctx):
    """between-test SD of ln t600 (pooled repeats), SD of the TGA prediction of a product mean (leave-one-out),
    between-product SD of the semi-noncombustible charring products"""
    y = {m: np.log([v for v in observed(ctx, m)[1] if v is not None]) for m in CHAR}
    ss = sum(((v - v.mean()) ** 2).sum() for v in y.values())
    dof = sum(len(v) - 1 for v in y.values())
    s_rep = float(np.sqrt(ss / dof))
    mu_tga = {m: float(np.log(tga_t600(ctx, m))) for m in CHAR}
    resid = {m: float(y[m].mean() - mu_tga[m]) for m in CHAR}
    tau2 = float(np.mean([resid[m] ** 2 - s_rep ** 2 / len(y[m]) for m in CHAR]))
    means = {m: float(y[m].mean()) for m in CHAR}
    return dict(s_rep=s_rep, tau=float(np.sqrt(max(tau2, TAU_MIN ** 2))), tau2_raw=tau2,
                sd_between=float(np.std(list(means.values()), ddof=1)), ybar=means,
                n={m: len(y[m]) for m in CHAR}, mu_tga=mu_tga, resid=resid)


# ------------------------------------------------------------------------------------------------ evidence modes
def specs_for(mode, ctx, st):
    """per product: base t600 values (equal weights, None = never reached 600 K), SD of a common log shift of the base
    values (epistemic, product mean), p_B beta parameters ('qfra' = the Jeffreys draws of facade_fra/qfra.py), source"""
    out = {}
    s_rep = st["s_rep"]
    two = lambda mu: [float(np.exp(mu - s_rep)), float(np.exp(mu + s_rep))]
    for m in MATS:
        obs, vals, k, n = observed(ctx, m)
        f400 = float(ctx["air"].loc[m, "f_above400_mean"])
        a_t, b_t = pb_prior(f400)
        grade, melt = qf.GRADE[m], f400 < F_MELT
        others = [st["ybar"][c] for c in CHAR if c != m]
        mu_pop = float(np.mean(others))
        if mode == "L2":
            sp = dict(base=vals, sd=0.0, pb="qfra", a=k + 0.5, b=n - k + 0.5, src="중규모 시험값")
        elif mode == "L0":
            sp = (dict(base=vals, sd=0.0, pb="beta", a=0.5, b=0.5, src="가연성 DB(보수적 경계)") if grade == "c" else
                  dict(base=two(mu_pop), sd=st["sd_between"], pb="beta", a=0.5, b=0.5, src="준불연 모집단"))
        elif mode == "L1":
            if grade == "c":
                sp = dict(base=vals, sd=0.0, pb="beta", a=0.5, b=0.5, src="게이트: 가연성 → TGA 미적용(DB 경계)")
            elif melt:
                sp = dict(base=two(mu_pop), sd=st["sd_between"], pb="beta", a=a_t, b=b_t, src="게이트: 용융형 → 확산은 모집단")
            else:
                sp = dict(base=two(np.log(tga_t600(ctx, m))), sd=st["tau"], pb="beta", a=a_t, b=b_t, src="TGA 전이함수")
        elif mode == "L1u":
            sp = dict(base=two(np.log(tga_t600(ctx, m))), sd=st["tau"], pb="beta", a=a_t, b=b_t, src="TGA 단독")
        elif mode == "L1+L2":
            if m in CHAR:
                prec0, precd = 1.0 / st["tau"] ** 2, st["n"][m] / s_rep ** 2
                mu_post = (st["mu_tga"][m] * prec0 + st["ybar"][m] * precd) / (prec0 + precd)
                shift = float(np.exp(mu_post - st["ybar"][m]))
                sp = dict(base=[v * shift for v in vals], sd=float(1 / np.sqrt(prec0 + precd)), pb="beta",
                          a=a_t + k, b=b_t + n - k, src="TGA 사전분포 × 시험", mu_post=float(mu_post))
            elif grade == "c":
                sp = dict(base=vals, sd=0.0, pb="beta", a=k + 0.5, b=n - k + 0.5, src="시험(가연성: TGA 미적용)")
            else:
                sp = dict(base=vals, sd=0.0, pb="beta", a=a_t + k, b=b_t + n - k, src="시험 + TGA(지속성만)")
        else:
            raise ValueError(mode)
        out[m] = sp
    return out


# ------------------------------------------------------------------------------------------------ building risk
def point_risk(specs, s=None, pb_override=None):
    p0 = {k: qf.seconds(k, v) for k, v in qf.P.items()}
    if s is not None:
        p0["s"] = s
    cons = {("L", 1): qf.consequence_table(p0, [None], 0.0, 1, 1.0, cavity=False)}
    for m in MATS:
        sp = specs[m]
        pb = sp["a"] / (sp["a"] + sp["b"]) if pb_override is None else pb_override
        for ie in (1, 2):
            for barrier in (False, True):
                q = 1.0 - p0["p_cb"] if barrier else 1.0
                cons[(m, ie, barrier)] = qf.consequence_table(p0, sp["base"], pb, ie, q)
    return {c: qf.risk(p0, cons, *c) for c in CASES}


def run_mc(specs, n_outer=qf.N_OUTER, n_inner=qf.N_INNER, seed=qf.SEED, s_fixed=None, ctx=None, pb_zero=False):
    """two-loop Monte Carlo of facade_fra/qfra.py with the product inputs of one evidence mode. With the default sizes
    and seed the building-side random numbers are those of qfra.py, so mode L2 reproduces its results."""
    rng = np.random.default_rng(seed)
    eps = qf.sample_epistemic(rng, n_outer)
    use = ctx["tests"][ctx["tests"].used]
    counts = qf.flag_b_counts(use)
    pb_q = {m: rng.beta(k + 0.5, n - k + 0.5, n_outer) for m, (k, n) in counts.items()}
    lhs = lambda: (rng.permutation(n_inner) + rng.uniform(0, 1, n_inner)) / n_inner
    z = np.vectorize(NormalDist().inv_cdf)
    ale = dict(z_fs=z(lhs()), z_rset=z(lhs()), z_l=z(lhs()), u_exp=lhs())
    ale = {k: v.reshape(1, -1, 1) for k, v in ale.items()}
    rng2 = np.random.default_rng(seed + 1000)
    z_mu = {m: rng2.standard_normal(n_outer) for m in MATS}
    u_b = {m: rng2.uniform(size=n_outer) for m in MATS}
    res = {c: {"R_S": [], "R_P": [], "exceed": []} for c in CASES}
    for start in range(0, n_outer, qf.BLOCK):
        sl = slice(start, start + qf.BLOCK)
        p = {k: qf.seconds(k, v) for k, v in qf.P.items()}
        for k in qf.EPISTEMIC:
            p[k] = qf.seconds(k, eps[k][sl]).reshape(-1, 1, 1)
        if s_fixed is not None:
            p["s"] = np.full_like(p["s"], s_fixed)
        cons = {("L", 1): qf.consequence_table(p, [None], 0.0, 1, 1.0, cavity=False, aleatory=ale)}
        for m in MATS:
            sp = specs[m]
            shift = np.exp(sp["sd"] * z_mu[m][sl]).reshape(-1, 1, 1)
            t600 = [None if v is None else v * shift for v in sp["base"]]
            pb = pb_q[m][sl] if sp["pb"] == "qfra" else beta_dist.ppf(u_b[m][sl], sp["a"], sp["b"])
            pb = pb.reshape(-1, 1, 1) * (0.0 if pb_zero else 1.0)
            for ie in (1, 2):
                for barrier in (False, True):
                    q = 1.0 - p["p_cb"] if barrier else 1.0
                    cons[(m, ie, barrier)] = qf.consequence_table(p, t600, pb, ie, q, aleatory=ale)
        pf = {k: (v.reshape(-1) if isinstance(v, np.ndarray) else v) for k, v in p.items()}
        for c in CASES:
            r = qf.risk(pf, cons, *c)
            for k in res[c]:
                res[c][k].append(np.asarray(r[k]))
    return {c: {k: np.concatenate(v, axis=0) for k, v in d.items()} for c, d in res.items()}


def summarize(mode, pt, mc):
    rows = []
    for (m, b, g) in CASES:
        r, base = mc[(m, b, g)], mc[("NC", b, g)]
        rs, ex = r["R_S"], r["exceed"]
        conf = float((ex[:, 1] <= qf.P["crit_f2"]).mean())
        verdict = "충족" if conf >= qf.P["crit_conf"] else ("불충족" if conf <= 1 - qf.P["crit_conf"] else "판단 유보")
        q = lambda v, a: float(np.quantile(v, a))
        rows.append(dict(mode=mode, mode_ko=MODE_KO[mode], material=m, label=LABEL[m], mitigation=g,
                         R_S_point=float(pt[(m, b, g)]["R_S"]), F_n2_point=float(pt[(m, b, g)]["exceed"][1]),
                         MF_point=float(pt[(m, b, g)]["R_S"] / pt[("NC", b, g)]["R_S"]),
                         R_S_mean=float(rs.mean()), R_S_p05=q(rs, .05), R_S_p50=q(rs, .5), R_S_p95=q(rs, .95),
                         R_P_mean=float(r["R_P"].mean()), F_n2_mean=float(ex[:, 1].mean()),
                         F_n2_p05=q(ex[:, 1], .05), F_n2_p95=q(ex[:, 1], .95),
                         MF_p50=q(rs / base["R_S"], .5), MF_p05=q(rs / base["R_S"], .05), MF_p95=q(rs / base["R_S"], .95),
                         conf_meets_F2=conf, verdict=verdict))
    return rows


# ------------------------------------------------------------------------------------------------ cross-scale tables
def beta_binom_pmf(k, n, a, b):
    lc = gammaln(n + 1) - gammaln(k + 1) - gammaln(n - k + 1)
    return float(np.exp(lc + betaln(k + a, n - k + b) - betaln(a, b)))


def consistency(ctx, st):
    rows = []
    for m in MATS:
        obs, vals, k, n = observed(ctx, m)
        reached = [v for v in vals if v is not None]
        f400 = float(ctx["air"].loc[m, "f_above400_mean"])
        grade, melt = qf.GRADE[m], f400 < F_MELT
        mu_hat = np.log(tga_t600(ctx, m))
        ybar = float(np.mean(np.log(reached)))
        sd = np.sqrt(st["tau"] ** 2 + st["s_rep"] ** 2 / len(reached))
        z = float((ybar - mu_hat) / sd)
        gate = "가연성 → 미적용" if grade == "c" else ("용융형 → 미적용" if melt else "적용")
        a_t, b_t = pb_prior(f400)
        p_pred = beta_binom_pmf(k, n, a_t, b_t)
        b_applicable = m != "PIR_C"
        a_cnt = int(sum(ctx["rm"].loc[t, "abrupt_flag"] == "A" for t in obs.test if t in ctx["rm"].index))
        rows.append(dict(material=m, label=LABEL[m], grade=grade, gate=gate,
                         t600_tga_s=float(np.exp(mu_hat)), t600_obs_geo_s=float(np.exp(ybar)), n_reached=len(reached), n_tests=n,
                         z_spread=z, spread_ok=bool(abs(z) <= 2), spread_used=gate == "적용",
                         f400_air=f400, pB_prior_mean=a_t / (a_t + b_t), B_k=k, B_n=n,
                         p_pred_B=p_pred if b_applicable else None, persist_ok=(p_pred >= 0.05) if b_applicable else None,
                         melt_type=bool(melt), A_count=a_cnt))
    return pd.DataFrame(rows)


def profile(ctx, st, specs_int):
    """six hazard dimensions x scales, in mid-scale-equivalent terms"""
    rows = []
    for m in ALL:
        a = ctx["air"].loc[m]
        f400 = float(a.f_above400_mean)
        grade = qf.GRADE[m]
        melt = f400 < F_MELT
        rec = dict(material=m, label=LABEL[m], grade=grade, tga_verdict=ctx["v_air"].get(m),
                   T_onset_air=float(a.T_onset_C_mean), T_a05_air=float(a.T_a05_C_mean),
                   r_max_air=float(a.r_max_pct_C_mean), f400_air=f400, melt=bool(melt),
                   t600_tga=tga_t600(ctx, m), pB_tga=pb_prior(f400)[0] / N0)
        rec["FRI_tga"] = 900.0 / rec["t600_tga"]
        if m in MATS:
            obs, vals, k, n = observed(ctx, m)
            reached = [v for v in vals if v is not None]
            idx, rm, ig = ctx["idx"], ctx["rm"], ctx["ig"]
            tl = list(obs.test)
            tig = [ig.loc[t, "t_ig_s"] for t in tl if t in ig.index and not pd.isna(ig.loc[t, "t_ig_s"])]
            tsp = [ig.loc[t, "t_spread_s"] for t in tl if t in ig.index and not pd.isna(ig.loc[t, "t_spread_s"])]
            rec.update(t_ig_obs=float(np.median(tig)) if tig else None, t_sp_obs=float(np.median(tsp)) if tsp else None,
                       t600_obs=float(np.exp(np.mean(np.log(reached)))) if reached else None, n_reached=len(reached), n_tests=n,
                       FRI_obs=float(900.0 / np.exp(np.mean(np.log(reached)))) if reached else None,
                       P_max=int(max(rm.loc[t, "P"] for t in tl)), S_max=int(max(rm.loc[t, "S"] for t in tl)),
                       R_max=int(max(rm.loc[t, "R"] for t in tl)),
                       H600_med=float(np.median([idx.loc[t, "excess_600K_Kmin"] for t in tl])),
                       peak_med=float(np.median([idx.loc[t, "peak_C"] for t in tl])),
                       B_k=k, B_n=n, A_count=int(sum(rm.loc[t, "abrupt_flag"] == "A" for t in tl)),
                       pB_post=specs_int[m]["a"] / (specs_int[m]["a"] + specs_int[m]["b"]))
        rows.append(rec)
    return pd.DataFrame(rows)


def effect_tier2():
    """EFFECT (Lamont & Ingolfsson 2018) Tier 2 for the example building: non-combustible cladding (assumed), 100 mm
    air cavity, residential all-out, 30-50 m, sprinklered, ignition sources at the base of the building"""
    fam = {"EPS_C": "SPF/XPS/EPS/PUR", "EPS_SC": "SPF/XPS/EPS/PUR", "PIR_C": "PIR/Phenolic 또는 PUR",
           "PIR_SC": "PIR/Phenolic 또는 PUR", "PF_out": "PIR/Phenolic", "PF_in": "PIR/Phenolic"}
    rows = []
    for m in ALL:
        rows.append(dict(material=m, label=LABEL[m], family=fam[m], fuel="High", cladding="Low (불연 가정)",
                         ignition="Medium (외벽 하부 발화원)", vertical="연결 > 3 m → 한 단계 상향",
                         likelihood="High", consequence="Moderate–Extreme (주거 all-out, 30–50 m)", outcome="D (Substantial)"))
    return pd.DataFrame(rows)


# ------------------------------------------------------------------------------------------------ main
def main():
    for d in ("results", "figures", "report"):
        (HERE / d).mkdir(exist_ok=True)
    ctx = load()
    st = fit_stats(ctx)
    specs = {mode: specs_for(mode, ctx, st) for mode in MODES}

    # building risk per evidence mode (point estimate and two-loop Monte Carlo)
    rows, mcs = [], {}
    for mode in MODES:
        pt = point_risk(specs[mode])
        mcs[mode] = cached(f"mc_{mode}", lambda: run_mc(specs[mode], ctx=ctx))
        rows += summarize(mode, pt, mcs[mode])
    risk = pd.DataFrame(rows)
    risk.to_csv(HERE / "results" / "risk_by_mode.csv", index=False, encoding="utf-8")

    # L2 must be facade_fra/qfra.py exactly
    q_pt = pd.read_csv(ROOT / "facade_fra" / "results" / "point_estimate.csv")
    q_mc = pd.read_csv(ROOT / "facade_fra" / "results" / "mc_summary.csv")
    worst = 0.0
    for r in risk[risk["mode"] == "L2"].itertuples():
        a = q_pt[(q_pt.material == r.material) & (q_pt.building == "B1") & (q_pt.mitigation == r.mitigation)].iloc[0]
        b = q_mc[(q_mc.material == r.material) & (q_mc.building == "B1") & (q_mc.mitigation == r.mitigation)].iloc[0]
        for x, y in ((r.R_S_point, a.R_S), (r.R_S_mean, b.R_S_mean), (r.R_S_p95, b.R_S_p95), (r.conf_meets_F2, b.conf_meets_F2)):
            worst = max(worst, abs(x - y) / max(abs(y), 1e-30))
    assert worst < 1e-9, f"L2 does not reproduce facade_fra/qfra.py ({worst})"

    # Bayesian fusion of TGA and the cavity tests (charring products), and the spread of the product-mean t600
    fusion = []
    for m in CHAR:
        sp = specs["L1+L2"][m]
        prec0, precd = 1 / st["tau"] ** 2, st["n"][m] / st["s_rep"] ** 2
        obs, vals, k, n = observed(ctx, m)
        a_t, b_t = pb_prior(float(ctx["air"].loc[m, "f_above400_mean"]))
        fusion.append(dict(material=m, label=LABEL[m], r_max_air=float(ctx["air"].loc[m, "r_max_pct_C_mean"]),
                           t600_tga_loo=float(np.exp(st["mu_tga"][m])), tau=st["tau"], t600_obs_geo=float(np.exp(st["ybar"][m])),
                           n=st["n"][m], s_rep=st["s_rep"], w_tga=prec0 / (prec0 + precd),
                           t600_post=float(np.exp(sp["mu_post"])), sd_post=sp["sd"],
                           f400_air=float(ctx["air"].loc[m, "f_above400_mean"]), a0=a_t, b0=b_t, B_k=k, B_n=n,
                           pB_prior=a_t / (a_t + b_t), pB_jeffreys=(k + .5) / (n + 1), pB_post=(a_t + k) / (a_t + b_t + n)))
    fusion = pd.DataFrame(fusion)
    fusion.to_csv(HERE / "results" / "bayes_fusion.csv", index=False, encoding="utf-8")
    z90 = 2 * z_band(0.95)
    width = []
    for m in CHAR:
        sd = {"L0": st["sd_between"], "L1": st["tau"], "L2": st["s_rep"] / np.sqrt(st["n"][m]), "L1+L2": specs["L1+L2"][m]["sd"]}
        width.append(dict(material=m, **{f"ratio90_{k}": float(np.exp(z90 * v)) for k, v in sd.items()},
                          **{f"sd_{k}": float(v) for k, v in sd.items()}))
    width = pd.DataFrame(width)
    width.to_csv(HERE / "results" / "t600_uncertainty.csv", index=False, encoding="utf-8")

    cons = consistency(ctx, st)
    cons.to_csv(HERE / "results" / "consistency.csv", index=False, encoding="utf-8")
    prof = profile(ctx, st, specs["L1+L2"])
    prof.to_csv(HERE / "results" / "material_profile.csv", index=False, encoding="utf-8")
    eff = effect_tier2()
    eff.to_csv(HERE / "results" / "effect_tier2.csv", index=False, encoding="utf-8")

    # material factor and its split into spread and self-sustained burning (L1+L2). Monte Carlo means: at the point
    # estimate the fire service (15 min) acts before the plume ends (20 min), which hides self-sustained burning
    mc_int, mc_nob = mcs["L1+L2"], cached("mc_noB", lambda: run_mc(specs["L1+L2"], ctx=ctx, pb_zero=True))
    pt_int = point_risk(specs["L1+L2"])
    dec = []
    for g in ("M0", "M1"):
        nc = float(mc_int[("NC", "B1", g)]["R_S"].mean())
        for m in MATS:
            r, r0 = float(mc_int[(m, "B1", g)]["R_S"].mean()), float(mc_nob[(m, "B1", g)]["R_S"].mean())
            dec.append(dict(material=m, label=LABEL[m], mitigation=g, R_S_mean=r, R_S_mean_noB=r0, R_S_NC=nc,
                            MF_mean=r / nc, spread_part=(r0 - nc) / (r - nc), persist_part=(r - r0) / (r - nc),
                            R_S_point=float(pt_int[(m, "B1", g)]["R_S"])))
    dec = pd.DataFrame(dec)
    dec.to_csv(HERE / "results" / "material_factor.csv", index=False, encoding="utf-8")

    # value of knowing s (what simulation or a full-scale test has to deliver)
    voi = []
    for s in S_GRID:
        mc = cached(f"voi_{s}", lambda: run_mc(specs["L1+L2"], n_outer=VOI_OUTER, n_inner=VOI_INNER,
                                                  seed=qf.SEED + 7, s_fixed=s, ctx=ctx))
        for (m, b, g) in CASES:
            ex = mc[(m, b, g)]["exceed"][:, 1]
            voi.append(dict(s=s, material=m, mitigation=g, conf_meets_F2=float((ex <= qf.P["crit_f2"]).mean()),
                            F_n2_mean=float(ex.mean()), R_S_mean=float(mc[(m, b, g)]["R_S"].mean())))
    voi = pd.DataFrame(voi)
    voi.to_csv(HERE / "results" / "voi_s.csv", index=False, encoding="utf-8")
    lo, hi = (next(p for p in qf.PARAMS if p[0] == "s")[5], next(p for p in qf.PARAMS if p[0] == "s")[6])
    p_ge = lambda x: float(np.clip(np.log(hi / x) / np.log(hi / lo), 0, 1))     # P(s >= x), s ~ log-uniform(lo, hi)
    thr = []
    for g in ("M0", "M1"):
        for m in MATS:
            d = voi[(voi.material == m) & (voi.mitigation == g)].sort_values("s")
            ss, cf = np.log(d.s.to_numpy()), d.conf_meets_F2.to_numpy()
            lv_meet, lv_fail = qf.P["crit_conf"], 1 - qf.P["crit_conf"]
            s_meet = s_fail = None
            if cf[-1] >= lv_meet:                    # smallest s from which it meets
                i = int(np.flatnonzero(cf >= lv_meet)[0])
                s_meet = float(np.exp(ss[0] if i == 0 else np.interp(lv_meet, cf[i - 1:i + 1], ss[i - 1:i + 1])))
            if cf[0] <= lv_fail:                     # largest s up to which it fails
                i = int(np.flatnonzero(cf <= lv_fail)[-1])
                s_fail = float(np.exp(ss[-1] if i == len(ss) - 1 else np.interp(lv_fail, cf[i:i + 2], ss[i:i + 2])))
            decided = (p_ge(s_meet) if s_meet is not None else 0.0) + (1 - p_ge(s_fail) if s_fail is not None else 0.0)
            thr.append(dict(material=m, label=LABEL[m], mitigation=g, s_meet=s_meet, s_fail=s_fail,
                            conf_min=float(cf.min()), conf_max=float(cf.max()), P_meet_after_L3=p_ge(s_meet) if s_meet else 0.0,
                            P_fail_after_L3=(1 - p_ge(s_fail)) if s_fail else 0.0, P_decided_after_L3=float(min(decided, 1.0))))
    thr = pd.DataFrame(thr)
    thr.to_csv(HERE / "results" / "voi_thresholds.csv", index=False, encoding="utf-8")

    data = dict(
        stats={k: v for k, v in st.items() if k not in ("ybar", "n", "mu_tga", "resid")} | dict(
            ybar={m: float(np.exp(v)) for m, v in st["ybar"].items()}, resid=st["resid"]),
        constants=dict(t_ig_sys=ctx["t_ig_sys"], c_sys=ctx["c_sys"], c_sys_cv=ctx["c_sys_cv"], f_ramp=F_RAMP, n0=N0,
                       f_melt=F_MELT, crit_f2=qf.P["crit_f2"], crit_conf=qf.P["crit_conf"], s_range=[lo, hi]),
        modes=MODES, mode_ko=MODE_KO, label=LABEL, grade=qf.GRADE, color=COLOR, materials=MATS, char=CHAR,
        specs={mode: {m: {k: v for k, v in sp.items()} for m, sp in d.items()} for mode, d in specs.items()},
        risk=json.loads(risk.to_json(orient="records", force_ascii=False)),
        fusion=json.loads(fusion.to_json(orient="records", force_ascii=False)),
        width=json.loads(width.to_json(orient="records", force_ascii=False)),
        consistency=json.loads(cons.to_json(orient="records", force_ascii=False)),
        profile=json.loads(prof.to_json(orient="records", force_ascii=False)),
        effect=json.loads(eff.to_json(orient="records", force_ascii=False)),
        decomposition=json.loads(dec.to_json(orient="records", force_ascii=False)),
        voi=json.loads(voi.to_json(orient="records", force_ascii=False)),
        voi_thresholds=json.loads(thr.to_json(orient="records", force_ascii=False)),
        n_outer=qf.N_OUTER, n_inner=qf.N_INNER, voi_outer=VOI_OUTER, voi_inner=VOI_INNER, s_grid=S_GRID,
        l2_check_rel_err=worst,
    )
    (HERE / "report" / "ms_data.json").write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    figures(ctx, st, risk, cons, voi)

    pd.set_option("display.width", 250)
    pd.set_option("display.max_columns", 40)
    print({k: round(v, 4) if isinstance(v, float) else v for k, v in st.items() if k in ("s_rep", "tau", "tau2_raw", "sd_between")})
    print(fusion.round(4).to_string())
    print(width.round(3).to_string())
    print(cons.round(3).to_string())
    print(risk[risk.mitigation == "M0"][["mode", "material", "R_S_point", "R_S_mean", "R_S_p05", "R_S_p95", "F_n2_mean",
                                          "MF_p50", "conf_meets_F2", "verdict"]].round(6).to_string())
    print(risk[risk.mitigation == "M1"][["mode", "material", "R_S_mean", "conf_meets_F2", "verdict"]].round(6).to_string())
    print(dec.round(4).to_string())
    print(thr.round(3).to_string())
    print("L2 vs facade_fra/qfra.py, worst relative difference:", worst)


def figures(ctx, st, risk, cons, voi):
    plt.rcParams.update({"font.family": qf.korean_font(), "font.size": 10, "axes.unicode_minus": False,
                         "mathtext.fontset": "dejavusans", "axes.spines.top": False, "axes.spines.right": False})
    import math
    pow10 = FuncFormatter(lambda v, _: f"$10^{{{int(round(math.log10(v)))}}}$" if v > 0 else "")

    # 1 TGA-predicted vs measured t600 (cavity test), single-test 95 % band
    fig, ax = plt.subplots(figsize=(6.4, 5.4))
    sd1 = np.sqrt(st["tau"] ** 2 + st["s_rep"] ** 2)
    xs = np.array([20, 3000])
    ax.plot(xs, xs, color="#4B5563", lw=1)
    ax.fill_between(xs, xs * np.exp(-1.96 * sd1), xs * np.exp(1.96 * sd1), color="#9CA3AF", alpha=0.2, lw=0)
    for m in MATS:
        obs, vals, k, n = observed(ctx, m)
        x = tga_t600(ctx, m)
        used = m in CHAR
        for v in vals:
            if v is None:
                ax.annotate("", xy=(x, 2400), xytext=(x, 1800), arrowprops=dict(arrowstyle="->", color=COLOR[m], lw=1.5))
                ax.plot(x, 1800, marker="^", ms=8, color=COLOR[m], mfc="white", mew=1.6)
            else:
                ax.plot(x, v, marker="o", ms=8, color=COLOR[m], mfc=COLOR[m] if used else "white", mew=1.6, ls="none")
        ax.annotate(LABEL[m], (x, [v for v in vals if v is not None][0]), textcoords="offset points",
                    xytext={"PIR_C": (8, -4), "EPS_SC": (-8, 10), "PF_in": (8, -12), "PF_out": (-8, 10), "PIR_SC": (8, -14)}[m],
                    ha="right" if m in ("EPS_SC", "PF_out") else "left", fontsize=9)
    ax.set_xscale("log")
    ax.set_yscale("log")
    ax.set_xlim(300, 2000)
    ax.set_ylim(20, 3000)
    ax.set_xlabel(r"TGA로 예측한 $t_{600}$ (s)  [준불연 탄화형: 해당 제품을 빼고 보정]")
    ax.set_ylabel(r"중공층 시험 $t_{600}$ (s)")
    ax.grid(color="#E5E7EB", lw=0.75, which="both")
    ax.plot([], [], "o", color="#4B5563", label="전이함수 적용(준불연 탄화형)")
    ax.plot([], [], "o", color="#4B5563", mfc="white", label="게이트로 제외(가연성·용융형)")
    ax.plot([], [], "^", color="#4B5563", mfc="white", label="30분 안에 미도달")
    ax.fill_between([], [], [], color="#9CA3AF", alpha=0.2, label="시험 1회 예측 95 % 구간")
    ax.legend(loc="upper left", fontsize=8.5, frameon=False)
    ax.text(1900, 1900 * 0.93, "1:1", ha="right", va="top", fontsize=9, color="#4B5563")
    ax.set_title(r"소형시료(TGA) → 중규모 등가 $t_{600}$: 예측과 시험", loc="left", fontsize=11.5)
    fig.tight_layout()
    fig.savefig(HERE / "figures" / "fig1_tga_vs_midscale.png", dpi=200)
    plt.close(fig)

    # 2 building risk per evidence mode (M0)
    modes = ["L0", "L1", "L2", "L1+L2"]
    marker = {"L0": ("o", "white"), "L1": ("s", "white"), "L2": ("o", None), "L1+L2": ("D", None)}
    fig, ax = plt.subplots(figsize=(10, 4.4))
    order = ["NC"] + MATS
    for i, m in enumerate(order):
        for j, mode in enumerate(modes):
            r = risk[(risk["mode"] == mode) & (risk.material == m) & (risk.mitigation == "M0")].iloc[0]
            x = i + (j - 1.5) * 0.17
            mk, fc = marker[mode]
            ax.plot([x, x], [r.R_S_p05, r.R_S_p95], color=COLOR[m], lw=1.8, alpha=0.85)
            ax.plot(x, r.R_S_mean, mk, color=COLOR[m], mfc=fc or COLOR[m], ms=7, mew=1.5)
    ax.set_yscale("log")
    ax.set_ylim(2e-6, 2e-1)
    ax.yaxis.set_major_formatter(pow10)
    ax.set_xticks(range(len(order)), [LABEL[m].replace(" ", "\n", 1) for m in order], fontsize=9)
    ax.set_ylabel(r"$R_S$  [확산 층 수/동·년]")
    ax.grid(axis="y", color="#E5E7EB", lw=0.75)
    for mode in modes:
        mk, fc = marker[mode]
        ax.plot([], [], mk, color="#4B5563", mfc=fc or "#4B5563", label=MODE_KO[mode])
    ax.legend(ncol=4, frameon=False, fontsize=9, loc="upper left")
    ax.set_title("증거 수준별 외벽 확산 위험도 (스프링클러 설치, 대책 없음, 평균과 90 % 구간)", loc="left", fontsize=11.5)
    fig.tight_layout()
    fig.savefig(HERE / "figures" / "fig2_risk_by_evidence.png", dpi=200)
    plt.close(fig)

    # 3 value of knowing s
    fig, axes = plt.subplots(1, 2, figsize=(10.5, 4.0), sharey=True)
    for ax, g in zip(axes, ("M0", "M1")):
        for m in MATS:
            d = voi[(voi.material == m) & (voi.mitigation == g)].sort_values("s")
            ax.plot(d.s, d.conf_meets_F2, marker="o", ms=4, color=COLOR[m], lw=2, label=LABEL[m])
        ax.axhline(qf.P["crit_conf"], color="#4B5563", lw=0.8, ls="--")
        ax.axhline(1 - qf.P["crit_conf"], color="#4B5563", lw=0.8, ls="--")
        ax.set_xscale("log")
        ax.xaxis.set_major_locator(FixedLocator(S_GRID[::2]))
        ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:.2f}"))
        ax.xaxis.set_minor_locator(NullLocator())
        ax.xaxis.set_minor_formatter(NullFormatter())
        ax.set_xlabel("실물/시험 층당 확산시간 비 s (L3·실대형으로 확정할 값)")
        ax.set_title({"M0": "M0 대책 없음", "M1": "M1 차단재 + 층간 배리어"}[g], loc="left", fontsize=11)
        ax.grid(color="#E5E7EB", lw=0.75)
    axes[0].set_ylabel("기준 충족 비율 (인식론적 표본)")
    axes[0].legend(fontsize=8.5, frameon=False, loc="center right")
    fig.suptitle("시뮬레이션(L3)이 s를 확정하면 판정이 어떻게 갈리는가 (통합 L1+L2, 스프링클러 설치)", x=0.01, ha="left", fontsize=11.5)
    fig.tight_layout()
    fig.savefig(HERE / "figures" / "fig3_value_of_simulation.png", dpi=200)
    plt.close(fig)


if __name__ == "__main__":
    main()
