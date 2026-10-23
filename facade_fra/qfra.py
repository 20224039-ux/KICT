#!/usr/bin/env python3
"""Quantitative façade fire risk assessment (QFRA) for cavity exterior insulation, a worked example of NFPA 551 5.5.

Risk of one building per year, R_j = sum_i F_i C_ij (NFPA 551 eq. 5.5.1.3): an event tree gives the frequency of each
façade exposure, and a storey-climb model gives its consequence. The material enters only through the 3 m cavity
test: t600 at 3000 mm is the time the cavity fire needs to climb one storey (the FRI idea, cavity_fire_index/RISK_MODEL.md),
and flag B says whether the cavity keeps burning after the burner is off (integrated_model).

Event tree (one building, per year)
  IE1  fire in a dwelling -> sprinkler absent or fails -> fully developed -> flames out of the window
  IE2  external source at the foot of the façade (vehicle, waste, piloti parking) that exposes the façade
  then: cavity / insulation exposed (cavity closers) -> climb in the cavity (cavity barriers, fire service),
  and for IE1 window-to-window spread (leapfrog) in parallel.

Consequence of one event (storeys counted above the storey of origin, j = 1 .. N-1)
  tau_F = s t600 h / 3 m                 time per storey in the cavity (s: full scale / test, epistemic)
  n_F   = floor(min(T_FS, T_stop) / tau_F), T_stop = exposure duration unless the cavity is self-sustaining (B)
  a_j   = q^j [j <= n_F],  q = 1 - p_cb   (a cavity barrier stops a crossing with p_cb)
  b_j   = p_L^j [j <= n_L], n_L = floor(T_FS / tau_L)                    (leapfrog, IE1 only)
  P(n_above >= j) = w_j (1 - (1 - a_j)(1 - b_j)),  w_j = (N - j) / N (storey of origin uniform, IE1), 1 (IE2)
  storey j adds its occupants to the exposed count if a path reaches it before RSET (the first path to arrive)

Uncertainty (NFPA 551 4.5, A.5.5.6) runs in two loops. The outer loop samples what we do not know (frequencies, branch
probabilities, s, the medians of T_FS / RSET / tau_L, p_B as Beta(k + 1/2, n - k + 1/2) from the flag-B tests). The
inner loop averages what varies from fire to fire (T_FS, RSET, tau_L, exposure duration, which repeat test the event
behaves like, B); repeat tests and B are enumerated exactly. The point estimate (best values, medians, repeat tests
weighted equally) is the closed form that the Excel calculator reproduces with cell formulas.

Every frequency and branch probability is an illustrative placeholder (inputs/parameters.csv says what to replace it
with). Only t600 and flag B come from measurements.

Writes inputs/parameters.csv, inputs/test_inputs.csv, results/*.csv, figures/*.png, report/qfra_data.json
    python qfra.py
"""
from __future__ import annotations

import glob
import json
import math
from pathlib import Path
from statistics import NormalDist

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib import font_manager
from matplotlib.ticker import FuncFormatter

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
SEED = 551
N_OUTER = 2000          # epistemic samples
N_INNER = 400           # aleatory samples of the continuous variables (Latin hypercube)
BLOCK = 100             # outer samples per vectorised block

MATERIALS = ["EPS_SC", "PIR_C", "PIR_SC", "PF_out", "PF_in"]
LABEL = {"NC": "불연 단열재(기준안)", "EPS_C": "EPS 가연성", "EPS_SC": "EPS 준불연", "PIR_C": "PIR(PUR) 가연성",
         "PIR_SC": "PIR(PUR) 준불연", "PF_out": "PF 준불연 외부용", "PF_in": "PF 준불연 내부용"}
GRADE = {"EPS_C": "c", "EPS_SC": "sc", "PIR_C": "c", "PIR_SC": "sc", "PF_out": "sc", "PF_in": "sc"}
COLOR = {"NC": "#8A8F98", "EPS_SC": "#E87BA4", "PIR_C": "#6250D6", "PIR_SC": "#008300", "PF_out": "#2A78D6", "PF_in": "#E34948"}
BUILDINGS = {"B1": 1, "B0": 0}                      # sprinklers in the dwellings: installed / not installed
MITIGATION = {"M0": "대책 없음", "M1": "개구부 차단재 + 층간 캐비티 배리어"}

# key, symbol, item, unit, best, low, high, distribution, NFPA 551, basis / what replaces it.
# Times in minutes here; the model works in seconds.
PARAMS = [
    ("N", "N", "층수", "층", 15, None, None, "fixed", "4.4.1", "예제 건물(15층 공동주택). 평가 대상 건물 값으로 바꿈"),
    ("h", "h", "층고", "m", 3.0, None, None, "fixed", "4.4.1", "중공층 시험 높이 3 m = 1개 층. 층고가 다르면 τ를 h/3 m 비율로 늘림"),
    ("units", "n_u", "층당 세대 수", "세대", 4, None, None, "fixed", "4.4.2.1", "예제 건물"),
    ("occ", "o", "세대당 재실자", "명", 2.5, None, None, "fixed", "4.4.2.1", "예제 건물(주거). 야간·주간 재실 분포로 바꿀 수 있음"),
    ("n_aff", "n_a", "확산 띠 안의 층당 세대 수", "세대", 2, None, None, "fixed", "4.4.2.3",
     "외벽 화재가 위로 번질 때 한 층에서 닿는 세대(발화 세대 열 + 이웃 1열)"),
    ("lam1", "λ₁", "세대 화재 발생빈도", "/세대·년", 2.0e-4, 1.0e-4, 4.0e-4, "logu", "5.3.1.1",
     "예시값. 국가화재정보시스템(NFDS) 공동주택 화재 건수 ÷ 세대 수(주택총조사)로 대체"),
    ("lam2", "λ₂", "외부 발화원 화재빈도(외벽 하부)", "/동·년", 2.0e-3, 5.0e-4, 8.0e-3, "logu", "5.3.1.1, 6.2.1.3",
     "예시값(자료 부족). 필로티·주차 차량, 쓰레기, 외부 설비 화재 통계로 대체"),
    ("p_spr", "p_s", "스프링클러 작동·제어 확률", "–", 0.90, 0.80, 0.97, "tri", "5.1.4.5",
     "예시값. 국내 스프링클러 작동·효과 통계로 대체"),
    ("p_fo", "p_FO", "제어 실패 화재의 최성기 도달 확률", "–", 0.15, 0.05, 0.30, "tri", "5.1.1.4",
     "예시값. NFDS 연소 확대 범위(발화실 밖 확대) 비율로 대체"),
    ("p_vent", "p_v", "최성기 화재의 창 분출화염 확률", "–", 0.70, 0.50, 0.90, "tri", "4.4.2.3",
     "예시값. 창호 형식(복층유리·방화유리) 반영"),
    ("p_x2", "p_x", "외부 화원의 외벽 노출 확률", "–", 0.30, 0.10, 0.60, "tri", "4.4.2.3",
     "예시값. 화원–외벽 거리, 필로티 구조 반영"),
    ("p_exp0", "p_e0", "중공층·단열재 노출 확률(차단재 없음)", "–", 0.90, 0.70, 1.00, "tri", "4.4.2.3",
     "개구부 주위·하부가 열려 있을 때. 현장조사(EFFECT Tier 2 수준)로 확인"),
    ("p_exp1", "p_e1", "중공층·단열재 노출 확률(불연 차단재 설치)", "–", 0.30, 0.10, 0.50, "tri", "5.1.4.5",
     "예시값. 차단재 시공 품질·시험 근거로 대체"),
    ("p_cb", "p_cb", "캐비티 배리어가 한 층 경계를 막을 확률", "–", 0.70, 0.40, 0.90, "tri", "5.1.4.5",
     "예시값. 배리어를 넣은 중공층 시험·실대형 시험(KS F 8414)으로 대체"),
    ("p_L", "p_L", "창→창 도약 확산 확률(한 층)", "–", 0.15, 0.05, 0.30, "tri", "4.4.2.3",
     "예시값. 스팬드럴 높이·발코니 유무 반영"),
    ("tau_L", "τ_L", "도약 확산 한 층 소요시간(중앙값)", "분", 12, 8, 20, "tri", "5.4",
     "예시값. 구획화재 성장시간(FDS·CFAST)으로 대체"),
    ("t_fs", "T_FS", "외벽 화재 진압 유효 시각(중앙값, 분출 기준)", "분", 15, 10, 25, "tri", "5.1.1.4",
     "예시값. 관할 소방서 도착시간 + 전개시간(고가사다리 도달 높이 포함)으로 대체"),
    ("rset", "RSET", "상부층 피난 완료 시각(중앙값, 분출 기준)", "분", 15, 10, 25, "tri", "4.4.2.4",
     "예시값. 경보 방식·재실자 특성으로 피난 계산해 대체"),
    ("s", "s", "실물/시험 층당 확산시간 비(τ = s·t₆₀₀)", "–", 0.5, 0.25, 1.0, "logu", "5.4.4, 6.2.2.5",
     "가정. 3 m·30 kW 시험을 실물 분출화염 조건으로 옮기는 계수. 실대형 시험(KS F 8414)으로 보정"),
    ("p_d", "p_d", "노출 재실자 사망 확률(선택)", "–", 0.05, 0.01, 0.20, "tri", "A.5.1.2",
     "예시값. 인명 지표로 환산할 때만 사용"),
    ("sig_fs", "σ_FS", "진압 시각의 사건별 변동(로그 표준편차)", "–", 0.5, None, None, "fixed", "A.5.5.6", "우연적 변동 폭(가정)"),
    ("sig_rset", "σ_R", "피난 완료 시각의 사건별 변동(로그 표준편차)", "–", 0.5, None, None, "fixed", "A.5.5.6", "우연적 변동 폭(가정)"),
    ("sig_L", "σ_L", "도약 소요시간의 사건별 변동(로그 표준편차)", "–", 0.3, None, None, "fixed", "A.5.5.6", "우연적 변동 폭(가정)"),
    ("texp1_lo", "t_e1", "분출화염 지속시간 하한(IE1)", "분", 10, None, None, "fixed", "A.5.5.6", "균등분포 하한(가정)"),
    ("texp1_hi", "t_e1", "분출화염 지속시간 상한(IE1)", "분", 30, None, None, "fixed", "A.5.5.6", "균등분포 상한(가정)"),
    ("texp2_lo", "t_e2", "외부 화원 지속시간 하한(IE2)", "분", 5, None, None, "fixed", "A.5.5.6", "균등분포 하한(가정)"),
    ("texp2_hi", "t_e2", "외부 화원 지속시간 상한(IE2)", "분", 20, None, None, "fixed", "A.5.5.6", "균등분포 상한(가정)"),
    ("crit_f2", "F₂*", "허용기준 예시: 발화층 위 2개 층 이상 확산 빈도", "/동·년", 1.0e-4, None, None, "fixed",
     "4.4.3, A.5.1.2", "예시값(규제 기준 아님). 사전협의에서 정함"),
    ("crit_conf", "c*", "기준 충족 판단에 쓰는 신뢰수준", "–", 0.90, None, None, "fixed", "4.5",
     "인식론적 표본 중 기준을 만족하는 비율의 하한(가정)"),
]
P = {p[0]: p[4] for p in PARAMS}
EPISTEMIC = [p[0] for p in PARAMS if p[7] in ("tri", "logu")]
MINUTES = {"tau_L", "t_fs", "rset", "texp1_lo", "texp1_hi", "texp2_lo", "texp2_hi"}


def korean_font():
    """A Hangul font: registered by matplotlib, or found in the usual Windows / macOS / Linux folders."""
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


def seconds(key, v):
    return v * 60.0 if key in MINUTES else v


def test_inputs():
    """t600 at 3000 mm and flag B per cavity test, from the earlier analyses (no raw data needed)."""
    idx = pd.read_csv(ROOT / "cavity_fire_index" / "results" / "index_3000mm.csv")
    fin = pd.read_csv(ROOT / "integrated_model" / "results" / "final_by_test.csv").set_index("test")
    rows = []
    for r in idx.itertuples():
        if r.repeat == 0:                           # PF_out-0: aborted pre-test
            continue
        used = r.material in MATERIALS
        flag_b = int(r.test in fin.index and fin.loc[r.test, "flag_B"] == "B")
        note = ""
        if not used:
            note = f"기록 {r.record_s:.0f} s – t₆₀₀ 산정 불가, 정량평가 제외(재시험 필요, 6.2.1.3)"
        elif pd.isna(r.t600_s):
            note = "30분 안에 3000 mm 600 K 미도달 → 이 사건은 층간 확산 없음"
        elif r.material == "PIR_C":
            note = "전체 화염(30 s): 소화 전에 연료 소진 → 플래그 B 판정 대상 아님"
        rows.append(dict(test=r.test, material=r.material, label=LABEL[r.material], grade=GRADE[r.material],
                         t600_s=None if pd.isna(r.t600_s) else float(r.t600_s), flag_B=flag_b, used=used, note=note))
    return pd.DataFrame(rows)


def flag_b_counts(ti):
    out = {}
    for m in MATERIALS:
        sub = ti[ti.material == m]
        out[m] = (int(sub.flag_B.sum()), int(len(sub)))
    return out


def climb(tau_f, t_stop, t_fs, rset, tau_l, q, p_l, j, w, occ_st):
    """One exposure event. Arguments broadcast against the storey axis j (last axis); tau_f = inf: no cavity path.
    Returns E[storeys above the storey of origin], E[exposed persons], P(n_above >= j) (weights applied)."""
    tau_f = np.asarray(tau_f, dtype=float)
    finite = np.isfinite(tau_f)
    safe = np.where(finite, tau_f, 1.0)
    n_f = np.where(finite, np.floor(np.minimum(t_fs, t_stop) / safe), 0.0)
    n_l = np.floor(t_fs / tau_l)
    a = q ** j * (j <= n_f)
    b = p_l ** j * (j <= n_l)
    reach = 1.0 - (1.0 - a) * (1.0 - b)
    t_a = np.where(finite, safe, np.inf) * j
    t_b = tau_l * j
    seen = a * b * (np.minimum(t_a, t_b) < rset) + a * (1.0 - b) * (t_a < rset) + (1.0 - a) * b * (t_b < rset)
    return (w * reach).sum(-1), occ_st * (w * seen).sum(-1), w * reach


def weights(n_storeys, ie):
    j = np.arange(1, int(n_storeys), dtype=float)
    return j, ((n_storeys - j) / n_storeys if ie == 1 else np.ones_like(j))


def consequence_table(p, t600s, p_b, ie, q, cavity=True, aleatory=None):
    """E[storeys], E[persons], exceedance over the repeat tests (equal weights) and B (weight p_b).
    p: parameter values in seconds (scalars, or arrays shaped (E, 1, 1) in the Monte Carlo).
    aleatory: None for the point estimate, else dict of inner-sample arrays shaped (1, I, 1)."""
    j, w = weights(P["N"], ie)
    occ_st = P["n_aff"] * P["occ"]
    p_l = p["p_L"] if ie == 1 else 0.0
    if aleatory is None:
        t_fs, rset, tau_l = p["t_fs"], p["rset"], p["tau_L"]
        t_exp = (p["texp1_lo"] + p["texp1_hi"]) / 2 if ie == 1 else (p["texp2_lo"] + p["texp2_hi"]) / 2
    else:
        t_fs = p["t_fs"] * np.exp(P["sig_fs"] * aleatory["z_fs"])
        rset = p["rset"] * np.exp(P["sig_rset"] * aleatory["z_rset"])
        tau_l = p["tau_L"] * np.exp(P["sig_L"] * aleatory["z_l"])
        lo, hi = (p["texp1_lo"], p["texp1_hi"]) if ie == 1 else (p["texp2_lo"], p["texp2_hi"])
        t_exp = lo + (hi - lo) * aleatory["u_exp"]
    tests = t600s if cavity else [None]
    en = ep = ex = 0.0
    for t600 in tests:
        tau_f = np.inf if t600 is None else p["s"] * t600 * P["h"] / 3.0
        for flag, wb in ((1, p_b), (0, 1.0 - p_b)):
            if not cavity and flag == 1:
                continue
            wt = (1.0 if not cavity else wb) / len(tests)
            t_stop = np.inf if flag else t_exp
            a, b, c = climb(tau_f, t_stop, t_fs, rset, tau_l, q, p_l, j, w, occ_st)
            if aleatory is not None:
                a, b, c = a.mean(-1), b.mean(-1), c.mean(-2)
            en, ep, ex = en + _sq(wt) * a, ep + _sq(wt) * b, ex + _sq(wt, 1) * c
    return en, ep, ex


def _sq(x, extra=0):
    """weights shaped (E, 1, 1) broadcast against results shaped (E,) or (E, J)."""
    x = np.asarray(x, dtype=float)
    if x.ndim == 0:
        return x
    x = x.reshape(x.shape[0])
    return x[:, None] if extra else x


def _col(x, like):
    """x shaped (E,) against results shaped (E, J): add the storey axis."""
    return np.asarray(x)[:, None] if np.ndim(like) == 2 and np.ndim(x) == 1 else x


def risk(p, cons, material, building, mitigation):
    """Combine frequencies and consequences. cons[(material, ie, barrier_on)] = (E storeys, E persons, exceedance),
    cons[("L", 1)] = the leapfrog-only path of IE1. Returns R_S, R_P, exceedance, F_IE1 ext, F_IE2 ext, share of IE2."""
    spr = BUILDINGS[building]
    credit = mitigation == "M1" and GRADE.get(material) == "sc"     # no credit for c-grade without evidence
    f1 = p["lam1"] * P["units"] * P["N"] * (1 - p["p_spr"] * spr) * p["p_fo"] * p["p_vent"]
    f2 = p["lam2"] * p["p_x2"]
    leap = cons[("L", 1)]
    if material == "NC":            # no cavity path: IE1 events spread by leapfrog only, IE2 does not spread
        pe, s1, s2 = 0.0, leap, None
    else:
        pe = p["p_exp1"] if credit else p["p_exp0"]
        s1, s2 = cons[(material, 1, credit)], cons[(material, 2, credit)]
    out = dict(F1=f1, F2=f2)
    for k, name in enumerate(("R_S", "R_P", "exceed")):
        like = np.asarray(leap[k])
        r1 = _col(f1, like) * (_col(pe, like) * s1[k] + (1 - _col(pe, like)) * leap[k])
        r2 = _col(f2, like) * _col(pe, like) * s2[k] if s2 is not None else 0.0 * like
        out[name] = r1 + r2
        if name == "R_S":
            tot = np.asarray(r1 + r2, dtype=float)
            out["share_IE2"] = np.where(tot > 0, np.asarray(r2, dtype=float) / np.where(tot > 0, tot, 1.0), 0.0)
    return out


def all_consequences(p, t600, pb, aleatory=None):
    cons = {("L", 1): consequence_table(p, [None], 0.0, 1, 1.0, cavity=False, aleatory=aleatory)}
    for m in MATERIALS:
        for ie in (1, 2):
            for barrier in (False, True):
                q = 1.0 - p["p_cb"] if barrier else 1.0
                cons[(m, ie, barrier)] = consequence_table(p, t600[m], pb[m], ie, q, aleatory=aleatory)
    return cons


CASES = [(m, b, g) for b in BUILDINGS for g in MITIGATION for m in ["NC"] + MATERIALS]


def point_estimate(p, t600, pb):
    cons = all_consequences(p, t600, pb)
    out = {}
    for m, b, g in CASES:
        out[(m, b, g)] = risk(p, cons, m, b, g)
    return out, cons


def sample_epistemic(rng, n):
    s = {}
    for key, _, _, _, best, lo, hi, dist, _, _ in PARAMS:
        if dist == "tri":
            s[key] = rng.triangular(lo, best, hi, n)
        elif dist == "logu":
            s[key] = np.exp(rng.uniform(math.log(lo), math.log(hi), n))
    return s


def monte_carlo(t600, counts):
    rng = np.random.default_rng(SEED)
    eps = sample_epistemic(rng, N_OUTER)
    pb = {m: rng.beta(k + 0.5, n - k + 0.5, N_OUTER) for m, (k, n) in counts.items()}
    lhs = lambda: (rng.permutation(N_INNER) + rng.uniform(0, 1, N_INNER)) / N_INNER
    z = np.vectorize(NormalDist().inv_cdf)
    ale = dict(z_fs=z(lhs()), z_rset=z(lhs()), z_l=z(lhs()), u_exp=lhs())
    ale = {k: v.reshape(1, -1, 1) for k, v in ale.items()}
    res = {c: {"R_S": [], "R_P": [], "exceed": []} for c in CASES}
    for start in range(0, N_OUTER, BLOCK):
        sl = slice(start, start + BLOCK)
        p = {k: seconds(k, v) for k, v in P.items()}
        for k in EPISTEMIC:
            p[k] = seconds(k, eps[k][sl]).reshape(-1, 1, 1)
        pbs = {m: pb[m][sl].reshape(-1, 1, 1) for m in MATERIALS}
        cons = all_consequences(p, t600, pbs, aleatory=ale)
        pf = {k: (v.reshape(-1) if isinstance(v, np.ndarray) else v) for k, v in p.items()}
        for c in CASES:
            r = risk(pf, cons, *c)
            for k in ("R_S", "R_P", "exceed"):
                res[c][k].append(np.asarray(r[k]))
    for c in CASES:
        for k in res[c]:
            res[c][k] = np.concatenate(res[c][k], axis=0)
    return res, eps, pb


def rank_corr(x, y):
    rx, ry = pd.Series(x).rank().to_numpy(), pd.Series(y).rank().to_numpy()
    return float(np.corrcoef(rx, ry)[0, 1])


def main():
    (HERE / "inputs").mkdir(exist_ok=True)
    (HERE / "results").mkdir(exist_ok=True)
    (HERE / "figures").mkdir(exist_ok=True)
    pd.DataFrame(PARAMS, columns=["key", "symbol", "item", "unit", "best", "low", "high", "distribution", "nfpa551", "basis"]) \
        .to_csv(HERE / "inputs" / "parameters.csv", index=False, encoding="utf-8")
    ti = test_inputs()
    ti.to_csv(HERE / "inputs" / "test_inputs.csv", index=False, encoding="utf-8")
    use = ti[ti.used]
    t600 = {m: [None if pd.isna(v) else float(v) for v in use[use.material == m].t600_s] for m in MATERIALS}
    counts = flag_b_counts(use)
    pb_best = {m: (k + 0.5) / (n + 1.0) for m, (k, n) in counts.items()}

    # ------------------------------------------------------------------ point estimate (= Excel calculator)
    p0 = {k: seconds(k, v) for k, v in P.items()}
    det, cons = point_estimate(p0, t600, pb_best)
    occupants = P["N"] * P["units"] * P["occ"]
    rows = []
    for (m, b, g), r in det.items():
        base = det[("NC", b, g)]
        rows.append(dict(material=m, label=LABEL[m], building=b, mitigation=g, F_IE1_ext=r["F1"], F_IE2_ext=r["F2"],
                         R_S=r["R_S"], R_P=r["R_P"], R_fatal=P["p_d"] * r["R_P"], IR=P["p_d"] * r["R_P"] / occupants,
                         F_n1=r["exceed"][0], F_n2=r["exceed"][1], F_n5=r["exceed"][4],
                         RR_S=r["R_S"] / base["R_S"], share_IE2=float(r["share_IE2"]),
                         meets_F2=bool(r["exceed"][1] <= P["crit_f2"])))
    det_df = pd.DataFrame(rows)
    det_df.to_csv(HERE / "results" / "point_estimate.csv", index=False, encoding="utf-8")

    # conditional consequences per event (what the cavity tests become)
    crow = []
    for m in ["NC"] + MATERIALS:
        for ie in (1, 2):
            for barrier in (False, True):
                if m == "NC":
                    c = cons[("L", 1)] if ie == 1 else (0.0, 0.0, np.zeros(int(P["N"]) - 1))
                else:
                    c = cons[(m, ie, barrier)]
                crow.append(dict(material=m, ie=f"IE{ie}", barrier=barrier, E_storeys=float(c[0]), E_persons=float(c[1]),
                                 P_ge1=float(c[2][0]), P_ge2=float(c[2][1]), P_ge5=float(c[2][4])))
    cons_df = pd.DataFrame(crow)
    cons_df.to_csv(HERE / "results" / "conditional_consequence.csv", index=False, encoding="utf-8")

    # ------------------------------------------------------------------ Monte Carlo (two loops)
    mc, eps, pb = monte_carlo(t600, counts)
    srows, erows = [], []
    for c in CASES:
        m, b, g = c
        base = mc[("NC", b, g)]
        rs, rp, ex = mc[c]["R_S"], mc[c]["R_P"], mc[c]["exceed"]
        rr = rs / base["R_S"]
        meet = ex[:, 1] <= P["crit_f2"]
        conf = float(meet.mean())
        verdict = "충족" if conf >= P["crit_conf"] else ("불충족" if conf <= 1 - P["crit_conf"] else "판단 유보")
        q = lambda v, a: float(np.quantile(v, a))
        srows.append(dict(material=m, label=LABEL[m], building=b, mitigation=g,
                          R_S_mean=float(rs.mean()), R_S_p05=q(rs, .05), R_S_p50=q(rs, .5), R_S_p95=q(rs, .95),
                          R_P_mean=float(rp.mean()), R_P_p05=q(rp, .05), R_P_p95=q(rp, .95),
                          F_n2_mean=float(ex[:, 1].mean()), F_n2_p05=q(ex[:, 1], .05), F_n2_p50=q(ex[:, 1], .5), F_n2_p95=q(ex[:, 1], .95),
                          RR_S_p05=q(rr, .05), RR_S_p50=q(rr, .5), RR_S_p95=q(rr, .95),
                          conf_meets_F2=conf, verdict=verdict))
        for n in range(1, int(P["N"])):
            erows.append(dict(material=m, building=b, mitigation=g, n=n, F_point=float(det[c]["exceed"][n - 1]),
                              F_mean=float(ex[:, n - 1].mean()), F_p05=q(ex[:, n - 1], .05), F_p50=q(ex[:, n - 1], .5),
                              F_p95=q(ex[:, n - 1], .95)))
    mc_df = pd.DataFrame(srows)
    mc_df.to_csv(HERE / "results" / "mc_summary.csv", index=False, encoding="utf-8")
    ex_df = pd.DataFrame(erows)
    ex_df.to_csv(HERE / "results" / "exceedance.csv", index=False, encoding="utf-8")

    # ------------------------------------------------------------------ sensitivity: rank correlation and tornado
    sens = []
    for g in MITIGATION:
        for m in MATERIALS:
            c = (m, "B1", g)
            y = np.log(mc[c]["R_S"])
            y2 = np.log(mc[c]["R_S"] / mc[("NC", "B1", g)]["R_S"])
            y3 = mc[c]["exceed"][:, 1]
            for k in EPISTEMIC + ["p_B"]:
                x = pb[m] if k == "p_B" else eps[k]
                sens.append(dict(mitigation=g, material=m, parameter=k, rho_R_S=rank_corr(x, y),
                                 rho_RR_S=rank_corr(x, y2), rho_F_n2=rank_corr(x, y3)))
    sens_df = pd.DataFrame(sens)
    sens_df.to_csv(HERE / "results" / "sensitivity_rank.csv", index=False, encoding="utf-8")

    torn = []
    for k in EPISTEMIC + ["p_B"]:
        for side in ("low", "high"):
            p1 = dict(p0)
            pbv = dict(pb_best)
            if k == "p_B":
                pbv = {m: (0.5 / (counts[m][1] + 1) if side == "low" else (counts[m][1] + 0.5) / (counts[m][1] + 1)) for m in MATERIALS}
            else:
                spec = next(x for x in PARAMS if x[0] == k)
                p1[k] = seconds(k, spec[5] if side == "low" else spec[6])
            d1, _ = point_estimate(p1, t600, pbv)
            for m in ["NC"] + MATERIALS:
                torn.append(dict(parameter=k, side=side, material=m, R_S=float(d1[(m, "B1", "M0")]["R_S"]),
                                 F_n2=float(d1[(m, "B1", "M0")]["exceed"][1])))
    torn_df = pd.DataFrame(torn)
    torn_df.to_csv(HERE / "results" / "tornado.csv", index=False, encoding="utf-8")

    # ------------------------------------------------------------------ check: the inner loop collapses to the point estimate
    p_chk = {k: (np.array([[[v]]]) if k in EPISTEMIC else v) for k, v in p0.items()}
    zero = dict(z_fs=np.zeros((1, 1, 1)), z_rset=np.zeros((1, 1, 1)), z_l=np.zeros((1, 1, 1)), u_exp=np.full((1, 1, 1), 0.5))
    pbc = {m: np.array([[[pb_best[m]]]]) for m in MATERIALS}
    cons_chk = all_consequences(p_chk, t600, pbc, aleatory=zero)
    pf = {k: (v.reshape(-1) if isinstance(v, np.ndarray) else v) for k, v in p_chk.items()}
    worst = max(abs(float(np.asarray(risk(pf, cons_chk, *c)["R_S"]).ravel()[0]) - float(det[c]["R_S"])) / max(float(det[c]["R_S"]), 1e-30)
                for c in CASES)
    assert worst < 1e-9, worst

    data = dict(
        params=[dict(zip(["key", "symbol", "item", "unit", "best", "low", "high", "distribution", "nfpa551", "basis"], p)) for p in PARAMS],
        tests=json.loads(ti.to_json(orient="records", force_ascii=False)),
        flag_b={m: dict(k=k, n=n, p_best=pb_best[m]) for m, (k, n) in counts.items()},
        label=LABEL, grade=GRADE, materials=MATERIALS, color=COLOR, mitigation=MITIGATION,
        point=json.loads(det_df.to_json(orient="records", force_ascii=False)),
        conditional=json.loads(cons_df.to_json(orient="records", force_ascii=False)),
        mc=json.loads(mc_df.to_json(orient="records", force_ascii=False)),
        exceedance=json.loads(ex_df.to_json(orient="records", force_ascii=False)),
        sensitivity=json.loads(sens_df.to_json(orient="records", force_ascii=False)),
        tornado=json.loads(torn_df.to_json(orient="records", force_ascii=False)),
        n_outer=N_OUTER, n_inner=N_INNER, seed=SEED, occupants=occupants, mc_check_rel_err=worst,
    )
    (HERE / "report").mkdir(exist_ok=True)
    (HERE / "report" / "qfra_data.json").write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    figures(det, ex_df, mc_df, sens_df)

    pd.set_option("display.width", 220)
    pd.set_option("display.max_columns", 30)
    print(ti.to_string())
    print(det_df[det_df.building == "B1"].round(6).to_string())
    print(mc_df[mc_df.building == "B1"].round(6).to_string())
    print(cons_df.round(4).to_string())
    print("MC inner-loop check, worst relative difference:", worst)


def figures(det, ex_df, mc_df, sens_df):
    plt.rcParams.update({"font.family": korean_font(), "font.size": 10, "axes.unicode_minus": False,
                         "mathtext.fontset": "dejavusans", "axes.spines.top": False, "axes.spines.right": False})
    pow10 = FuncFormatter(lambda v, _: f"$10^{{{int(round(math.log10(v)))}}}$" if v > 0 else "")
    order = ["NC"] + MATERIALS
    # 1 exceedance curves, one panel per material (B1): M0 and M1 means with 90 % epistemic bands, baseline in grey
    fig, axes = plt.subplots(1, len(MATERIALS), figsize=(13, 3.9), sharey=True)
    nc = ex_df[(ex_df.material == "NC") & (ex_df.building == "B1") & (ex_df.mitigation == "M0")]
    for ax, m in zip(axes, MATERIALS):
        ax.plot(nc.n, nc.F_mean.clip(lower=1e-9), color=COLOR["NC"], lw=1.5, label="불연 기준안")
        for g, ls, alpha in (("M0", "-", 0.16), ("M1", "--", 0.10)):
            d = ex_df[(ex_df.material == m) & (ex_df.building == "B1") & (ex_df.mitigation == g)]
            ax.fill_between(d.n, d.F_p05.clip(lower=1e-9), d.F_p95.clip(lower=1e-9), color=COLOR[m], alpha=alpha, lw=0)
            ax.plot(d.n, d.F_mean.clip(lower=1e-9), color=COLOR[m], lw=2, ls=ls, label=f"{g} {MITIGATION[g]}")
        ax.plot([2], [P["crit_f2"]], "o", color="#23303D", ms=5, zorder=5)
        ax.axhline(P["crit_f2"], color="#4B5563", lw=0.8, ls=":")
        ax.set_yscale("log")
        ax.yaxis.set_major_formatter(pow10)
        ax.set_ylim(1e-8, 1e-2)
        ax.set_xlim(1, 14)
        ax.set_xticks([1, 2, 5, 10, 14])
        ax.set_title(LABEL[m], fontsize=10.5, loc="left", color="#23303D")
        if GRADE[m] == "c":
            ax.text(13.6, 2.2e-5, "M1 = M0\n(가연성 등급은\n대책 효과 불인정)", ha="right", va="top", fontsize=8.5, color="#4B5563")
        ax.grid(axis="y", color="#E5E7EB", lw=0.75)
        ax.set_xlabel("확산 층 수 n")
    axes[0].set_ylabel("F(발화층 위 확산 ≥ n)  [/동·년]")
    handles = [plt.Line2D([], [], color="#4B5563", lw=2, ls="-"), plt.Line2D([], [], color="#4B5563", lw=2, ls="--"),
               plt.Line2D([], [], color=COLOR["NC"], lw=1.5),
               plt.Line2D([], [], color="#23303D", marker="o", ls=":", ms=5, lw=0.8)]
    fig.legend(handles, ["M0 대책 없음", "M1 개구부 차단재 + 캐비티 배리어", "불연 단열재(기준안)",
                         r"허용기준 예시: F(n ≥ 2) ≤ $1\times10^{-4}$"],
               loc="lower center", ncol=4, frameon=False, fontsize=9.5, bbox_to_anchor=(0.5, -0.01))
    fig.suptitle("외벽 확산 초과빈도 곡선 — 스프링클러 설치 건물, 선: 인식론적 평균, 음영: 90 % 구간", x=0.01, ha="left", fontsize=12)
    fig.tight_layout(rect=(0, 0.07, 1, 1))
    fig.savefig(HERE / "figures" / "fig1_exceedance.png", dpi=200)
    plt.close(fig)

    # 2 R_S with intervals
    fig, ax = plt.subplots(figsize=(9, 4.2))
    xs = np.arange(len(order))
    for g, dx, mk in (("M0", -0.15, "o"), ("M1", 0.15, "s")):
        for i, m in enumerate(order):
            r = mc_df[(mc_df.material == m) & (mc_df.building == "B1") & (mc_df.mitigation == g)].iloc[0]
            ax.plot([i + dx, i + dx], [r.R_S_p05, r.R_S_p95], color=COLOR[m], lw=2, alpha=0.9 if g == "M0" else 0.55)
            ax.plot(i + dx, r.R_S_mean, mk, color=COLOR[m], ms=7, mfc=COLOR[m] if g == "M0" else "white", mew=1.6)
    ax.set_yscale("log")
    ax.yaxis.set_major_formatter(pow10)
    ax.set_xticks(xs, [LABEL[m].replace(" ", "\n", 1) for m in order], fontsize=9)
    ax.set_ylabel("R_S  [확산 층 수/동·년]")
    ax.grid(axis="y", color="#E5E7EB", lw=0.75)
    ax.plot([], [], "o", color="#4B5563", label="M0 대책 없음 (평균, 90 % 구간)")
    ax.plot([], [], "s", color="#4B5563", mfc="white", label="M1 차단재 + 캐비티 배리어")
    ax.legend(frameon=False, fontsize=9, loc="upper left")
    ax.set_title("외벽 확산 위험도 R_S (스프링클러 설치 건물)", loc="left", fontsize=12)
    fig.tight_layout()
    fig.savefig(HERE / "figures" / "fig2_risk.png", dpi=200)
    plt.close(fig)

    # 3 rank correlation of each epistemic input with log R_S (B1), M0 and M1
    keys = [k for k in EPISTEMIC if k != "p_d"] + ["p_B"]
    sym = dict(lam1=r"$\lambda_1$", lam2=r"$\lambda_2$", p_spr=r"$p_s$", p_fo=r"$p_{FO}$", p_vent=r"$p_v$", p_x2=r"$p_x$",
               p_exp0=r"$p_{e0}$", p_exp1=r"$p_{e1}$", p_cb=r"$p_{cb}$", p_L=r"$p_L$", tau_L=r"$\tau_L$", t_fs=r"$T_{FS}$",
               rset="RSET", s=r"$s$", p_B=r"$p_B$")
    fig, axes = plt.subplots(1, 2, figsize=(11, 6), sharey=True)
    for ax, g in zip(axes, MITIGATION):
        mat = np.array([[sens_df[(sens_df.mitigation == g) & (sens_df.material == m) & (sens_df.parameter == k)].rho_R_S.iloc[0]
                         for m in MATERIALS] for k in keys])
        im = ax.imshow(mat, cmap="RdBu_r", vmin=-1, vmax=1, aspect="auto")
        ax.set_xticks(range(len(MATERIALS)), [LABEL[m].replace(" ", "\n", 1) for m in MATERIALS], fontsize=8.5)
        ax.set_yticks(range(len(keys)), [sym[k] for k in keys], fontsize=9)
        for i in range(len(keys)):
            for k in range(len(MATERIALS)):
                ax.text(k, i, f"{mat[i, k]:+.2f}", ha="center", va="center", fontsize=8,
                        color="white" if abs(mat[i, k]) > 0.55 else "#23303D")
        ax.set_title(f"{g} {MITIGATION[g]}", loc="left", fontsize=11)
    fig.colorbar(im, ax=axes, fraction=0.03, label="순위상관 ρ (log R_S)")
    fig.suptitle("인식론적 입력이 외벽 확산 위험도에 미치는 영향 (스프링클러 설치 건물)", x=0.01, ha="left", fontsize=12)
    fig.savefig(HERE / "figures" / "fig3_sensitivity.png", dpi=200, bbox_inches="tight")
    plt.close(fig)


if __name__ == "__main__":
    main()
