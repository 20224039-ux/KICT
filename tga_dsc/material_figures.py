#!/usr/bin/env python3
"""Material TGA and DSC figures with Korean labels: the six materials side by side, and one sheet per material.

    figures/material_tga.png        mass and DTG, six materials overlaid, N2 and air (O) panels
    figures/material_dsc.png        local heat flow, six materials offset, N2 and air (O) panels
    figures/material_<KEY>.png      one material: mass, DTG, DSC (N2 grey, air in the material colour) and its numbers

Curves are the mean of the two runs with the run-to-run range shaded; the DSC shows both runs (solid #1, dashed #2)
because their sizes differ with the unrecorded sample mass. Run after analysis.py, qa_reliability.py and
pyrosim_inputs.py (it reads their results).

    python material_figures.py
"""
from __future__ import annotations

import glob
import importlib.util
import os
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402
from matplotlib import font_manager  # noqa: E402

HERE = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location("tga_analysis", HERE / "analysis.py")
tga = importlib.util.module_from_spec(spec)
spec.loader.exec_module(tga)

GRID = np.arange(60.0, 786.0, 1.0)
DSC_WIN = (150.0, 720.0)          # same window as the reproducibility check (away from the ramp ends)
ORDER = ["EPS_C", "EPS_SC", "PIR_C", "PIR_SC", "PF_out", "PF_in"]
LABEL = {key: ko for key, ko, _ in tga.MATERIALS.values()}
# categorical slots in an order that passes the palette check (adjacent CVD and normal-vision separation)
COLOR = {"EPS_C": "#eda100", "EPS_SC": "#e87ba4", "PIR_C": "#6250d6", "PIR_SC": "#008300", "PF_out": "#2a78d6",
         "PF_in": "#e34948"}
N2_COLOR = "#52514e"
ATM = {"N": "N$_2$ (열분해)", "O": "공기(O) (산화)"}
DSC_EVENTS = {("EPS_C", "N"): (398, "398 °C 흡열 (해중합) — 두 반복 재현"),
              ("PF_in", "O"): (474, "474 °C 발열 (탄화층 산화) — 두 반복 재현")}
DSC_SCALE = {("PF_in", "O"): 0.2}      # drawn at 1/5 so the other traces stay readable
NOTE_FONT = 9


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


def load():
    runs = {}
    for atm in ("N", "O"):
        for f in sorted(glob.glob(os.path.join(HERE / "data", atm, "*.csv"))):
            stem, rep = tga.stem(os.path.basename(f))
            r = tga.process(f, atm)
            runs[(tga.MATERIALS[stem][0], atm, rep)] = {k: np.interp(GRID, r["T"], r[k]) for k in ("rem", "dtg", "h")}
    return runs


def band(runs, m, atm, key):
    v = np.array([runs[(m, atm, k)][key] for k in (1, 2)])
    return v.mean(axis=0), v.min(axis=0), v.max(axis=0)


def peak_label(ax, x, y, text, color, dx=0, dy=6, ha="center", at=None):
    """Dot on the peak and its label; `at` = (x, y) in data units puts the label there with a leader line."""
    ax.plot([x], [y], "o", ms=5, color=color, mec="white", mew=1.0, zorder=5)
    if at is None:
        ax.annotate(text, (x, y), xytext=(dx, dy), textcoords="offset points", ha=ha, va="bottom",
                    fontsize=8.5, color=tga.INK, zorder=6)
    else:
        ax.annotate(text, (x, y), xytext=at, textcoords="data", ha=ha, va="center", fontsize=8.5, color=tga.INK,
                    zorder=6, arrowprops=dict(arrowstyle="-", color=tga.INK2, lw=0.7, shrinkA=2, shrinkB=3))


def fig_overview_tga(runs, stages, path):
    fig, axs = plt.subplots(2, 2, figsize=(15, 9.4), sharex=True, gridspec_kw=dict(height_ratios=[1, 1.1]))
    for c, atm in enumerate(("N", "O")):
        for m in ORDER:
            for r_, key in enumerate(("rem", "dtg")):
                mean, lo, hi = band(runs, m, atm, key)
                ax = axs[r_, c]
                ax.fill_between(GRID, lo, hi, color=COLOR[m], alpha=0.16, lw=0)
                ax.plot(GRID, mean, color=COLOR[m], lw=2.0, label=LABEL[m] if (r_ == 0 and c == 0) else None)
        axs[0, c].set_title(f"{'(a)' if c == 0 else '(b)'} 질량 잔류 — {ATM[atm]}")
        axs[1, c].set_title(f"{'(c)' if c == 0 else '(d)'} 질량감소속도 DTG — {ATM[atm]}")
        axs[0, c].set_ylim(-5, 112)
        axs[1, c].set_ylim(-0.1, 2.95)
        axs[1, c].set_xlabel("온도 (°C, 시료 온도)")
    axs[0, 0].set_ylabel("잔류 질량 100·(1-α) (%)")
    axs[1, 0].set_ylabel("dα/dT (%/°C)")
    for ax in axs.ravel():
        ax.set_xlim(60, 790)
    # main peaks of every material and atmosphere (stage with the highest peak), labelled with temperature and rate
    # crowded peaks get their label in free space with a leader line (data units)
    placed = {("PIR_C", "N"): (175, 1.02), ("PIR_SC", "N"): (175, 0.80), ("PIR_C", "O"): (610, 1.10),
              ("PIR_SC", "O"): (610, 0.90), ("PF_out", "O"): (610, 0.70)}
    for (m, atm), g in stages.groupby(["material", "atm"]):
        s = g.sort_values("peak_rate_pct_per_C").iloc[-1]
        if not s["finished"]:
            continue
        ax = axs[1, 0 if atm == "N" else 1]
        at = placed.get((m, atm))
        peak_label(ax, s["REFERENCE_TEMPERATURE_C"], s["peak_rate_pct_per_C"],
                   f"{s['REFERENCE_TEMPERATURE_C']:.0f} °C · {s['peak_rate_pct_per_C']:.2f}", COLOR[m],
                   ha="left" if at else "center", at=at)
    notes_n = [(95, 42, "EPS: 한 단계 급격 분해\n(가연성 397 °C, 준불연 426 °C)", "left"),
               (95, 22, "PIR(PUR): 300 °C대 1단계 뒤\n탄화층이 천천히 분해", "left"),
               (500, 100, "PF: 완만한 탄화 — 790 °C에서도 진행", "left")]
    notes_o = [(110, 30, "EPS: 공기에서 주분해가\n80–100 °C 앞당겨짐", "left"),
               (560, 78, "PF 내부용: 472 °C 급속 전소\nPF 외부용: 490 °C 완만", "left"),
               (560, 45, "PIR(PUR): 525–534 °C\n탄화층 산화", "left")]
    for ax, notes in ((axs[0, 0], notes_n), (axs[0, 1], notes_o)):
        for x, y, t, ha in notes:
            ax.text(x, y, t, fontsize=NOTE_FONT, color=tga.INK2, ha=ha, va="center", linespacing=1.35)
    axs[1, 1].text(0.99, 0.97, "점 = 주 피크 (온도 · 최대속도 %/°C)\n띠 = 두 반복의 범위", transform=axs[1, 1].transAxes,
                   ha="right", va="top", fontsize=NOTE_FONT, color=tga.INK2)
    axs[1, 0].text(0.99, 0.97, "PF 2종은 N$_2$에서 뚜렷한 주 피크 없이\n790 °C까지 감소가 이어짐(표시 생략)", transform=axs[1, 0].transAxes,
                   ha="right", va="top", fontsize=NOTE_FONT, color=tga.INK2)
    fig.legend(loc="upper center", ncol=6, bbox_to_anchor=(0.5, 1.0), frameon=False, fontsize=10, handlelength=2.4)
    fig.suptitle("단열재 6종 TGA — 질량 잔류와 질량감소속도 (5 °C/min, 두 반복 평균)", y=1.045, fontsize=13,
                 fontweight="bold", color=tga.INK, x=0.06, ha="left")
    fig.tight_layout(rect=(0, 0, 1, 0.965))
    fig.savefig(path, dpi=160, bbox_inches="tight")
    plt.close(fig)


def fig_overview_dsc(runs, stages, path):
    fig, axs = plt.subplots(1, 2, figsize=(15, 8.6), sharey=True)
    sel = (GRID >= DSC_WIN[0]) & (GRID <= DSC_WIN[1])
    step = 4.5
    for c, atm in enumerate(("N", "O")):
        ax = axs[c]
        for i, m in enumerate(ORDER):
            base = -i * step
            sc = DSC_SCALE.get((m, atm), 1.0)
            ax.axhline(base, color=tga.GRID, lw=0.8, zorder=0)
            for rep, ls in ((1, "-"), (2, (0, (4, 2)))):
                ax.plot(GRID[sel], base + sc * runs[(m, atm, rep)]["h"][sel], color=COLOR[m], lw=1.5 if rep == 1 else 1.2, ls=ls)
            name = LABEL[m] + (" (×1/5)" if sc != 1.0 else "")
            ax.text(DSC_WIN[0] + 4, base + 1.1, name, fontsize=9.5, color=tga.INK, fontweight="bold", va="bottom")
            ev = DSC_EVENTS.get((m, atm))
            if ev:
                t, txt = ev
                y = base + sc * np.interp(t, GRID, runs[(m, atm, 2)]["h"])
                at = (225, base - 2.6) if y < base - 1 else (t + 40, base + 1.8)
                ax.annotate(txt, (t, y), xytext=at, textcoords="data", fontsize=NOTE_FONT, color=tga.INK, va="center",
                            arrowprops=dict(arrowstyle="-", color=tga.INK2, lw=0.8, shrinkA=2, shrinkB=2))
            else:
                ax.text(DSC_WIN[0] + 4, base - 0.45, "재현 사건 없음", fontsize=8.5, color=tga.MUTED, ha="left",
                        va="top")
            # main DTG peak of this material and atmosphere, so the heat flow can be read against the mass loss
            g = stages[(stages.material == m) & (stages.atm == atm) & stages.finished]
            if len(g):
                tp = g.sort_values("peak_rate_pct_per_C").iloc[-1]["REFERENCE_TEMPERATURE_C"]
                ax.plot([tp, tp], [base - 1.4, base + 0.9], color=COLOR[m], lw=1.0, ls=(0, (1, 1.5)), zorder=1)
        ax.set_title(f"{'(a)' if c == 0 else '(b)'} {ATM[atm]} — 흡열 ↑ · 발열 ↓")
        ax.set_xlim(DSC_WIN[0] - 5, DSC_WIN[1] + 5)
        ax.set_xlabel("온도 (°C, 시료 온도)")
        ax.set_yticks([])
        ax.spines["left"].set_visible(False)
        # scale bar instead of a y axis (traces are offset)
        x0 = DSC_WIN[1] - 18
        ax.plot([x0, x0], [-len(ORDER) * step + 0.6, -len(ORDER) * step + 2.6], color=tga.INK, lw=1.6)
        ax.text(x0 - 6, -len(ORDER) * step + 1.6, "2 mW", ha="right", va="center", fontsize=9, color=tga.INK)
    axs[0].set_ylim(-len(ORDER) * step + 0.2, 3.2)
    axs[1].text(1.0, 1.015, "실선 #1 · 점선 #2 · 세로 점선 = DTG 주 피크", transform=axs[1].transAxes, fontsize=9,
                color=tga.INK2, ha="right", va="bottom")
    fig.suptitle("단열재 6종 DSC — 국소 열류 (미보정 열류에서 100 °C 창 추세를 뺀 값, 정성)", y=1.0, fontsize=13,
                 fontweight="bold", color=tga.INK, x=0.06, ha="left")
    fig.tight_layout(rect=(0, 0, 1, 0.97))
    fig.savefig(path, dpi=160, bbox_inches="tight")
    plt.close(fig)


def sheet(runs, stages, summ, verdict, m, path):
    fig = plt.figure(figsize=(15, 9.6))
    gs = fig.add_gridspec(3, 2, width_ratios=[2.15, 1], height_ratios=[1, 1, 0.9], wspace=0.06, hspace=0.16,
                          top=0.92, bottom=0.07)
    axs = [fig.add_subplot(gs[i, 0]) for i in range(3)]
    for ax in axs[1:]:
        ax.sharex(axs[0])
    colors = {"N": N2_COLOR, "O": COLOR[m]}
    for atm in ("N", "O"):
        for ax, key in zip(axs[:2], ("rem", "dtg")):
            mean, lo, hi = band(runs, m, atm, key)
            ax.fill_between(GRID, lo, hi, color=colors[atm], alpha=0.16, lw=0)
            ax.plot(GRID, mean, color=colors[atm], lw=2.0, label=ATM[atm] if key == "rem" else None)
        sel = (GRID >= DSC_WIN[0]) & (GRID <= DSC_WIN[1])
        for rep, ls in ((1, "-"), (2, (0, (4, 2)))):
            axs[2].plot(GRID[sel], runs[(m, atm, rep)]["h"][sel], color=colors[atm], lw=1.4, ls=ls,
                        label=f"{'N$_2$' if atm == 'N' else '공기(O)'} #{rep}")
        for _, s in stages[(stages.material == m) & (stages.atm == atm) & stages.finished].iterrows():
            if s["peak_rate_pct_per_C"] < 0.08:
                continue
            peak_label(axs[1], s["REFERENCE_TEMPERATURE_C"], s["peak_rate_pct_per_C"],
                       f"{s['REFERENCE_TEMPERATURE_C']:.0f} °C", colors[atm], dx=-6 if atm == "N" else 6, dy=5,
                       ha="right" if atm == "N" else "left")
        ev = DSC_EVENTS.get((m, atm))
        if ev:
            axs[2].annotate(ev[1], (ev[0], np.interp(ev[0], GRID, runs[(m, atm, 2)]["h"])), xytext=(26, 0),
                            textcoords="offset points", fontsize=NOTE_FONT, color=tga.INK, va="center",
                            arrowprops=dict(arrowstyle="-", color=tga.INK2, lw=0.8))
    axs[0].set_ylim(-5, 112)
    dmax = max(band(runs, m, a, "dtg")[2].max() for a in ("N", "O"))
    axs[1].set_ylim(-0.05 * dmax, dmax * 1.18)
    hmin = min(runs[(m, a, k)]["h"][(GRID >= DSC_WIN[0]) & (GRID <= DSC_WIN[1])].min() for a in ("N", "O") for k in (1, 2))
    hmax = max(runs[(m, a, k)]["h"][(GRID >= DSC_WIN[0]) & (GRID <= DSC_WIN[1])].max() for a in ("N", "O") for k in (1, 2))
    pad = 0.12 * (hmax - hmin)
    axs[2].set_ylim(hmin - pad, hmax + pad)
    axs[2].axhline(0, color=tga.AXIS, lw=0.8)
    axs[0].set_ylabel("잔류 질량 (%)")
    axs[1].set_ylabel("dα/dT (%/°C)")
    axs[2].set_ylabel("국소 열류 (mW, 흡열 +)")
    axs[2].set_xlabel("온도 (°C, 시료 온도)")
    axs[0].set_title(f"(a) TGA 질량 잔류", loc="left")
    axs[1].set_title("(b) DTG — 점 = 단계별 피크", loc="left")
    axs[2].set_title("(c) DSC 국소 열류 (정성) — 실선 #1, 점선 #2", loc="left")
    axs[0].legend(loc="lower left")
    axs[2].legend(loc="lower left", ncol=4, fontsize=8.5)
    for ax in axs[:2]:
        plt.setp(ax.get_xticklabels(), visible=False)
    axs[0].set_xlim(60, 790)

    # numbers panel
    tx = fig.add_subplot(gs[:, 1])
    tx.axis("off")
    lines = []

    def row(atm):
        s = summ[(summ.material == m) & (summ.atmosphere == ("N2" if atm == "N" else "oxidative (O)"))].iloc[0]
        v = verdict[(verdict.material == m) & (verdict.atmosphere == ("N2" if atm == "N" else "oxidative (O)"))].iloc[0]
        return (f"  T$_{{onset}}$ {s.T_onset_C_mean:.0f} °C · T$_{{peak}}$ {s.T_peak_C_mean:.0f} °C\n"
                f"  r$_{{max}}$ {s.r_max_pct_C_mean:.2f} %/°C · T$_{{\\alpha 50}}$ {s.T_a50_C_mean:.0f} °C\n"
                f"  f$_{{400}}$ {s.f_above400_mean:.2f} · 자료 판정: {VERDICT_KO[v.overall]}")
    lines.append(("열분해 (N$_2$)", row("N")))
    lines.append(("산화 (공기)", row("O")))
    dsc = [t for (mm, a), (_, t) in DSC_EVENTS.items() if mm == m]
    lines.append(("DSC", "  " + (dsc[0] if dsc else "두 반복에서 재현된 사건 없음 (정성 참고)")))
    rx = []
    for _, s in stages[(stages.material == m)].iterrows():
        if not str(s["role"]).startswith("반응"):
            continue
        rx.append(f"  {s['role'].split(':')[0]} ({'N$_2$' if s['atm'] == 'N' else '공기'}): T$_{{ref}}$ {s['REFERENCE_TEMPERATURE_C']:.0f} °C,\n"
                  f"    REFERENCE_RATE {s['REFERENCE_RATE_per_s']:.2e} 1/s")
    lines.append(("PyroSim 반응 시작값 (HEATING_RATE 5)", "\n".join(rx) if rx else "  —"))
    y = 0.98
    tx.text(0.02, y, LABEL[m], fontsize=16, fontweight="bold", color=tga.INK, va="top", transform=tx.transAxes)
    y -= 0.075
    for head, body in lines:
        tx.text(0.02, y, head, fontsize=11, fontweight="bold", color=tga.INK, va="top", transform=tx.transAxes)
        y -= 0.04
        tx.text(0.02, y, body, fontsize=9.8, color=tga.INK2, va="top", transform=tx.transAxes, linespacing=1.5)
        y -= 0.034 * (body.count("\n") + 1) + 0.045
    tx.text(0.02, 0.02, "밀도 · 비열 · 열전도율 · 반응열은 이 시험으로 구할 수 없음\n(실측 또는 제품 시험성적서 — PyroSim 입력 자료 참고)",
            fontsize=9, color=tga.MUTED, va="bottom", transform=tx.transAxes, linespacing=1.4)
    fig.suptitle(f"{LABEL[m]} — TGA · DSC 재료 특성 (5 °C/min, N$_2$ 회색 · 공기(O) 색)", x=0.06, ha="left", y=0.995,
                 fontsize=13, fontweight="bold", color=tga.INK)
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)


VERDICT_KO = {"reliable": "신뢰", "conditional": "조건부", "endpoint": "종점 의존", "remeasure": "재측정"}


def main():
    tga.style()
    plt.rcParams.update({"font.family": korean_font(), "axes.unicode_minus": False})
    runs = load()
    stages = pd.read_csv(HERE / "results" / "pyrosim_reactions.csv")
    stages["atm"] = stages["atmosphere"].map({"N₂": "N", "공기(O)": "O"})
    summ = pd.read_csv(HERE / "results" / "tga_summary.csv")
    verdict = pd.read_csv(HERE / "results" / "qa_verdict.csv")
    out = HERE / "figures"
    fig_overview_tga(runs, stages, out / "material_tga.png")
    fig_overview_dsc(runs, stages, out / "material_dsc.png")
    for m in ORDER:
        sheet(runs, stages, summ, verdict, m, out / f"material_{m}.png")
    print("figures:", ", ".join(p.name for p in sorted(out.glob("material_*.png"))))


if __name__ == "__main__":
    main()
