#!/usr/bin/env python3
"""QA: is '(2) EPS_C' a 30-min record? Height values are max of the three thermocouples per height.

From 263 s every row of '(2) EPS_C' equals the same row of '(1) PF_semi_out' (all 18 channels, including the
CH14 'BURNOUT' cells at 418-453 s), so only 0-262 s is EPS combustible test #2.

    python qa_eps_c2.py data/EPS.xlsx data/PF.xlsx
"""
from __future__ import annotations

import argparse
import importlib.util
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402

HERE = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location("cavity_analysis", HERE / "analysis.py")
cfa = importlib.util.module_from_spec(spec)
spec.loader.exec_module(cfa)

SPLICE_ROW = 263


def height_max(ch):
    """Max of the three thermocouples per height (°C); a height with all three blank stays blank."""
    with np.errstate(all="ignore"):
        return np.column_stack([np.nanmax(ch[:, 3 * i:3 * i + 3], axis=1) for i in range(6)])


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("eps", type=Path, nargs="?", default=HERE / "data" / "EPS.xlsx")
    ap.add_argument("pf", type=Path, nargs="?", default=HERE / "data" / "PF.xlsx")
    args = ap.parse_args()
    tests = cfa.read_tests([args.eps, args.pf])
    E, P, E1 = tests["EPS_C-2"]["ch"], tests["PF_out-1"]["ch"], tests["EPS_C-1"]["ch"]
    n = min(len(E), len(P))
    same = (np.isclose(E[:n], P[:n], atol=0.05) | (~np.isfinite(E[:n]) & ~np.isfinite(P[:n]))).all(axis=1)
    HE, HP, H1 = height_max(E[:n]), height_max(P[:n]), height_max(E1)
    print(f"identical rows from {SPLICE_ROW} s: {int(same[SPLICE_ROW:].sum())}/{n - SPLICE_ROW}; "
          f"before: {int(same[:SPLICE_ROW].sum())}/{SPLICE_ROW}")
    print(f"500 mm (max CH1-3) at 262 s / 263 s: EPS_C-2 {HE[262, 0]:.0f} / {HE[263, 0]:.0f} °C, "
          f"PF_out-1 {HP[262, 0]:.0f} / {HP[263, 0]:.0f} °C")

    cfa.style()
    t = np.arange(n)
    eps_c, pf_c = cfa.MATERIALS["EPS_C"][2], cfa.MATERIALS["PF_out"][2]
    fig, axs = plt.subplots(1, 2, figsize=(14, 5.4), gridspec_kw=dict(width_ratios=[1.35, 1]))
    ax = axs[0]
    ax.plot(t, HP[:, 0], color=pf_c, lw=3.2, alpha=0.45, label="'(1) PF_semi_out' (PF exterior-use #1)")
    ax.plot(t, HE[:, 0], color=eps_c, lw=1.1, label="'(2) EPS_C' sheet as received")
    ax.plot(np.arange(len(H1)), H1[:, 0], color=cfa.INK2, lw=1.0, ls=(0, (4, 2)), label="'(1) EPS_C' (240 s record)")
    ax.axvline(SPLICE_ROW, color=cfa.INK2, lw=0.8, ls=(0, (2, 2)))
    ax.text(SPLICE_ROW + 20, 1210, f"from {SPLICE_ROW} s: {int(same[SPLICE_ROW:].sum())} of {n - SPLICE_ROW} rows identical\n"
            "in all 18 channels, incl. the CH14 'BURNOUT'\ncells at 418–453 s", fontsize=8.5, color=cfa.INK, va="top")
    ax.set_xlim(0, 1800)
    ax.set_ylim(0, 1250)
    ax.set_xlabel("Time (s)")
    ax.set_ylabel("500 mm temperature, max of CH1–CH3 (°C)")
    ax.set_title("(a) 500 mm: the EPS sheet follows the PF sheet exactly from 263 s")
    ax.legend(loc="lower right")
    ax = axs[1]
    for i, (ls, z) in enumerate(((("-"), 500), ((0, (5, 2)), 1000))):
        ax.plot(t, HP[:, i], color=pf_c, lw=3.2, alpha=0.45, ls=ls)
        ax.plot(t, HE[:, i], color=eps_c, lw=1.4, ls=ls, label=f"{z} mm")
    ax.axvline(SPLICE_ROW, color=cfa.INK2, lw=0.8, ls=(0, (2, 2)))
    ax.annotate(f"262 → 263 s: {HE[262, 0]:.0f} → {HE[263, 0]:.0f} °C\nin one second at every height",
                (SPLICE_ROW, HE[263, 0]), xytext=(-14, -40), textcoords="offset points", ha="right", fontsize=8.5,
                color=cfa.INK, arrowprops=dict(arrowstyle="-", color=cfa.INK2, lw=0.8))
    ax.set_xlim(200, 320)
    ax.set_ylim(0, 900)
    ax.set_xlabel("Time (s)")
    ax.set_ylabel("Max of the three thermocouples (°C)")
    ax.set_title("(b) Zoom: amber = '(2) EPS_C', wide blue = '(1) PF_semi_out'")
    ax.legend(loc="center left", title="height", title_fontsize=8.5)
    fig.tight_layout()
    out = HERE / "figures" / "qa_eps_c2_splice.png"
    fig.savefig(out, dpi=150)
    plt.close(fig)
    print("figure:", out)


if __name__ == "__main__":
    main()
