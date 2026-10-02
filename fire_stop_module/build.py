#!/usr/bin/env python3
"""샌드위치패널 외벽 화재확산방지 모듈 – 상세도 생성.

    python build.py

- figures/d0*.svg : 단독 SVG (README용, 라이트/다크 자동 전환)
- index.html      : template.html 의 <!--D0n--> 자리에 같은 도면을 인라인으로 넣은 페이지

단위 mm. 얇은 강판과 틈은 보이도록 과장했고(NTS), 깊이 방향과 높이 방향의 축척도 다르다.
좌표: x = 기존 외부강판 바깥면에서 실내 쪽으로(+), y = 하부 절단선(EL±0)에서 위로(+).
"""
from __future__ import annotations

import math
from pathlib import Path

HERE = Path(__file__).resolve().parent
FIG_DIR = HERE / "figures"

# ------------------------------------------------------------------ style
LIGHT = {
    "sheet": "#ffffff", "ink": "#1d2329", "muted": "#5f6973", "rule": "#d6dbe1",
    "new": "#c4511a", "new-tint": "#fbe6da", "water": "#1769c2",
    "bad": "#b0124a", "ok": "#2f7d4f", "core": "#ece6d3", "core-dot": "#b3a77f",
}
DARK = {
    "sheet": "#1b1f25", "ink": "#e4e8ec", "muted": "#9aa4ae", "rule": "#333a43",
    "new": "#f08043", "new-tint": "#46281a", "water": "#5aa9f6",
    "bad": "#f0679a", "ok": "#62c48e", "core": "#343128", "core-dot": "#71684f",
}
BODY = "var(--f-body,'IBM Plex Sans KR','Noto Sans KR','Apple SD Gothic Neo','Malgun Gothic',sans-serif)"
MONO = "var(--f-mono,'IBM Plex Mono',ui-monospace,Menlo,monospace)"

# class rules shared by the page (index.html) and the standalone SVG files
RULES = f"""
.dwg{{font-family:{BODY}}}
.dwg .bg{{fill:var(--sheet)}}
.dwg .fl{{stroke:none}}
.dwg .sk{{fill:none;stroke:var(--ink);stroke-width:2.2;stroke-linecap:square}}
.dwg .gt{{fill:none;stroke:var(--muted);stroke-width:3}}
.dwg .nw{{fill:none;stroke:var(--new);stroke-width:2.4;stroke-linejoin:round;stroke-linecap:round}}
.dwg .nw2{{fill:none;stroke:var(--new);stroke-width:1.7;stroke-linejoin:round;stroke-linecap:round}}
.dwg .pc-bg{{fill:var(--core)}}
.dwg .pc-dot{{fill:var(--core-dot)}}
.dwg .pm-bg{{fill:var(--new-tint)}}
.dwg .pm-ln{{fill:none;stroke:var(--new);stroke-width:.7;opacity:.7}}
.dwg .ps-bg{{fill:var(--sheet)}}
.dwg .ps-ln{{fill:none;stroke:var(--new);stroke-width:.6;opacity:.55}}
.dwg .void{{fill:var(--sheet);stroke:var(--muted);stroke-width:.8;stroke-dasharray:3 2}}
.dwg .brk{{fill:none;stroke:var(--muted);stroke-width:1}}
.dwg .wt{{fill:none;stroke:var(--water);stroke-width:1.6;stroke-dasharray:5 3}}
.dwg .wt2{{fill:none;stroke:var(--water);stroke-width:1.2;stroke-dasharray:3 2}}
.dwg .lk{{fill:none;stroke:var(--bad);stroke-width:1.6;stroke-dasharray:5 3}}
.dwg .mv{{fill:none;stroke:var(--ink);stroke-width:1.5}}
.dwg .mk-w{{fill:var(--water)}}
.dwg .mk-k{{fill:var(--ink)}}
.dwg .mk-b{{fill:var(--bad)}}
.dwg .dm{{fill:none;stroke:var(--muted);stroke-width:.8}}
.dwg .dx{{fill:var(--muted);font-size:11px;font-variant-numeric:tabular-nums}}
.dwg .tx{{fill:var(--ink);font-size:12px}}
.dwg .ts{{fill:var(--muted);font-size:11px}}
.dwg .th{{fill:var(--ink);font-size:13.5px;font-weight:700}}
.dwg .tw{{fill:var(--water);font-size:11.5px}}
.dwg .tn{{fill:var(--new);font-size:12px;font-weight:600}}
.dwg .tb{{fill:var(--bad);font-size:12px;font-weight:600}}
.dwg .tk{{fill:var(--ok);font-size:12px;font-weight:600}}
.dwg .halo{{paint-order:stroke;stroke:var(--sheet);stroke-width:3.5px;stroke-linejoin:round}}
.dwg .bc{{fill:var(--sheet);stroke:var(--new);stroke-width:1.2}}
.dwg .bn{{fill:var(--new);font-family:{MONO};font-size:10.5px;font-weight:600}}
.dwg .ld{{fill:none;stroke:var(--muted);stroke-width:.7}}
.dwg .ldd{{fill:var(--muted)}}
.dwg .sc{{fill:none;stroke:var(--ink);stroke-width:1.8}}
.dwg .sl{{fill:var(--bad)}}
.dwg .hl{{fill:none;stroke:var(--new);stroke-width:1.1;stroke-dasharray:2.5 2}}
.dwg .jt{{fill:none;stroke:var(--muted);stroke-width:.9;stroke-dasharray:3 2}}
"""


def token_css(selector: str) -> str:
    light = ";".join(f"--{k}:{v}" for k, v in LIGHT.items())
    dark = ";".join(f"--{k}:{v}" for k, v in DARK.items())
    return f"{selector}{{{light}}}@media (prefers-color-scheme:dark){{{selector}{{{dark}}}}}"


# ------------------------------------------------------------------ svg helpers
def esc(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


class Frame:
    """mm → px (anisotropic)."""

    def __init__(self, ox, oy, sx, sy):
        self.ox, self.oy, self.sx, self.sy = ox, oy, sx, sy

    def X(self, x):
        return self.ox + x * self.sx

    def Y(self, y):
        return self.oy - y * self.sy

    def __call__(self, pts):
        return [(self.X(x), self.Y(y)) for x, y in pts]


class BrokenFrame(Frame):
    """Frame whose middle band (y_lo..y_hi) is squeezed into `gap` px, so both lap joints draw large."""

    def __init__(self, ox, top_px, sx, sy, ytop, y_hi, y_lo, gap):
        super().__init__(ox, 0, sx, sy)
        self.y_hi, self.y_lo, self.gap = y_hi, y_lo, gap
        self.Yh = top_px + (ytop - y_hi) * sy

    def Y(self, y):
        if y >= self.y_hi:
            return self.Yh - (y - self.y_hi) * self.sy
        if y <= self.y_lo:
            return self.Yh + self.gap + (self.y_lo - y) * self.sy
        return self.Yh + self.gap * (self.y_hi - y) / (self.y_hi - self.y_lo)

    def cover(self, f, x0=-20.0, x1=112.0):
        """blank the squeezed band and mark it with two break lines."""
        a, b = self.X(x0), self.X(x1)
        f.rect(a, self.Yh, b - a, self.gap, "bg")
        m = (a + b) / 2
        for yy in (self.Yh + 2, self.Yh + self.gap - 2):
            f.poly([(a, yy), (m - 7, yy), (m - 3, yy - 6), (m + 3, yy + 6), (m + 7, yy), (b, yy)], "brk")


class Fig:
    def __init__(self, fid, w, h, aria):
        self.fid, self.w, self.h, self.aria = fid, w, h, aria
        self.out: list[str] = []

    def ref(self, name):
        return f"url(#{self.fid}-{name})"

    def poly(self, pts, cls, closed=False, fill=None, end=None):
        tag = "polygon" if closed else "polyline"
        extra = f' fill="{self.ref(fill)}"' if fill else ""
        if end:
            extra += f' marker-end="{self.ref(end)}"'
        p = " ".join(f"{x:.1f},{y:.1f}" for x, y in pts)
        self.out.append(f'<{tag} class="{cls}" points="{p}"{extra}/>')

    def path(self, d, cls, end=None):
        extra = f' marker-end="{self.ref(end)}"' if end else ""
        self.out.append(f'<path class="{cls}" d="{d}"{extra}/>')

    def line(self, x1, y1, x2, y2, cls, end=None):
        self.poly([(x1, y1), (x2, y2)], cls, end=end)

    def rect(self, x, y, w, h, cls, fill=None):
        extra = f' fill="{self.ref(fill)}"' if fill else ""
        self.out.append(f'<rect class="{cls}" x="{x:.1f}" y="{y:.1f}" width="{w:.1f}" height="{h:.1f}"{extra}/>')

    def circle(self, x, y, r, cls):
        self.out.append(f'<circle class="{cls}" cx="{x:.1f}" cy="{y:.1f}" r="{r:.1f}"/>')

    def text(self, x, y, s, cls="tx", anchor="start", rot=None):
        if rot is None:
            self.out.append(f'<text class="{cls}" x="{x:.1f}" y="{y:.1f}" text-anchor="{anchor}">{esc(s)}</text>')
        else:
            self.out.append(f'<text class="{cls}" text-anchor="{anchor}" '
                            f'transform="translate({x:.1f} {y:.1f}) rotate({rot})">{esc(s)}</text>')

    def render(self, standalone=False) -> str:
        i = self.fid
        defs = (
            "<defs>"
            f'<pattern id="{i}-core" width="7" height="7" patternUnits="userSpaceOnUse">'
            '<rect class="pc-bg" width="7" height="7"/><circle class="pc-dot" cx="1.75" cy="1.75" r=".8"/>'
            '<circle class="pc-dot" cx="5.25" cy="5.25" r=".8"/></pattern>'
            f'<pattern id="{i}-mw" width="10" height="7" patternUnits="userSpaceOnUse">'
            '<rect class="pm-bg" width="10" height="7"/><polyline class="pm-ln" points="0,3.5 2.5,1 5,3.5 7.5,6 10,3.5"/></pattern>'
            f'<pattern id="{i}-ms" width="12" height="9" patternUnits="userSpaceOnUse">'
            '<rect class="ps-bg" width="12" height="9"/><polyline class="ps-ln" points="0,4.5 3,1.5 6,4.5 9,7.5 12,4.5"/></pattern>'
            + "".join(
                f'<marker id="{i}-{n}" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="6" markerHeight="6" '
                f'orient="auto-start-reverse"><path class="{c}" d="M0,0 L10,5 L0,10 z"/></marker>'
                for n, c in (("aw", "mk-w"), ("ak", "mk-k"), ("ab", "mk-b"))
            )
            + "</defs>"
        )
        size = f' width="{self.w}" height="{self.h}"' if standalone else ""
        head = (f'<svg class="dwg" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {self.w} {self.h}"{size} '
                f'role="img" aria-label="{esc(self.aria)}">')
        style = f"<style>{token_css('svg.dwg')}{RULES}</style>" if standalone else ""
        bg = f'<rect class="bg" width="{self.w}" height="{self.h}"/>'
        return head + style + defs + bg + "".join(self.out) + "</svg>"


def dim_v(f: Fig, X, Y1, Y2, label, side=-1):
    """vertical dimension line at px X between px Y1..Y2, text rotated along it."""
    f.line(X, Y1, X, Y2, "dm")
    for yy in (Y1, Y2):
        f.line(X - 3.5, yy + 3.5, X + 3.5, yy - 3.5, "dm")
    tx = X - 4 if side < 0 else X + 12
    f.text(tx, (Y1 + Y2) / 2, label, "dx", "middle", rot=-90)


def dim_h(f: Fig, Y, X1, X2, label):
    f.line(X1, Y, X2, Y, "dm")
    for xx in (X1, X2):
        f.line(xx - 3.5, Y + 3.5, xx + 3.5, Y - 3.5, "dm")
    f.text((X1 + X2) / 2, Y - 5, label, "dx", "middle")


def ext(f: Fig, x1, y1, x2, y2):
    f.line(x1, y1, x2, y2, "dm")


def callout(f: Fig, target, bx, by, n, label, side):
    tx, ty = target
    edge = bx + 9 if side == "L" else bx - 9
    f.line(tx, ty, edge, by, "ld")
    f.circle(tx, ty, 1.8, "ldd")
    f.circle(bx, by, 9, "bc")
    f.text(bx, by + 3.6, str(n), "bn", "middle")
    if side == "L":
        f.text(bx - 14, by + 4, label, "tx", "end")
    else:
        f.text(bx + 14, by + 4, label, "tx", "start")


def brk(f: Fig, fr: Frame, y, x0=-9.0, x1=109.0):
    """horizontal break line across the wall section at elevation y."""
    m = (x0 + x1) / 2
    f.poly(fr([(x0, y), (m - 6, y), (m - 3, y + 12), (m + 3, y - 12), (m + 6, y), (x1, y)]), "brk")


# ------------------------------------------------------------------ geometry (mm)
T = 100.0        # example core thickness (t); the system works for t = 50–200
H = 460.0        # cut opening: outer skin + core removed between EL±0 and EL+460
LAP = 70.0       # tongue behind the upper skin
REVEAL = 15.0    # joggle top sits this far below the upper skin edge (room for the 12 mm lift)
POCKET_D, POCKET_H = 10.0, 100.0   # insulation removed behind the upper skin only
LIFT = 12.0      # cassette lift during install; both hooks engage 8 mm
TOP_PLATE = H - REVEAL - 15.0      # 430: joggle bottom = top plate level
BOT_PLATE = 12.0                   # bottom plate rests on the starter-rail rib (12 high)

STARTER = [(98, 50), (98, 1.5), (24, 1.5), (24, BOT_PLATE), (19, BOT_PLATE), (19, 1.5),
           (-2, 1.5), (-2, -50), (-7, -57)]
STARTER_PLAIN = [(98, 50), (98, 1.5), (-2, 1.5), (-2, -50), (-7, -57)]
TOPRAIL = [(98.5, TOP_PLATE), (98.5, TOP_PLATE - 33), (88, TOP_PLATE - 33), (88, TOP_PLATE - 10)]


def module(lift=0.0, rot_deg=0.0):
    """cassette outline points; lifted by `lift` and tilted bottom-out about the tongue tip."""
    tip = H + LAP
    pivot = (4.5, tip + lift)
    rot = math.radians(rot_deg)

    def tf(pts):
        out = []
        for x, y in pts:
            y += lift
            if rot:
                rx, ry = x - pivot[0], y - pivot[1]
                c, s = math.cos(rot), math.sin(rot)
                x, y = pivot[0] + rx * c + ry * s, pivot[1] - rx * s + ry * c
            out.append((x, y))
        return out

    tp, bp = TOP_PLATE, BOT_PLATE
    jtop = H - REVEAL
    sb = tp + lift          # the strip squeezes against the upper core (H) when lifted
    return {
        "front": tf([(8.5, tip - 10), (8.5, tip), (4.5, tip), (4.5, jtop), (-4.5, tp), (-4.5, -40), (-10, -47)]),
        "top": tf([(-2.5, tp - 20), (-2.5, tp), (94, tp), (94, tp - 18)]),
        "bot": tf([(-2.5, bp + 20), (-2.5, bp), (26, bp), (26, bp - 8), (30, bp - 8), (30, bp), (88, bp), (88, bp + 10)]),
        "wool": tf([(-4, bp), (97, bp), (97, 54), (99.5, 54), (99.5, tp - 39), (86, tp - 39), (86, tp), (-4, tp)]),
        "strip": [(-3, sb), (96, sb), (96, H), (5.5, H), (5.5, min(H, jtop + lift))],
    }


def draw_existing(f: Fig, fr: Frame, ytop, ybot, pocket=True):
    ptop = min(H + POCKET_H, ytop)
    up = [(0, ytop), (T, ytop), (T, H)]
    up += [(POCKET_D, H), (POCKET_D, ptop), (0, ptop)] if pocket else [(0, H)]
    f.poly(fr(up), "fl", closed=True, fill="core")
    f.poly(fr([(0, 0), (T, 0), (T, ybot), (0, ybot)]), "fl", closed=True, fill="core")
    if pocket:
        f.poly(fr([(0, H), (POCKET_D, H), (POCKET_D, ptop), (0, ptop)]), "void", closed=True)


def draw_skins(f: Fig, fr: Frame, ytop, ybot):
    f.poly(fr([(0, H), (0, ytop)]), "sk")
    f.poly(fr([(0, ybot), (0, 0)]), "sk")
    f.poly(fr([(T, ybot), (T, ytop)]), "sk")
    brk(f, fr, ytop)
    brk(f, fr, ybot)


def draw_module(f: Fig, fr: Frame, m, fills=True):
    if fills:
        f.poly(fr(m["wool"]), "fl", closed=True, fill="mw")
        f.poly(fr(m["strip"]), "fl", closed=True, fill="ms")
    f.poly(fr(m["top"]), "nw2")
    f.poly(fr(m["bot"]), "nw2")
    f.poly(fr(m["front"]), "nw")


def screw(f: Fig, fr: Frame, y, x_head, x_tip):
    hx, ty = fr.X(x_head), fr.Y(y)
    f.line(hx, ty - 5, hx, ty + 5, "sc")
    f.line(hx, ty, fr.X(x_tip), ty, "sc")


def legend(f: Fig, x, y, cols=1):
    rows = [
        ("line", "wt", "빗물 경로"), ("line", "nw", "신설 부재"),
        ("line", "sk", "기존 부재"), ("fill", "core", "기존 단열재(EPS·PIR)"),
        ("fill", "mw", "미네랄울(불연)"), ("fill", "ms", "압축 미네랄울 스트립"),
    ]
    for k, (kind, c, label) in enumerate(rows):
        cx = x + (k % cols) * 150
        cy = y + (k // cols) * 18
        if kind == "line":
            f.line(cx, cy - 4, cx + 24, cy - 4, c)
        else:
            f.rect(cx, cy - 10, 24, 11, "fl", fill=c)
        f.text(cx + 30, cy, label, "ts")


# ------------------------------------------------------------------ D-01 principle
def d01() -> Fig:
    f = Fig("d01", 920, 410,
            "원리 비교: 모듈을 기존 외부강판 바깥에 덧대면 상부 겹침이 물 흐름과 반대가 되어 실리콘이 유일한 방수선이 되고, "
            "상부는 기존 강판 뒤로 넣고 하부는 기존 강판 위로 덮으면 모든 겹침이 위가 아래를 덮어 실링 없이 배수된다")
    YT, YB = 548, -70
    titles = [("기존 방식 · 바깥에 덧댐", "상부 겹침이 물 흐름과 반대 → 실리콘이 유일한 방수선", "tb"),
              ("제안 · 위는 안쪽, 아래는 바깥", "위가 아래를 덮는 겹침 → 실링 없이 중력 배수", "tk")]
    for i, (title, verdict, vcls) in enumerate(titles):
        fr = BrokenFrame(235 + i * 465, 72, 1.6, 1.0, YT, 395, 60, 30)
        X = fr.X
        f.text(36 + i * 465, 28, title, "th")
        f.text(36 + i * 465, 48, verdict, vcls)
        if i == 0:
            draw_existing(f, fr, YT, YB, pocket=False)
            f.poly(fr([(-5, 2), (99.5, 2), (99.5, H - 1), (-5, H - 1)]), "fl", closed=True, fill="mw")
            draw_skins(f, fr, YT, YB)
            f.poly(fr([(-1, 510), (-6, 510), (-6, -40), (-11, -47)]), "nw")
            f.poly(fr([(-6.5, 510), (-0.5, 510), (-0.5, 530)]), "sl", closed=True)
            f.poly(fr([(-3, 546), (-3, 516)]), "wt", end="aw")
            f.poly(fr([(-3, 506), (-3, 474), (14, 446), (36, 415)]), "lk", end="ab")
            f.poly(fr([(-11, 500), (-11, -36), (-16, -46)]), "wt", end="aw")
            fr.cover(f)
            f.text(X(-24), fr.Y(536), "기존 외부강판(상부)", "tx", "end")
            f.text(X(-24), fr.Y(514), "실리콘 (시간이 지나면 열화)", "tb", "end")
            f.text(X(-24), fr.Y(440), "모듈 전면판", "tx", "end")
            f.text(X(-24), fr.Y(422), "기존 강판 바깥에 덧댐", "ts", "end")
            f.text(X(-24), fr.Y(-26), "기존 외부강판(하부)", "tx", "end")
            f.text(X(108), fr.Y(430), "누수 → 내부로", "tb")
        else:
            m = module()
            draw_existing(f, fr, YT, YB)
            f.poly(fr([(-3, 2), (99.5, 2), (99.5, H - 1), (5.5, H - 1), (5.5, H - REVEAL), (-3, TOP_PLATE)]),
                   "fl", closed=True, fill="mw")
            draw_skins(f, fr, YT, YB)
            f.poly(fr(m["front"]), "nw")
            f.poly(fr([(-10, 546), (-10, -38), (-16, -50)]), "wt", end="aw")
            fr.cover(f)
            f.text(X(-24), fr.Y(536), "기존 외부강판(상부)", "tx", "end")
            f.text(X(-24), fr.Y(500), "텅: 기존 강판 뒤로 70", "tn", "end")
            f.text(X(-24), fr.Y(430), "모듈 전면판", "tx", "end")
            f.text(X(-24), fr.Y(-8), "스커트: 기존 강판 위로", "tn", "end")
            f.text(X(-24), fr.Y(-34), "기존 외부강판(하부)", "tx", "end")
    return f


# ------------------------------------------------------------------ D-02 recommended section
def d02() -> Fig:
    f = Fig("d02", 980, 830,
            "추천안 수직 단면: 카세트의 텅이 기존 상부 외부강판 뒤 포켓에 70 mm 들어가고, 스커트가 스타터레일과 기존 하부 "
            "외부강판 바깥을 덮는다. 카세트는 상부 걸이레일과 스타터레일 리브에 걸리며, 미네랄울이 외부강판 면에서 내부강판까지 "
            "418 mm 높이로 채워진다")
    YT, YB = 610, -130
    fr = BrokenFrame(305, 84, 2.6, 1.5, YT, 370, 70, 44)
    X, Y = fr.X, fr.Y

    draw_existing(f, fr, YT, YB)
    m = module()
    f.poly(fr(m["wool"]), "fl", closed=True, fill="mw")
    f.poly(fr(m["strip"]), "fl", closed=True, fill="ms")
    f.poly(fr([(112, 50), (101.5, 50), (101.5, 0), (160, 0)]), "gt")       # existing girt (C-section)
    f.poly(fr([(160, -8), (160, 8)]), "brk")
    draw_skins(f, fr, YT, YB)
    f.poly(fr(STARTER), "nw")
    f.poly(fr(TOPRAIL), "nw")
    draw_module(f, fr, m, fills=False)
    screw(f, fr, 25, 96.5, 110)
    screw(f, fr, TOP_PLATE - 14, 97, 106)

    # water
    f.poly(fr([(-9, YT - 6), (-9, -36), (-15, -47)]), "wt", end="aw")
    f.poly(fr([(-13, -62), (-6, -76), (-6, -122)]), "wt", end="aw")
    f.path(f"M{X(-17):.1f},{Y(437):.1f} Q{X(-2):.1f},{Y(441):.1f} {X(2.2):.1f},{Y(452):.1f} "
           f"L{X(2.2):.1f},{Y(505):.1f}", "wt2", end="aw")
    f.poly(fr([(16, 6), (-3.25, 6), (-3.25, -46), (-8, -60)]), "wt2", end="aw")
    fr.cover(f)
    f.text(X(-9) - 6, Y(YT - 4), "빗물", "tw", "end")

    # dimensions: chains on the outside
    xa, xb = 246, 224
    for yy, x0 in ((-40, -10), (0, -2), (H, 0), (H + LAP, 4.5)):
        ext(f, X(x0) - 3, Y(yy), xa - 4, Y(yy))
    dim_v(f, xa, Y(0), Y(-40), "40")
    dim_v(f, xa, Y(H), Y(0), "개구 460 (+10/−0)")
    dim_v(f, xa, Y(H + LAP), Y(H), "겹침 70")
    for yy, x0 in ((-50, -2), (H - REVEAL, -4.5), (H + POCKET_H, 0)):
        ext(f, X(x0) - 3, Y(yy), xb - 4, Y(yy))
    ext(f, xa - 4, Y(0), xb - 4, Y(0))
    ext(f, xa - 4, Y(H), xb - 4, Y(H))
    dim_v(f, xb, Y(0), Y(-50), "50")
    dim_v(f, xb, Y(H), Y(H - REVEAL), "15")
    dim_v(f, xb, Y(H + POCKET_H), Y(H), "포켓 100")
    # top: thickness
    ext(f, X(0), Y(YT) - 4, X(0), 54)
    ext(f, X(T), Y(YT) - 4, X(T), 54)
    dim_h(f, 58, X(0), X(T), "t = 기존 코어 두께 (도면 100)")
    f.text(X(-12), 26, "실외", "ts", "end")
    f.text(X(108), 26, "실내", "ts", "start")

    legend(f, 748, 22)

    left = [((0, 590), 106, 1, "기존 외부강판(상부)"), ((4.5, 510), 234, 2, "텅 + 백댐 헴"),
            ((0, 438), 342, 3, "45° 조글 · 들임 15"), ((-4.5, 390), 414, 4, "카세트 전면판"),
            ((-2, -22), 610, 14, "스타터레일(물끊기 겸용)"), ((-8, -44), 662, 5, "스커트 + 물끊기"),
            ((0, -100), 748, 6, "기존 외부강판(하부)")]
    right = [((9, 545), 184, 7, "상부 포켓 10×100"), ((50, 445), 296, 8, "압축 스트립 40 → 30"),
             ((60, TOP_PLATE), 332, 9, "상부판 + 테일"), ((88, TOP_PLATE - 14), 370, 10, "상부 걸이레일"),
             ((50, 390), 410, 11, "미네랄울 높이 418 (≥400)"), ((100, 62), 504, 12, "기존 내부강판(존치)"),
             ((28, 8), 548, 13, "하부판 + 걸림 리브"), ((140, 0), 628, 15, "기존 띠장 + 피스 @300")]
    for (pt, by, n, label) in left:
        callout(f, fr([pt])[0], 196, by, n, label, "L")
    for (pt, by, n, label) in right:
        callout(f, fr([pt])[0], 745, by, n, label, "R")
    return f


# ------------------------------------------------------------------ D-03 installation motion
def d03() -> Fig:
    f = Fig("d03", 920, 470,
            "설치 동작 3단계: 카세트를 12 mm 들어올려 텅을 포켓에 끼우고, 하단을 안으로 밀어 넣은 뒤, 12 mm 내려놓으면 "
            "상부 테일이 걸이레일에, 하부 리브가 스타터레일 리브 뒤에 동시에 걸린다")
    steps = [("① 들어올려 텅 삽입", "비스듬히 12 mm 올려 텅을 포켓에", LIFT, 1.5),
             ("② 하단 밀어넣기", "하부 리브가 레일 리브 위를 지나감", LIFT, 0.0),
             ("③ 내려 걸기", "12 mm 내리면 상·하부가 함께 걸림", 0.0, 0.0)]
    YT, YB = 578, -75
    for i, (title, sub, lift, rot) in enumerate(steps):
        fr = BrokenFrame(100 + i * 298, 80, 1.3, 1.0, YT, 378, 62, 28)
        X, Y = fr.X, fr.Y
        f.text(X(-60), 30, title, "th")
        f.text(X(-60), 50, sub, "ts")
        draw_existing(f, fr, YT, YB)
        m = module(lift, rot)
        f.poly(fr(m["wool"]), "fl", closed=True, fill="mw")
        f.poly(fr(m["strip"]), "fl", closed=True, fill="ms")
        draw_skins(f, fr, YT, YB)
        f.poly(fr(STARTER), "nw")
        f.poly(fr(TOPRAIL), "nw")
        draw_module(f, fr, m, fills=False)
        fr.cover(f, -24, 112)
        if i == 0:
            f.line(X(-40), Y(470), X(-40), Y(540), "mv", end="ak")
            f.text(X(-46), Y(500), "↑12", "tx", "end")
        elif i == 1:
            f.line(X(-62), Y(25), X(-22), Y(25), "mv", end="ak")
            f.text(X(-62), Y(42), "밀기", "tx")
        else:
            f.line(X(-40), Y(540), X(-40), Y(470), "mv", end="ak")
            f.text(X(-46), Y(500), "↓12", "tx", "end")
            for cx, cy, label in ((92, TOP_PLATE - 12, "상부 걸림"), (26, 8, "하부 걸림")):
                f.circle(X(cx), Y(cy), 13, "hl")
                f.line(X(cx) + 13, Y(cy), X(112), Y(cy), "ld")
                f.text(X(114), Y(cy) + 4, label, "tn")
    return f


# ------------------------------------------------------------------ D-04 plan joint
def d04() -> Fig:
    f = Fig("d04", 920, 440,
            "모듈 간 이음 수평 단면: 카세트 A의 전면판이 50 mm 연장되어 카세트 B의 전면판 뒤에 겹치고, 연장부의 배수 채널이 "
            "이음으로 들어온 물을 받아 아래 스타터레일로 보낸다. 미네랄울은 50 mm 단차로 맞물려 이음부에도 끊기지 않는다")
    ox, oy, ss, sd = 460, 112, 1.4, 2.6

    def P(pts):
        return [(ox + s * ss, oy + d * sd) for s, d in pts]

    S0, S1 = -290, 290
    f.poly(P([(S0, -3.5), (0, -3.5), (0, 48), (50, 48), (50, 99.5), (S0, 99.5)]), "fl", closed=True, fill="mw")
    f.poly(P([(0, 11), (50, 11), (50, -3.5), (S1, -3.5), (S1, 99.5), (50, 99.5), (50, 48), (0, 48)]),
           "fl", closed=True, fill="mw")
    f.poly(P([(0, -3.5), (0, 48), (50, 48), (50, 99.5)]), "jt")
    f.poly(P([(S0, 100), (S1, 100)]), "sk")
    f.poly(P([(S0, -4.5), (-4, -4.5), (0, -2.5), (19, -2.5), (19, 8), (31, 8), (31, -2.5), (50, -2.5), (50, 3)]), "nw")
    f.poly(P([(0, -4.5), (S1, -4.5)]), "nw")
    for s in (S0, S1):
        x = ox + s * ss
        f.line(x, oy - 22, x, oy + 280, "brk")

    f.poly(P([(-70, -30), (-6, -9), (8, -4), (25, 4)]), "wt", end="aw")
    f.text(ox - 72 * ss, oy - 33 * sd, "바람에 밀린 빗물", "tw", "end")
    f.line(ox + 31 * ss, oy + 8 * sd, ox + 64 * ss, oy - 12 * sd, "ld")
    f.text(ox + 66 * ss, oy - 12 * sd + 4, "배수 채널 12×10 → 아래 스타터레일로", "tw")

    x0, x50 = ox, ox + 50 * ss
    f.line(x0, oy - 6 * sd, x0, oy - 27 * sd, "dm")
    f.line(x50, oy - 3 * sd, x50, oy - 27 * sd, "dm")
    dim_h(f, oy - 24 * sd, x0, x50, "겹침 50")

    f.text(ox - 150 * ss, oy + 52 * sd, "카세트 A", "th halo", "middle")
    f.text(ox + 175 * ss, oy + 52 * sd, "카세트 B", "th halo", "middle")
    f.line(ox + 50 * ss, oy + 74 * sd, ox + 82 * ss, oy + 74 * sd, "ld")
    f.text(ox + 84 * ss, oy + 72 * sd, "미네랄울 50 단차 겹침", "tn halo")
    f.text(ox + 84 * ss, oy + 72 * sd + 17, "이음부에도 불연재가 이어짐", "tx halo")
    f.text(ox + (S0 + 6) * ss, oy - 30 * sd, "실외", "ts")
    f.text(ox + (S0 + 6) * ss, oy + 118 * sd, "실내", "ts")
    f.text(ox + (S1 - 6) * ss, oy + 113 * sd, "기존 내부강판", "ts", "end")
    f.text(ox - 150 * ss, oy - 9 * sd, "전면판 (길이 1,000)", "ts", "middle")
    return f


# ------------------------------------------------------------------ D-05 options
def d05() -> Fig:
    f = Fig("d05", 920, 440,
            "대안 비교 단면: A안은 공장에서 일체화한 카세트를 레일 두 줄에 거는 방식, B안은 미네랄울 인서트를 먼저 넣고 전면 "
            "커버를 덮는 방식, C안은 기성 불연 샌드위치패널을 띠 모양으로 잘라 플래싱과 관통 피스로 붙이는 방식")
    panels = [("A안 · 걸이식 일체 카세트 (추천)", "현장 부재 3종 · 노출 체결 없음"),
              ("B안 · 3부재 분리형", "현장 부재 4종 · 커버 하단 리벳"),
              ("C안 · 기성 불연패널 스트립", "금형 불필요 · 관통 피스 노출")]
    YT, YB = 578, -75
    for i, (title, sub) in enumerate(panels):
        fr = BrokenFrame(50 + i * 305, 78, 1.1, 0.9, YT, 378, 62, 28)
        X, Y = fr.X, fr.Y
        f.text(X(-25), 28, title, "th")
        f.text(X(-25), 47, sub, "ts")
        draw_existing(f, fr, YT, YB)
        if i == 0:
            m = module()
            f.poly(fr(m["wool"]), "fl", closed=True, fill="mw")
            f.poly(fr(m["strip"]), "fl", closed=True, fill="ms")
            draw_skins(f, fr, YT, YB)
            f.poly(fr(STARTER), "nw")
            f.poly(fr(TOPRAIL), "nw")
            draw_module(f, fr, m, fills=False)
            labels = [((92, TOP_PLATE - 12), 500, "상부 걸이레일"), ((50, 400), 420, "카세트 (공장 일체)"),
                      ((60, 1.5), 40, "스타터레일")]
        elif i == 1:
            m = module()
            f.poly(fr([(-4, 4), (99.5, 4), (99.5, H - 1), (-4, H - 1)]), "fl", closed=True, fill="mw")
            draw_skins(f, fr, YT, YB)
            f.poly(fr(STARTER_PLAIN), "nw")
            f.poly(fr(TOPRAIL), "nw")
            f.poly(fr(m["top"]), "nw2")
            f.poly(fr(m["front"]), "nw")
            f.circle(X(-4.5), Y(-22), 2.6, "ldd")
            labels = [((92, TOP_PLATE - 12), 500, "상부 걸이레일"), ((-4.5, 400), 455, "전면 커버 (공용)"),
                      ((50, 395), 395, "미네랄울 인서트"), ((-4.5, -22), 45, "하단 리벳 (노출)"),
                      ((60, 1.5), 10, "스타터레일")]
        else:
            f.poly(fr([(-4.5, 12), (99, 12), (99, H - 6), (-4.5, H - 6)]), "fl", closed=True, fill="mw")
            draw_skins(f, fr, YT, YB)
            f.poly(fr([(-4.5, 12), (-4.5, H - 6)]), "nw")
            f.poly(fr([(99, 12), (99, H - 6)]), "nw2")
            f.poly(fr([(-7, 400), (-7, 428), (4.5, 443), (4.5, H + LAP), (8.5, H + LAP)]), "nw")
            f.circle(X(-7), Y(412), 2.6, "ldd")
            f.poly(fr(STARTER_PLAIN), "nw")
            screw(f, fr, 30, -8, 112)
            labels = [((-7, 412), 520, "Z-플래싱 + 리벳"), ((50, 395), 420, "불연패널 스트립"),
                      ((40, 30), 52, "관통 피스 (노출)"), ((60, 1.5), 8, "하부 물끊기")]
        fr.cover(f, -24, 112)
        for (pt, ly, label) in labels:
            tx, ty = fr([pt])[0]
            lx = X(112)
            f.poly([(tx, ty), (lx - 4, Y(ly) - 4)], "ld")
            f.circle(tx, ty, 1.6, "ldd")
            f.text(lx, Y(ly), label, "ts")
    return f


# ------------------------------------------------------------------ build
FIGS = {"D01": (d01, "d01_principle"), "D02": (d02, "d02_section"), "D03": (d03, "d03_install"),
        "D04": (d04, "d04_joint_plan"), "D05": (d05, "d05_options")}


def main():
    FIG_DIR.mkdir(exist_ok=True)
    page = (HERE / "template.html").read_text(encoding="utf-8")
    page = page.replace("/*SVG-RULES*/", RULES.strip())
    for key, (fn, name) in FIGS.items():
        fig = fn()
        (FIG_DIR / f"{name}.svg").write_text(fig.render(standalone=True), encoding="utf-8")
        page = page.replace(f"<!--{key}-->", fig.render(standalone=False))
    (HERE / "index.html").write_text(page, encoding="utf-8")
    print("wrote", ", ".join(f"figures/{n}.svg" for _, n in FIGS.values()), "and index.html")


if __name__ == "__main__":
    main()
