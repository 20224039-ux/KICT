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


def dim_v(f: Fig, X, Y1, Y2, label, side=-1, cls="dx"):
    """vertical dimension line at px X between px Y1..Y2, text rotated along it."""
    f.line(X, Y1, X, Y2, "dm")
    for yy in (Y1, Y2):
        f.line(X - 3.5, yy + 3.5, X + 3.5, yy - 3.5, "dm")
    tx = X - 4 if side < 0 else X + 12
    f.text(tx, (Y1 + Y2) / 2, label, cls, "middle", rot=-90)


def dim_h(f: Fig, Y, X1, X2, label, cls="dx"):
    f.line(X1, Y, X2, Y, "dm")
    for xx in (X1, X2):
        f.line(xx - 3.5, Y + 3.5, xx + 3.5, Y - 3.5, "dm")
    f.text((X1 + X2) / 2, Y - 5, label, cls, "middle")


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


def brk(f: Fig, fr: Frame, y, x0=-9.0, x1=109.0, amp=12.0):
    """horizontal break line across the wall section at elevation y."""
    m = (x0 + x1) / 2
    w = amp / 2
    f.poly(fr([(x0, y), (m - w, y), (m - w / 2, y + amp), (m + w / 2, y - amp), (m + w, y), (x1, y)]), "brk")


def note(f: Fig, target, tx, ty, lines, side):
    """leader from a px target to a text block; first line ink, the rest muted."""
    x0, y0 = target
    ex = tx + 6 if side == "L" else tx - 6
    f.line(x0, y0, ex, ty - 4, "ld")
    f.circle(x0, y0, 1.8, "ldd")
    anchor = "end" if side == "L" else "start"
    for k, line in enumerate(lines):
        f.text(tx, ty + k * 15, line, "tx" if k == 0 else "ts", anchor)


def clip_y(pts, ylo, yhi):
    """clip a polyline (mm) to ylo <= y <= yhi; returns the list of visible pieces."""
    pieces, cur = [], []

    def inside(p):
        return ylo <= p[1] <= yhi

    def cut(a, b, yc):
        t = (yc - a[1]) / (b[1] - a[1])
        return (a[0] + t * (b[0] - a[0]), yc)

    for a, b in zip(pts, pts[1:]):
        ia, ib = inside(a), inside(b)
        if ia and not cur:
            cur = [a]
        if ia and ib:
            cur.append(b)
        elif ia and not ib:
            cur.append(cut(a, b, yhi if b[1] > yhi else ylo))
            pieces.append(cur)
            cur = []
        elif not ia and ib:
            cur = [cut(a, b, yhi if a[1] > yhi else ylo), b]
        elif (a[1] - ylo) * (b[1] - ylo) < 0 or (a[1] - yhi) * (b[1] - yhi) < 0:
            pieces.append([cut(a, b, yhi if a[1] > yhi else ylo), cut(a, b, ylo if a[1] > yhi else yhi)])
    if len(cur) > 1:
        pieces.append(cur)
    return pieces


# ------------------------------------------------------------------ geometry (mm)
# 상부 단열재에는 손대지 않는다. 기존 외부강판과 단열재는 한 번의 절단선(H)에서 같이 잘린다.
# 겹침에 필요한 '뒤 공간'은 개구 안에서 만든다: 새 캡레일의 립이 기존 강판 면을 아래로 연장하고,
# 카세트 텅이 그 립 뒤로 들어간다.
T = 100.0        # example core thickness (t)
H = 460.0        # cut line: outer skin + core removed between EL±0 and EL+460 (one saw line each)
LIP = 90.0       # cap-rail lip below the cut line (continues the existing skin face downward)
LAP = 70.0       # tongue behind the lip
REVEAL = 15.0    # open reveal between lip and joggle (room for the 12 mm lift)
LIFT = 12.0      # cassette lift during install; both hooks engage 8 mm
LIP_BOT = H - LIP                  # 370
TOP = LIP_BOT + LAP                # 440: tongue top = cassette top plate (one bent sheet)
JTOP = LIP_BOT - REVEAL            # 355
JBOT = JTOP - 15.0                 # 340
BOT = 12.0                         # bottom plate rests on the starter-rail rib (12 high)

CAP = [(-3, LIP_BOT - 4), (0.8, LIP_BOT), (0.8, H), (8, H), (10, H - 4), (12, H),
       (98.5, H), (98.5, H - 60), (88, H - 60), (88, H - 30)]
STARTER = [(98, 50), (98, 1.5), (24, 1.5), (24, BOT), (19, BOT), (19, 1.5), (-2, 1.5), (-2, -50), (-7, -57)]
STARTER_PLAIN = [(98, 50), (98, 1.5), (-2, 1.5), (-2, -50), (-7, -57)]


def module(lift=0.0, rot_deg=0.0):
    """cassette outline (mm); lifted by `lift` and tilted bottom-out about the tongue top."""
    pivot = (4.5, TOP + lift)
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

    return {
        # one bent sheet: tail – top plate – tongue – joggle – face – skirt – kick-out
        "front": tf([(94, TOP - 18), (94, TOP), (4.5, TOP), (4.5, JTOP), (-4.5, JBOT), (-4.5, -40), (-10, -47)]),
        "bot": tf([(-2.5, BOT + 20), (-2.5, BOT), (26, BOT), (26, BOT - 8), (30, BOT - 8), (30, BOT),
                   (88, BOT), (88, BOT + 10)]),
        "wool": tf([(-4, BOT), (97, BOT), (97, 54), (99.5, 54), (99.5, H - 66), (86, H - 66), (86, TOP),
                    (5.5, TOP), (5.5, JTOP), (-4, JBOT)]),
        # compressible strip between top plate and cap rail; squeezes when lifted
        "strip": [(5.5, TOP + lift), (97, TOP + lift), (97, H), (5.5, H)],
    }


def draw_existing(f: Fig, fr: Frame, ytop, ybot):
    f.poly(fr([(0, ytop), (T, ytop), (T, H), (0, H)]), "fl", closed=True, fill="core")
    f.poly(fr([(0, 0), (T, 0), (T, ybot), (0, ybot)]), "fl", closed=True, fill="core")


def draw_skins(f: Fig, fr: Frame, ytop, ybot, amp=12.0):
    f.poly(fr([(0, H), (0, ytop)]), "sk")
    f.poly(fr([(0, ybot), (0, 0)]), "sk")
    f.poly(fr([(T, ybot), (T, ytop)]), "sk")
    brk(f, fr, ytop, amp=amp)
    brk(f, fr, ybot, amp=amp)


def draw_module(f: Fig, fr: Frame, m, fills=True):
    if fills:
        f.poly(fr(m["wool"]), "fl", closed=True, fill="mw")
        f.poly(fr(m["strip"]), "fl", closed=True, fill="ms")
    f.poly(fr(m["bot"]), "nw2")
    f.poly(fr(m["front"]), "nw")


def screw(f: Fig, fr: Frame, y, x_head, x_tip, head=5):
    hx, ty = fr.X(x_head), fr.Y(y)
    f.line(hx, ty - head, hx, ty + head, "sc")
    f.line(hx, ty, fr.X(x_tip), ty, "sc")


def legend(f: Fig, x, y, cols=1, rows=None):
    rows = rows or [
        ("line", "wt", "빗물 경로"), ("line", "nw", "신설 부재"),
        ("line", "sk", "기존 부재"), ("fill", "core", "기존 단열재(EPS·PIR)"),
        ("fill", "mw", "미네랄울(불연)"), ("fill", "ms", "압축 불연 스트립"),
    ]
    for k, (kind, c, label) in enumerate(rows):
        cx = x + (k % cols) * 150
        cy = y + (k // cols) * 18
        if kind == "line":
            f.line(cx, cy - 4, cx + 24, cy - 4, c)
        else:
            f.rect(cx, cy - 10, 24, 11, "fl", fill=c)
        f.text(cx + 30, cy, label, "ts")


def poly_clip(f: Fig, fr: Frame, pts, cls, ylo, yhi):
    for piece in clip_y(pts, ylo, yhi):
        f.poly(fr(piece), cls)


# ------------------------------------------------------------------ D-01 principle
def d01() -> Fig:
    f = Fig("d01", 920, 470,
            "원리 비교: 모듈을 기존 외부강판 바깥에 덧대면 상부 겹침이 물 흐름과 반대가 되어 실리콘이 유일한 방수선이 된다. "
            "제안안은 상부 단열재를 건드리지 않고, 새 캡레일 립이 기존 강판 면을 아래로 이어 주며 카세트 텅이 그 립 뒤로 겹친다")
    YT, YB = 540, -70
    titles = [("기존 방식 · 바깥에 덧댐", "상부 겹침이 물 흐름과 반대 → 실리콘이 유일한 방수선", "tb"),
              ("제안 · 캡레일 립 뒤로 겹침", "상부 단열재 그대로 · 모든 겹침이 위가 아래를 덮음", "tk")]
    for i, (title, verdict, vcls) in enumerate(titles):
        fr = BrokenFrame(235 + i * 465, 72, 1.6, 1.0, YT, 320, 60, 30)
        X = fr.X
        f.text(36 + i * 465, 28, title, "th")
        f.text(36 + i * 465, 48, verdict, vcls)
        draw_existing(f, fr, YT, YB)
        if i == 0:
            f.poly(fr([(-5, 2), (99.5, 2), (99.5, H - 1), (-5, H - 1)]), "fl", closed=True, fill="mw")
            draw_skins(f, fr, YT, YB)
            f.poly(fr([(-1, 510), (-6, 510), (-6, -40), (-11, -47)]), "nw")
            f.poly(fr([(-6.5, 510), (-0.5, 510), (-0.5, 530)]), "sl", closed=True)
            f.poly(fr([(-3, 538), (-3, 516)]), "wt", end="aw")
            f.poly(fr([(-3, 506), (-3, 474), (14, 446), (36, 405)]), "lk", end="ab")
            f.poly(fr([(-11, 500), (-11, -36), (-16, -46)]), "wt", end="aw")
            fr.cover(f)
            f.text(X(-24), fr.Y(530), "기존 외부강판(상부)", "tx", "end")
            f.text(X(-24), fr.Y(508), "실리콘 (시간이 지나면 열화)", "tb", "end")
            f.text(X(-24), fr.Y(420), "모듈 전면판", "tx", "end")
            f.text(X(-24), fr.Y(402), "기존 강판 바깥에 덧댐", "ts", "end")
            f.text(X(-24), fr.Y(-26), "기존 외부강판(하부)", "tx", "end")
            f.text(X(108), fr.Y(430), "누수 → 내부로", "tb")
        else:
            m = module()
            f.poly(fr([(-3, 2), (99.5, 2), (99.5, H - 1), (5.5, H - 1), (5.5, JTOP), (-3, JBOT)]),
                   "fl", closed=True, fill="mw")
            draw_skins(f, fr, YT, YB)
            f.poly(fr(CAP[:7]), "nw")
            f.poly(fr(m["front"][2:]), "nw")
            f.poly(fr([(-10, 538), (-10, -38), (-16, -50)]), "wt", end="aw")
            fr.cover(f)
            f.text(X(-24), fr.Y(530), "기존 외부강판(상부) · 단열재 그대로", "tx", "end")
            f.text(X(-24), fr.Y(430), "캡레일 립: 강판 면을 아래로 연장", "tn", "end")
            f.text(X(-24), fr.Y(372), "텅: 립 뒤로 70 겹침", "tn", "end")
            f.text(X(-24), fr.Y(335), "모듈 전면판", "tx", "end")
            f.text(X(-24), fr.Y(-8), "스커트: 기존 강판 위로", "tn", "end")
            f.text(X(-24), fr.Y(-34), "기존 외부강판(하부)", "tx", "end")
    return f


# ------------------------------------------------------------------ D-02 top joint (enlarged)
def d02() -> Fig:
    f = Fig("d02", 1000, 800,
            "상부 조인트 확대 상세: 기존 외부강판과 단열재는 같은 선에서 잘린 그대로 두고, 그 아래에 캡레일을 붙인다. "
            "캡레일 상판은 단열재 절단면에 밀착하고, 립은 기존 강판 면보다 0.5 mm 안쪽에서 90 mm 내려온다. "
            "카세트 텅은 립 뒤 3 mm 등압 공간에서 70 mm 겹치며, 립 하단과 조글 사이 15 mm 줄눈으로 물이 빠진다")
    YT, YB = 512, 318
    fr = Frame(345, 60 + YT * 3.6, 3.6, 3.6)
    X, Y = fr.X, fr.Y
    m = module()
    f.poly(fr([(0, YT), (T, YT), (T, H), (0, H)]), "fl", closed=True, fill="core")
    f.poly(fr([(p[0], max(p[1], YB)) for p in m["wool"]]), "fl", closed=True, fill="mw")
    f.poly(fr(m["strip"]), "fl", closed=True, fill="ms")
    f.poly(fr([(0, H), (0, YT)]), "sk")
    f.poly(fr([(T, YB), (T, YT)]), "sk")
    brk(f, fr, YT, -9, 109, amp=4)
    brk(f, fr, YB, -12, 109, amp=4)
    f.poly(fr(CAP), "nw")
    poly_clip(f, fr, m["front"], "nw", YB, YT)
    screw(f, fr, H - 50, 97, 106, head=6)

    # water
    f.poly(fr([(-2.2, 505), (-2.2, H + 1.5), (-1.6, H - 1.5), (-1.6, LIP_BOT + 1.5), (-5.2, LIP_BOT - 4.2)]),
           "wt", end="aw")
    f.poly(fr([(-5.4, LIP_BOT - 7), (-5.4, JBOT + 2.5)]), "wt2", end="aw")
    f.poly(fr([(-7, JBOT - 2), (-7, YB + 4)]), "wt", end="aw")
    f.path(f"M{X(-20):.1f},{Y(347):.1f} Q{X(-3):.1f},{Y(350):.1f} {X(2.9):.1f},{Y(364):.1f} "
           f"L{X(2.9):.1f},{Y(418):.1f}", "wt2", end="aw")
    f.poly(fr([(1.5, H + 1.6), (8.2, H + 1.6)]), "wt2", end="aw")
    f.text(X(-2.2) - 6, Y(507), "빗물", "tw", "end")

    # dimensions
    xd = X(-13)
    for yy, x0 in ((H, 0), (LIP_BOT, -3), (JTOP, 4.5)):
        ext(f, X(x0) - 3, Y(yy), xd - 4, Y(yy))
    dim_v(f, xd, Y(H), Y(LIP_BOT), "립 90")
    dim_v(f, xd, Y(LIP_BOT), Y(JTOP), "15")
    xl = X(26)
    ext(f, X(5), Y(TOP), xl + 4, Y(TOP))
    ext(f, X(2), Y(LIP_BOT), xl + 4, Y(LIP_BOT))
    dim_v(f, xl, Y(TOP), Y(LIP_BOT), "겹침 70", side=1, cls="dx halo")
    xs = X(74)
    dim_v(f, xs, Y(H), Y(TOP), "20", side=1, cls="dx halo")

    left = [((0, 492), 492, ["기존 외부강판(상부)", "그대로 둠 · 절단 끝 징크리치 방청"]),
            ((0.4, H), 466, ["맞댐부", "립 바깥면 = 강판 안쪽면 (0.5 mm 들어감)", "물이 계단처럼 아래 판으로 넘어감"]),
            ((0.8, 410), 414, ["캡레일 물끊기 립 90", "1.0t 도장강판 (기존 패널 색)"]),
            ((-2, LIP_BOT - 3), 378, ["립 하단 킥아웃", "물방울을 바깥으로 떨굼"]),
            ((0.5, 362), 356, ["들임 줄눈 15", "바깥과 통하는 개구 → 등압"]),
            ((0.5, 347), 334, ["45° 조글", "떨어진 물을 전면판 바깥으로"])]
    for pt, ly, lines in left:
        note(f, fr([pt])[0], 250, Y(ly), lines, "L")
    note(f, fr([(-4.5, 326)])[0], 250, Y(322), ["카세트 전면판 0.8t"], "L")

    right = [((50, 492), 496, ["기존 단열재(상부)", "자르거나 파내지 않음"]),
             ((60, H), 470, ["캡레일 상판", "단열재 절단면에 밀착 (화염·빗물 차단)"]),
             ((10, H - 3.5), 452, ["모세관 차단 홈", "맞댐부로 스민 물이 안쪽으로 못 감"]),
             ((60, 450), 434, ["압축 불연 스트립 30 → 20"]),
             ((40, TOP), 420, ["텅 + 상부판 일체 절곡"]),
             ((93.5, TOP - 16), 406, ["테일이 걸이 채널에 걸림 8"]),
             ((100, H - 50), 392, ["캡레일 피스 @150 → 내부강판"]),
             ((4.5, 396), 375, ["텅: 립 뒤로 겹침 70", "틈 3 등압 공간 → 물이 못 올라감"]),
             ((50, 345), 345, ["미네랄울 2호 이상 (t + 10 압축)"]),
             ((100, 328), 328, ["기존 내부강판"])]
    for pt, ly, lines in right:
        note(f, fr([pt])[0], 752, Y(ly), lines, "R")
    f.text(X(-12), 30, "실외", "ts", "end")
    f.text(X(108), 30, "실내", "ts")
    return f


# ------------------------------------------------------------------ D-03 bottom joint (enlarged)
def d03() -> Fig:
    f = Fig("d03", 1000, 600,
            "하부 조인트 확대 상세: 스타터레일 선반이 하부 단열재 절단면과 기존 강판 끝을 덮고 앞날개가 기존 강판 바깥을 50 mm "
            "덮는다. 카세트 스커트가 그 바깥을 다시 40 mm 덮고, 하부판은 레일 리브 위에 얹혀 하향 리브가 리브 뒤에 8 mm 걸린다")
    YT, YB = 64, -72
    fr = Frame(345, 60 + YT * 3.6, 3.6, 3.6)
    X, Y = fr.X, fr.Y
    m = module()
    f.poly(fr([(0, 0), (T, 0), (T, YB), (0, YB)]), "fl", closed=True, fill="core")
    f.poly(fr([(-4, BOT), (97, BOT), (97, 54), (99.5, 54), (99.5, YT), (-4, YT)]), "fl", closed=True, fill="mw")
    f.poly(fr([(110, 50), (101.5, 50), (101.5, 0), (112, 0)]), "gt")
    f.poly(fr([(0, YB), (0, 0)]), "sk")
    f.poly(fr([(T, YB), (T, YT)]), "sk")
    brk(f, fr, YT, -12, 109, amp=4)
    brk(f, fr, YB, -9, 109, amp=4)
    f.poly(fr(STARTER), "nw")
    f.poly(fr(m["bot"]), "nw2")
    poly_clip(f, fr, m["front"], "nw", YB, YT)
    screw(f, fr, 25, 96.5, 110, head=6)

    # water: face film → skirt kick-out → free drip; joint water → shelf → behind skirt → out
    f.poly(fr([(-7, 60), (-7, -40.5), (-12.5, -47.5)]), "wt", end="aw")
    f.poly(fr([(-12.5, -53), (-12.5, -70)]), "wt2", end="aw")
    f.poly(fr([(3, 60), (3, 5), (-0.5, 5), (-3.25, 2), (-3.25, -45), (-8.5, -59.5)]), "wt2", end="aw")
    f.poly(fr([(44, 5), (17, 5)]), "wt2", end="aw")
    f.text(X(3) + 6, Y(56), "세로 조인트 배수채널에서 내려온 물", "tw halo")

    xa, xb = X(-17), X(-22)
    for yy, x0, xx in ((0, -2, xb), (-40, -10, xa), (-50, -7, xb)):
        ext(f, X(x0) - 3, Y(yy), xx - 4, Y(yy))
    dim_v(f, xa, Y(0), Y(-40), "40")
    dim_v(f, xb, Y(0), Y(-50), "50")

    left = [((-4.5, 40), 46, ["카세트 전면판"]),
            ((3, BOT), 24, ["하부판 배수 노치", "세로 조인트 위치"]),
            ((-4.5, -14), -8, ["스커트: 레일 바깥 40 겹침"]),
            ((-2, -30), -25, ["레일 앞날개: 기존 강판 바깥 50 겹침"]),
            ((-8, -45), -42, ["스커트 킥아웃 (물끊기)"]),
            ((-5, -55.5), -57, ["레일 킥아웃 (물끊기)"])]
    for pt, ly, lines in left:
        note(f, fr([pt])[0], 250, Y(ly), lines, "L")
    note(f, fr([(0, -66)])[0], 250, Y(-70), ["기존 외부강판(하부)", "끝단은 레일 선반이 덮음"], "L")

    right = [((100, 60), 60, ["기존 내부강판"]),
             ((108, 50), 48, ["기존 띠장 (C형강)"]),
             ((50, 40), 36, ["미네랄울 2호 이상"]),
             ((109, 25), 24, ["뒷다리 피스 @300 → 띠장"]),
             ((60, BOT), 12, ["하부판 0.6t: 레일 리브 위에 얹힘"]),
             ((80, 1.5), 0, ["스타터레일 선반 1.0t STS", "하부 단열재 절단면과 강판 끝 덮개"]),
             ((21.5, BOT - 2), -16, ["레일 리브 12 + 배수 노치 @150"]),
             ((28, 5), -30, ["하향 리브: 레일 리브 뒤에 걸림 8"])]
    for pt, ly, lines in right:
        note(f, fr([pt])[0], 770, Y(ly), lines, "R")
    f.text(X(-12), 30, "실외", "ts", "end")
    f.text(X(108), 30, "실내", "ts")
    return f


# ------------------------------------------------------------------ D-04 vertical joint (plan)
def d04() -> Fig:
    f = Fig("d04", 920, 440,
            "카세트 세로 조인트 수평 단면: 카세트 A의 전면판이 50 mm 연장되어 카세트 B의 전면판 뒤에 겹치고, 연장부의 배수 채널이 "
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


# ------------------------------------------------------------------ D-05 overall section (where the joints are)
def d05() -> Fig:
    f = Fig("d05", 920, 640,
            "전체 수직 단면: 상부 조인트(상세 1)와 하부 조인트(상세 2)의 위치. 기존 외부강판과 단열재를 460 mm 걷어 낸 개구에 "
            "상부 캡레일과 스타터레일을 고정하고 카세트를 건다")
    YT, YB = 540, -110
    fr = BrokenFrame(360, 70, 2.2, 1.1, YT, 300, 70, 40)
    X, Y = fr.X, fr.Y
    draw_existing(f, fr, YT, YB)
    m = module()
    f.poly(fr(m["wool"]), "fl", closed=True, fill="mw")
    f.poly(fr(m["strip"]), "fl", closed=True, fill="ms")
    f.poly(fr([(112, 50), (101.5, 50), (101.5, 0), (150, 0)]), "gt")
    draw_skins(f, fr, YT, YB)
    f.poly(fr(CAP), "nw")
    f.poly(fr(STARTER), "nw")
    draw_module(f, fr, m, fills=False)
    f.poly(fr([(-9, YT - 6), (-9, -36), (-15, -47)]), "wt", end="aw")
    fr.cover(f, -20, 112)

    for (cx, cy, r, label) in ((4, 395, 62, "상세 ① 상부 조인트 (D-02)"), (-2, -12, 58, "상세 ② 하부 조인트 (D-03)")):
        f.circle(X(cx), Y(cy), r, "hl")
        f.line(X(cx) - r * 0.7, Y(cy) - r * 0.7, 236, Y(cy) - r * 0.7, "ld")
        f.text(232, Y(cy) - r * 0.7 + 4, label, "tn", "end")

    xa = X(-48)
    for yy, x0 in ((0, -2), (H, 0)):
        ext(f, X(x0) - 3, Y(yy), xa - 4, Y(yy))
    dim_v(f, xa, Y(H), Y(0), "개구 460 (+10/−0)")
    ext(f, X(99), Y(TOP), X(124), Y(TOP))
    ext(f, X(99), Y(BOT), X(124), Y(BOT))
    dim_v(f, X(121), Y(TOP), Y(BOT), "미네랄울 428 ≥ 400", side=1)

    labels = [((60, 505), 505, "기존 단열재(상부) · 그대로"), ((60, H), 470, "상부 캡레일"),
              ((50, 330), 330, "카세트 (1,000 단위)"), ((100, 62), 62, "기존 내부강판"),
              ((60, 1.5), 20, "스타터레일"), ((112, 50), -20, "기존 띠장")]
    for pt, ly, label in labels:
        tx, ty = fr([pt])[0]
        f.line(tx, ty, X(150) - 4, Y(ly) - 4, "ld")
        f.circle(tx, ty, 1.8, "ldd")
        f.text(X(150), Y(ly), label, "tx")
    legend(f, 20, 30, rows=[("line", "wt", "빗물 경로"), ("line", "nw", "신설 부재"), ("line", "sk", "기존 부재"),
                            ("fill", "core", "기존 단열재"), ("fill", "mw", "미네랄울")])
    return f


# ------------------------------------------------------------------ D-06 installation motion
def d06() -> Fig:
    f = Fig("d06", 920, 480,
            "설치 동작 3단계: 카세트를 12 mm 들어올려 텅을 캡레일 립 뒤로 넣고, 하단을 안으로 밀어 넣은 뒤, 12 mm 내려놓으면 "
            "상부 테일이 캡레일 걸이 채널에, 하부 리브가 스타터레일 리브 뒤에 동시에 걸린다")
    steps = [("① 들어올려 텅 삽입", "12 mm 올려 텅을 립 뒤로", LIFT, 1.5),
             ("② 하단 밀어넣기", "하부 리브가 레일 리브 위를 지나감", LIFT, 0.0),
             ("③ 내려 걸기", "12 mm 내리면 상·하부가 함께 걸림", 0.0, 0.0)]
    YT, YB = 520, -75
    for i, (title, sub, lift, rot) in enumerate(steps):
        fr = BrokenFrame(100 + i * 298, 80, 1.3, 1.1, YT, 320, 62, 28)
        X, Y = fr.X, fr.Y
        f.text(X(-60), 30, title, "th")
        f.text(X(-60), 50, sub, "ts")
        draw_existing(f, fr, YT, YB)
        m = module(lift, rot)
        f.poly(fr(m["wool"]), "fl", closed=True, fill="mw")
        f.poly(fr(m["strip"]), "fl", closed=True, fill="ms")
        draw_skins(f, fr, YT, YB)
        f.poly(fr(CAP), "nw")
        f.poly(fr(STARTER), "nw")
        draw_module(f, fr, m, fills=False)
        fr.cover(f, -24, 112)
        if i == 0:
            f.line(X(-40), Y(380), X(-40), Y(440), "mv", end="ak")
            f.text(X(-46), Y(405), "↑12", "tx", "end")
        elif i == 1:
            f.line(X(-62), Y(25), X(-22), Y(25), "mv", end="ak")
            f.text(X(-62), Y(42), "밀기", "tx")
        else:
            f.line(X(-40), Y(440), X(-40), Y(380), "mv", end="ak")
            f.text(X(-46), Y(405), "↓12", "tx", "end")
            for cx, cy, label in ((91, TOP - 13, "상부 걸림"), (26, 8, "하부 걸림")):
                f.circle(X(cx), Y(cy), 13, "hl")
                f.line(X(cx) + 13, Y(cy), X(112), Y(cy), "ld")
                f.text(X(114), Y(cy) + 4, label, "tn")
    return f


# ------------------------------------------------------------------ build
FIGS = {"D01": (d01, "d01_principle"), "D02": (d02, "d02_top_joint"), "D03": (d03, "d03_bottom_joint"),
        "D04": (d04, "d04_vertical_joint"), "D05": (d05, "d05_section"), "D06": (d06, "d06_install")}


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
