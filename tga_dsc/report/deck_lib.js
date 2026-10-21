// Shared frame for the TGA-DSC decks (pptxgenjs): theme, layouts, text helpers with real subscripts, native chart
// helpers with a manual plot area, and finish(), which writes the theme colours and repairs what pptxgenjs writes:
// replicate "#2" curves dashed, series without line or marker, blank chart points dropped, one <a:pPr> per paragraph
// (pptxgenjs writes one per run, and a later one would also cancel the bullet).
//
//   const { makeDeck } = require("./deck_lib");
//   const deck = makeDeck({ title, subject, footer });      // deck.pres, deck.content(...), ..., await deck.finish(file)
"use strict";
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const JSZip = require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));

const THEME = {
  name: "KICT TGA-DSC",
  headFontFace: "Malgun Gothic",
  bodyFontFace: "Malgun Gothic",
  colors: {
    dk1: "23303D", lt1: "FFFFFF", dk2: "4B5563", lt2: "F3F4F6",
    accent1: "EB6834", accent2: "2A78D6", accent3: "0CA30C", accent4: "FAB219", accent5: "EC835A", accent6: "D03B3B",
    hlink: "2A78D6", folHlink: "4B5563",
  },
};
// chart and status colours are data colours (validated palette, status palette never themed), so they stay hex
const HEX = {
  ink: "23303D", ink2: "4B5563", grid: "E5E7EB", axis: "9CA3AF", panel: "F3F4F6",
  n2: "2A78D6", n2Light: "9CC0EC", air: "EB6834", airLight: "F5B699", grey: "8A8F98",
};
const STATUS = {
  good: { color: "0CA30C", tint: "E2F3E2", label: "양호" },
  fair: { color: "FAB219", tint: "FEF2D3", label: "보통" },
  poor: { color: "D03B3B", tint: "F8DEDE", label: "미흡" },
  reliable: { color: "0CA30C", tint: "E2F3E2", label: "신뢰" },
  conditional: { color: "FAB219", tint: "FEF2D3", label: "조건부" },
  endpoint: { color: "EC835A", tint: "FCE6DC", label: "종점 의존" },
  remeasure: { color: "D03B3B", tint: "F8DEDE", label: "재측정" },
};
const W = 13.333, MX = 0.5, CW = 12.333;

function makeDeck({ title, subject, footer }) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.author = "KICT 중공층 외단열 화재 분석";
  pres.title = title;
  pres.subject = subject;
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;

  pres.defineSlideMaster({
    title: "TITLE_DARK",
    background: { color: THEME.colors.dk1 },
    objects: [
      { placeholder: { options: { name: "kicker", type: "body", x: 0.8, y: 1.3, w: 11.7, h: 0.4, fontSize: 14, bold: true, color: C.background2, margin: 0, align: "left", valign: "middle" }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 1.8, w: 11.7, h: 1.45, fontSize: 36, bold: true, color: C.background1, margin: 0, align: "left", valign: "top" }, text: "" } },
      { placeholder: { options: { name: "subtitle", type: "body", x: 0.8, y: 3.35, w: 11.7, h: 0.8, fontSize: 16, color: C.background2, margin: 0, align: "left", valign: "top" }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: THEME.colors.lt1 },
    margin: [0.5, 0.5, 0.6, 0.5],
    objects: [
      { placeholder: { options: { name: "kicker", type: "body", x: MX, y: 0.3, w: 9.0, h: 0.3, fontSize: 12, bold: true, color: C.text2, margin: 0, align: "left", valign: "middle" }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.68, w: CW, h: 0.92, fontSize: 20, bold: true, color: C.text1, margin: 0, align: "left", valign: "top" }, text: "" } },
      { text: { text: footer, options: { x: MX, y: 7.05, w: 8.0, h: 0.26, fontSize: 9, color: C.text2, margin: 0, valign: "middle" } } },
      { text: { text: "[", options: { x: 12.3, y: 7.05, w: 0.12, h: 0.26, fontSize: 10, color: C.text2, margin: 0, align: "left", valign: "middle" } } },
      { text: { text: "]", options: { x: 12.71, y: 7.05, w: 0.12, h: 0.26, fontSize: 10, color: C.text2, margin: 0, align: "right", valign: "middle" } } },
    ],
    slideNumber: { x: 12.3, y: 7.05, w: 0.53, h: 0.26, fontSize: 10, color: HEX.ink2, align: "center", margin: [0, 0, 0, 0] },
  });

  const fmt = (v, n = 2) => (v === null || v === undefined ? "—" : (Math.round((Number(v) + Math.sign(v) * Number.EPSILON) * 10 ** n) / 10 ** n).toFixed(n));
  // slide text writes T_onset, r_max, N₂, f₄₀₀ ... as base + real subscript (Unicode subscript digits render tiny or
  // fall back to another font); speaker notes keep the plain spelling
  const SUBS = { T_onset: ["T", "onset"], T_peak: ["T", "peak"], "T_α50": ["T", "α50"], "T_α90": ["T", "α90"], "T_α5": ["T", "α5"],
    T_lo: ["T", "lo"], r_max: ["r", "max"], W_ref: ["W", "ref"], W_end: ["W", "end"], c_p: ["c", "p"], "N₂": ["N", "2"], "O₂": ["O", "2"],
    "f₄₀₀": ["f", "400"], "SG₁₀₀": ["SG", "100"], "t₆₀₀": ["t", "600"] };
  const SUB_RE = new RegExp("(?<![A-Za-z_])(" + Object.keys(SUBS).sort((a, b) => b.length - a.length).join("|") + ")", "g");
  function rich(str, base = {}) {
    if (typeof str !== "string" || !SUB_RE.test(str)) return str;
    SUB_RE.lastIndex = 0;
    if (str.includes("\n")) {
      // explicit line breaks become breakLine runs; a "\n" inside one run of several would push the next run to a new line
      const lines = str.split("\n"), out = [];
      lines.forEach((ln, i) => {
        const r = rich(ln, base), runs = typeof r === "string" ? [{ text: r, options: { ...base } }] : r;
        if (i < lines.length - 1) runs[runs.length - 1].options = { ...runs[runs.length - 1].options, breakLine: true };
        out.push(...runs);
      });
      return out;
    }
    const out = [];
    let last = 0, mm;
    while ((mm = SUB_RE.exec(str)) !== null) {
      if (mm.index > last) out.push({ text: str.slice(last, mm.index), options: { ...base } });
      const [b, sub] = SUBS[mm[1]];
      out.push({ text: b, options: { ...base } }, { text: sub, options: { ...base, subscript: true } });
      last = mm.index + mm[0].length;
    }
    if (last < str.length) out.push({ text: str.slice(last), options: { ...base } });
    return out;
  }
  // split runs that contain subscript tokens; paragraph options stay on the first piece, breakLine on the last
  function expandRuns(runs) {
    const out = [];
    runs.forEach((r) => {
      const parts = rich(r.text, r.options || {});
      if (typeof parts === "string") return out.push(r);
      parts.forEach((p, i) => {
        const o = { ...p.options };
        if (i > 0) { delete o.bullet; delete o.paraSpaceAfter; }
        if (i < parts.length - 1) delete o.breakLine;
        out.push({ text: p.text, options: o });
      });
    });
    return out;
  }
  const richText = (t) => (Array.isArray(t) ? expandRuns(t) : rich(t));
  const LANG = "ko-KR";
  const richRows = (rows) => rows.map((row) => row.map((c) => (c && typeof c === "object" && "text" in c ? { ...c, text: richText(c.text), options: { lang: LANG, ...(c.options || {}) } } : c)));
  const tab = (rows, key) => rows.find((r) => Object.entries(key).every(([k, v]) => r[k] === v));
  const SER_STYLE = new Map(); // series name -> { dash, noLine, noMarker }, applied in finish()

  function content(section, kicker, title) {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
    s.addText(richText(kicker), { placeholder: "kicker", lang: LANG });
    s.addText(richText(title), { placeholder: "title", lang: LANG });
    return s;
  }
  function text(s, t, o) {
    s.addText(richText(t), { isTextBox: true, margin: 0, color: C.text1, fontSize: 14, valign: "top", fit: "none", lang: LANG, ...o });
  }
  // rotated (270°) axis title whose visible footprint is centred on (cx, cy) and runs len inches along the axis
  function vlabel(s, t, cx, cy, len) {
    text(s, t, { x: cx - len / 2, y: cy - 0.2, w: len, h: 0.4, fontSize: 11, color: C.text2, align: "center", valign: "middle", rotate: 270 });
  }
  function table(s, rows, o) { s.addTable(richRows(rows), { lang: LANG, ...o }); }
  // paragraphs: each item is a string or an array of [text, options] runs; bullets unless o.bullet === false
  function paras(items, o = {}) {
    const out = [];
    items.forEach((it, i) => {
      const runs = Array.isArray(it) ? it : [[it, {}]];
      runs.forEach(([t, ro], j) => {
        const op = { ...(ro || {}) };
        if (j === 0) {
          if (o.bullet !== false) op.bullet = { indent: 13 };
          op.paraSpaceAfter = o.space === undefined ? 5 : o.space;
        }
        if (j === runs.length - 1 && i < items.length - 1) op.breakLine = true;
        out.push({ text: t, options: op });
      });
    });
    return out;
  }
  function badge(s, n, x, y, label, w = 5.0) {
    s.addShape(pres.ShapeType.ellipse, { x, y, w: 0.3, h: 0.3, fill: { color: C.text1 }, line: { color: C.text1, width: 0.5 }, objectName: `panel ${n}` });
    s.addText(String(n), { x, y, w: 0.3, h: 0.3, fontSize: 11, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
    if (label) text(s, label, { x: x + 0.42, y: y - 0.02, w, h: 0.34, fontSize: 15, bold: true, valign: "middle" });
  }
  function legend(s, entries, x, y) {
    let cx = x;
    entries.forEach((e) => {
      if (e.kind === "box") s.addShape(pres.ShapeType.rect, { x: cx + 0.08, y: y + 0.06, w: 0.2, h: 0.16, fill: { color: e.color }, line: { color: e.color, width: 0.5 } });
      else if (e.kind === "dot") s.addShape(pres.ShapeType.ellipse, { x: cx + 0.1, y: y + 0.05, w: 0.17, h: 0.17, fill: { color: e.color }, line: { color: e.color, width: 0.5 } });
      else s.addShape(pres.ShapeType.line, { x: cx, y: y + 0.14, w: 0.36, h: 0, line: { color: e.color, width: 2, dashType: e.dash ? "dash" : "solid" } });
      text(s, e.label, { x: cx + 0.44, y, w: e.w, h: 0.28, fontSize: 11, color: C.text2, valign: "middle" });
      cx += 0.44 + e.w + 0.2;
    });
  }
  function chip(s, x, y, w, st, label) {
    const S = STATUS[st];
    s.addShape(pres.ShapeType.roundRect, { x, y, w, h: 0.3, rectRadius: 0.06, fill: { color: S.tint }, line: { color: S.color, width: 1 } });
    s.addShape(pres.ShapeType.ellipse, { x: x + 0.1, y: y + 0.085, w: 0.13, h: 0.13, fill: { color: S.color }, line: { color: S.color, width: 0.5 } });
    text(s, label || S.label, { x: x + 0.3, y, w: w - 0.34, h: 0.3, fontSize: 11, bold: true, valign: "middle" });
  }
  // shared chart text style; a fresh object per chart because pptxgenjs mutates options
  function axisStyle() {
    return {
      catAxisLabelColor: HEX.ink2, valAxisLabelColor: HEX.ink2, catAxisLabelFontSize: 10, valAxisLabelFontSize: 10,
      catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
      valGridLine: { color: HEX.grid, size: 0.75 }, catGridLine: { style: "none" },
      catAxisLineShow: true, catAxisLineColor: HEX.axis, valAxisLineShow: false,
      titleFontFace: "+mn-lt", titleFontSize: 12, titleColor: HEX.ink, titleBold: true,
      dataLabelFontFace: "+mn-lt", dataLabelFontSize: 10, dataLabelColor: HEX.ink,
      legendFontFace: "+mn-lt", legendFontSize: 10, legendColor: HEX.ink2, lang: LANG,
    };
  }
  // maps data values to slide inches for a chart drawn with a manual inner plot area L (fractions of the chart frame)
  function plotMap(ch) {
    const px = ch.x + ch.w * ch.L.x, py = ch.y + ch.h * ch.L.y, pw = ch.w * ch.L.w, ph = ch.h * ch.L.h;
    const fy = (v) => (ch.yLog ? (Math.log10(v) - Math.log10(ch.yMin)) / (Math.log10(ch.yMax) - Math.log10(ch.yMin)) : (v - ch.yMin) / (ch.yMax - ch.yMin));
    return { px, py, pw, ph, X: (v) => px + (pw * (v - ch.xMin)) / (ch.xMax - ch.xMin), Y: (v) => py + ph * (1 - fy(v)), cat: (i, n) => px + (pw * (i + 0.5)) / n };
  }
  // XY line chart (native scatter with lines, no markers)
  function curves(s, ch, xs, series, extra = {}) {
    series.forEach((sr) => sr.style && SER_STYLE.set(sr.name, sr.style));
    s.addChart(pres.ChartType.scatter, [{ name: "X", values: xs }, ...series.map((sr) => ({ name: sr.name, values: sr.ys, labels: sr.labels }))], {
      x: ch.x, y: ch.y, w: ch.w, h: ch.h, ...axisStyle(), layout: { ...ch.L },
      chartColors: series.map((sr) => sr.color), lineSize: 1.5, lineDataSymbol: "none",
      showTitle: !!ch.title, title: ch.title, showLegend: false,
      catAxisMinVal: ch.xMin, catAxisMaxVal: ch.xMax, catAxisMajorUnit: ch.xMajor,
      valAxisMinVal: ch.yMin, valAxisMaxVal: ch.yMax, valAxisMajorUnit: ch.yMajor, valAxisCrossesAt: ch.yMin,
      ...extra,
    });
  }
  // horizontal reference line across a plot area, label at the right end inside the plot
  function refLine(s, m, yv, label, color) {
    const y = m.Y(yv);
    s.addShape(pres.ShapeType.line, { x: m.px, y, w: m.pw, h: 0, line: { color, width: 1.25, dashType: "dash" } });
    if (label) text(s, label, { x: m.px + m.pw + 0.06, y: y - 0.2, w: 0.95, h: 0.4, fontSize: 10, color: C.text2, valign: "middle" });
  }
  const CAT = { EPS_C: "EPS\n가연성", EPS_SC: "EPS\n준불연", PIR_C: "PIR(PUR)\n가연성", PIR_SC: "PIR(PUR)\n준불연", PF_out: "PF\n외부용", PF_in: "PF\n내부용" };
  function statusOf(v) { return STATUS[v] || STATUS.poor; }

  async function finish(file) {
    const zip = await JSZip.loadAsync(fs.readFileSync(file));
    const slots = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
    const scheme = `<a:clrScheme name="${THEME.name}">` + slots.map((k) => `<a:${k}><a:srgbClr val="${THEME.colors[k]}"/></a:${k}>`).join("") + "</a:clrScheme>";
    const tp = "ppt/theme/theme1.xml";
    zip.file(tp, (await zip.file(tp).async("string")).replace(/<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/, () => scheme)
      .replace(/(<a:(?:theme|fontScheme)\b[^>]*?\bname=")[^"]*"/g, (_, head) => `${head}${THEME.name}"`));
    const styleSeries = (ser) => {
      const nm = ((ser.match(/<c:tx>[\s\S]*?<c:v>([\s\S]*?)<\/c:v>/) || [])[1] || "").replace(/&amp;/g, "&");
      const st = SER_STYLE.get(nm);
      if (!st) return ser;
      let out = ser;
      if (st.dash) out = out.replace('<a:prstDash val="solid"/>', `<a:prstDash val="${st.dash}"/>`);
      if (st.noLine) out = out.replace(/(<c:spPr>[\s\S]*?)<a:ln w="[^"]*"[^>]*>[\s\S]*?<\/a:ln>/, "$1<a:ln><a:noFill/></a:ln>");
      if (st.noMarker) out = out.replace(/<c:marker>[\s\S]*?<\/c:marker>/, '<c:marker><c:symbol val="none"/></c:marker>');
      return out;
    };
    const firstPPr = (xml) => xml.replace(/<a:p>([\s\S]*?)<\/a:p>/g, (whole, inner) => {
      let run = false;
      return "<a:p>" + inner.replace(/<a:pPr\b[^>]*\/>|<a:pPr\b[^>]*>[\s\S]*?<\/a:pPr>|<a:r>|<a:br\/>|<a:fld\b/g, (tok) => {
        if (!tok.startsWith("<a:pPr")) { run = true; return tok; }
        return run ? "" : tok;
      }) + "</a:p>";
    });
    for (const name of Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))) {
      zip.file(name, firstPPr(await zip.file(name).async("string")));
    }
    for (const name of Object.keys(zip.files).filter((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n))) {
      let x = await zip.file(name).async("string");
      x = x.replace(/<c:(scatterChart|lineChart)>[\s\S]*?<\/c:\1>/g, (blk) => blk.replace(/<c:ser>[\s\S]*?<\/c:ser>/g, styleSeries));
      x = x.replace(/<c:pt idx="\d+"><c:v><\/c:v><\/c:pt>/g, "");
      zip.file(name, x);
    }
    fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  }

  return {
    pres, C, THEME, HEX, STATUS, W, MX, CW, LANG, CAT, SER_STYLE,
    fmt, rich, richText, richRows, tab, statusOf,
    content, text, vlabel, table, paras, badge, legend, chip, axisStyle, plotMap, curves, refLine, finish,
  };
}

module.exports = { makeDeck, THEME, HEX, STATUS, W, MX, CW };
