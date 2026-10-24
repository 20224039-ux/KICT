#!/usr/bin/env node
// 기축 건축물 외벽 화재위험성 평가 방법(안) 덱 (pptxgenjs, native charts). Reads ebfra_data.json (../ebfra.py).
//
//   python ../ebfra.py && NODE_PATH=<folder holding pptxgenjs> node build_ebfra_deck.js [out.pptx]
"use strict";
const fs = require("fs");
const path = require("path");
const { makeDeck } = require("../../tga_dsc/report/deck_lib");

const D = JSON.parse(fs.readFileSync(path.join(__dirname, "ebfra_data.json"), "utf8"));
const OUT = process.argv[2] || path.join(__dirname, "기축건축물_외벽화재위험성_평가방법.pptx");

const SUBS = {
  p_B: ["p", "B"], p_e2: ["p", "e2"], p_e: ["p", "e"], p_cb: ["p", "cb"], p_L: ["p", "L"], p_s: ["p", "s"], p_FO: ["p", "FO"],
  p_v: ["p", "v"], p_x: ["p", "x"], p_st: ["p", "st"], "λ_res": ["λ", "res"], "λ_com": ["λ", "com"], "λ_car": ["λ", "car"],
  "λ_adj": ["λ", "adj"], "τ_L": ["τ", "L"], T_FS: ["T", "FS"], S_FS: ["S", "FS"], S_e: ["S", "e"], R_ref: ["R", "ref"], f_i: ["f", "i"],
  q_a: ["q", "a"], q_b: ["q", "b"], q_min: ["q", "min"], "r_max,air": ["r", "max,air"], "f₄₀₀,air": ["f", "400,air"], n_u: ["n", "u"],
  F_ign: ["F", "ign"], F_s: ["F", "s"], "t̂₆₀₀": ["t̂", "600"], "p̂_B": ["p̂", "B"], R1: ["R", "1"], R2: ["R", "2"], R3: ["R", "3"],
  "f_하부": ["f", "하부"], "a_s": ["a", "s"], "q_s": ["q", "s"], p_e0: ["p", "e0"], p_e1: ["p", "e1"], "T_FS·in": ["T", "FS,in"],
  "T_FS·out": ["T", "FS,out"], RSET_res: ["RSET", "res"],
};
const {
  pres, C, HEX, STATUS, MX, CW, LANG, content: content0, text: text0, vlabel, table, paras, badge, legend, plotMap, curves, finish,
} = makeDeck({
  title: "기축 건축물 외벽 화재위험성 평가 방법(안)",
  subject: "국내 기축 건축물을 형태와 관계없이 외벽 구간으로 나누어 평가하는 평가식과 단계별 절차(연구 방법론 제안)",
  footer: "KICT · 기축 건축물 외벽 화재위험성 평가 방법(안) — 평가식 · 5단계 절차 · 예제 3동",
  themeName: "KICT façade risk process",
  subs: SUBS,
});

// ------------------------------------------------------------------------------------------------ data
const SUM = (b) => D.summary.find((s) => s.building === b);
const OPT = (b, o) => D.remediation.find((r) => r.building === b && r.option === o);
const ZN = (b) => D.zones.filter((z) => z.building === b);
const BLD = (b) => D.buildings.find((x) => x.id === b);
const PAR = (k) => D.params.find((p) => p.key === k);
const P = (k) => PAR(k).best;
const UNC = (b) => D.uncertainty.find((u) => u.building === b);
const GK = { 1: "낮음", 2: "보통", 3: "높음", 4: "매우 높음", 5: "심각" };
const GC = { 1: "86C38B", 2: "B9D98C", 3: "F2C14E", 4: "EB8A3A", 5: "D64545" };       // ordinal grade scale (green → red)
const GT = { 1: "E2F3E2", 2: "EEF6DF", 3: "FDF0D5", 4: "FBE3D2", 5: "F8DEDE" };       // tints for table cells / bands
const BC = { A: "E34948", B: "2A78D6", C: "6250D6" };
const TYPE_C = { ETICS_EPS_C: "EDA100", CAV_PF_IN: "E34948", CAV_PIR_SC: "008300", CAV_PF_OUT: "2A78D6", NC: "8A8F98" };

// ------------------------------------------------------------------------------------------------ helpers
const SUPC = { "⁻": "−", "⁰": "0", "¹": "1", "²": "2", "³": "3", "⁴": "4", "⁵": "5", "⁶": "6", "⁷": "7", "⁸": "8", "⁹": "9" };
const SUP_RE = /[⁻⁰¹²³⁴⁵⁶⁷⁸⁹]/;
function supSplit(str, o) {
  const { breakLine, ...rest } = o;
  const out = [];
  const lines = str.split("\n");
  lines.forEach((line, li) => {
    line.split(/([⁻⁰¹²³⁴⁵⁶⁷⁸⁹]+)/).filter((p) => p !== "").forEach((p) => {
      const sup = /^[⁻⁰¹²³⁴⁵⁶⁷⁸⁹]+$/.test(p), opt = { ...rest };
      if (out.length > 0) { delete opt.bullet; delete opt.paraSpaceAfter; }
      if (sup) opt.superscript = true;
      out.push({ text: sup ? [...p].map((c) => SUPC[c]).join("") : p, options: opt });
    });
    if (li < lines.length - 1 && out.length) out[out.length - 1].options.breakLine = true;
  });
  if (breakLine && out.length) out[out.length - 1].options.breakLine = true;
  return out;
}
function S_(t) {
  if (typeof t === "string") return SUP_RE.test(t) ? supSplit(t, {}) : t;
  if (!Array.isArray(t)) return t;
  const objs = t.map((r) => (Array.isArray(r) ? { text: r[0], options: r[1] || {} } : r));
  if (!objs.some((r) => r && typeof r.text === "string" && SUP_RE.test(r.text))) return t;
  return objs.flatMap((r) => (typeof r.text === "string" && SUP_RE.test(r.text) ? supSplit(r.text, r.options || {}) : [r]));
}
const content = (sec, kicker, title) => content0(sec, kicker, S_(title));
const text = (s, t, o) => text0(s, S_(t), o);
const SUPD = ["⁰", "¹", "²", "³", "⁴", "⁵", "⁶", "⁷", "⁸", "⁹"];
function sci(v, d = 2) {
  if (v === 0) return "0";
  let e = Math.floor(Math.log10(Math.abs(v)));
  let m = v / 10 ** e;
  if (Math.abs(Number(m.toFixed(d - 1))) >= 10) { m /= 10; e += 1; }
  return `${m.toFixed(d - 1)}×10${e < 0 ? "⁻" : ""}${String(Math.abs(e)).split("").map((c) => SUPD[+c]).join("")}`;
}
const pct = (v, d = 0) => `${(100 * v).toFixed(d)} %`;
const f2 = (v, d = 2) => Number(v).toFixed(d);
const fx = (v, d = 0) => (v === null || v === undefined ? "—" : Number(v).toFixed(d));
const BORDER = { type: "solid", pt: 0.5, color: "D1D5DB" };
const TH = (t, o = {}) => ({ text: S_(t), options: { bold: true, color: "FFFFFF", fill: { color: HEX.ink }, fontSize: 11, align: "center", valign: "middle", ...o } });
const TD = (t, o = {}) => ({ text: S_(t), options: { fontSize: 11, color: HEX.ink, valign: "middle", ...o } });
const NOTE = (s, t, y, o = {}) => text(s, t, { x: MX, y, w: CW, h: 0.3, fontSize: 10.5, color: C.text2, ...o });
const tup = (it) => (Array.isArray(it) && it.length && it[0] && typeof it[0] === "object" && !Array.isArray(it[0]) ? it.map((r) => [r.text, r.options]) : it);
const para = (items, o) => paras(items.map((it) => tup(S_(it))), o);
function addNotes(s, lines) { s.addNotes(lines.join("\n")); }
function arrow(s, x1, y1, x2, y2, o = {}) {
  s.addShape(pres.ShapeType.line, {
    x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.max(Math.abs(x2 - x1), 0.001), h: Math.max(Math.abs(y2 - y1), 0.001),
    flipH: x2 < x1, flipV: y2 < y1, line: { color: o.color || HEX.ink2, width: o.width || 1.25, endArrowType: o.head === false ? undefined : "triangle", dashType: o.dash ? "dash" : "solid" },
  });
}
function box(s, x, y, w, h, o = {}) {
  s.addShape(o.rect ? pres.ShapeType.rect : pres.ShapeType.roundRect, { x, y, w, h, rectRadius: o.r === undefined ? 0.06 : o.r, fill: { color: o.fill || "FFFFFF", transparency: o.transparency },
    line: { color: o.line || "D1D5DB", width: o.lw || 0.75, dashType: o.dash ? "dash" : "solid" }, objectName: o.name });
}
function card(s, x, y, w, h, st, head, items, o = {}) {
  box(s, x, y, w, h, { fill: STATUS[st].tint, line: STATUS[st].tint, r: 0.08, name: `card ${head}` });
  s.addShape(pres.ShapeType.ellipse, { x: x + 0.2, y: y + 0.21, w: 0.2, h: 0.2, fill: { color: STATUS[st].color }, line: { color: STATUS[st].color, width: 0.5 } });
  text(s, head, { x: x + 0.5, y: y + 0.1, w: w - 0.65, h: 0.42, fontSize: o.headSize || 15, bold: true, valign: "middle" });
  text(s, para(items, { space: o.space || 4 }), { x: x + 0.22, y: y + 0.58, w: w - 0.42, h: h - 0.7, fontSize: o.size || 12 });
}
function gradeCell(t, g, o = {}) { return TD(t, { align: "center", bold: true, fill: { color: GT[g] }, ...o }); }
const labelOf = (x) => x.label.replace("(", " (").replace(")", ")");
// grade bands drawn behind a log-y chart (the plot area is transparent)
function gradeBands(s, mp, yMin, yMax, labels = true) {
  const edges = [yMin, P("g1"), P("g2"), P("g3"), P("g4"), yMax];
  for (let g = 1; g <= 5; g++) {
    const lo = Math.max(edges[g - 1], yMin), hi = Math.min(edges[g], yMax);
    if (hi <= lo) continue;
    s.addShape(pres.ShapeType.rect, { x: mp.px, y: mp.Y(hi), w: mp.pw, h: mp.Y(lo) - mp.Y(hi), fill: { color: GT[g] }, line: { color: GT[g], width: 0 }, objectName: `grade band ${g}` });
    if (labels) text(s, `${g}등급`, { x: mp.px + mp.pw + 0.05, y: (mp.Y(hi) + mp.Y(lo)) / 2 - 0.13, w: 0.7, h: 0.26, fontSize: 9.5, color: C.text2 });
  }
}

const A_ = SUM("A"), B_ = SUM("B"), C_ = SUM("C");
const share = (b) => {
  const z = ZN(b), R = SUM(b).R;
  return { ie1: z.reduce((a, x) => a + x.r1, 0) / R, ie2: z.reduce((a, x) => a + x.r2, 0) / R, ie3: z.reduce((a, x) => a + x.r3, 0) / R, z, R };
};
const evP = (lev) => D.evidence_mc.find((e) => e.level === lev);
const pGe3 = (e) => e.P_g3 + e.P_g4 + e.P_g5;

// ================================================================================================ 1 title
pres.addSection({ title: "요약" });
{
  const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "요약" });
  s.addText("외벽 화재위험성평가 · 기축 건축물 · 연구 방법론(안) · 2026-10-24", { placeholder: "kicker", lang: LANG });
  s.addText("기축 건축물 외벽 화재위험성\n평가 방법(안)", { placeholder: "title", lang: LANG });
  s.addText("건물 형태와 관계없이 외벽을 '구간'으로 나누어 하나의 평가식으로 계산 — 사진 · 현장에서 시료 시험 · 정밀 평가로 갈수록 정밀해지는 5단계",
    { placeholder: "subtitle", lang: LANG });
  const facts = [
    ["평가식 1개", "R = Σ 빈도 × 도달 확률 × 노출 인원", "→ 개인 위험 r → 1–5등급"],
    ["5단계 절차", "선별 → 현장 → 시료 → 정밀 → 판정", "재료 인자 증거 수준 E0–E3"],
    ["예제 3동", "필로티 · 아파트 · 주상복합", `${A_.grade} · ${B_.grade} · ${C_.grade}등급 · 대책별 등급 변화 계산`],
  ];
  facts.forEach(([big, lab, sub], i) => {
    const x = 0.8 + i * 3.95;
    s.addShape(pres.ShapeType.roundRect, { x, y: 4.55, w: 3.7, h: 1.65, rectRadius: 0.08, fill: { color: C.background1, transparency: 90 }, line: { color: C.background2, width: 0.75 } });
    text(s, big, { x: x + 0.25, y: 4.7, w: 3.3, h: 0.6, fontSize: 24, bold: true, color: C.background1 });
    text(s, lab, { x: x + 0.25, y: 5.32, w: 3.3, h: 0.32, fontSize: 12.5, bold: true, color: C.background1 });
    text(s, sub, { x: x + 0.25, y: 5.66, w: 3.3, h: 0.32, fontSize: 11, color: C.background2 });
  });
  text(s, "함께 드리는 파일: 기축건축물_외벽화재위험성_평가계산서.xlsx (평가식을 셀 수식으로 계산 · 예제 3동 · 빈 양식 · 대책 스위치)",
    { x: 0.8, y: 6.65, w: 11.7, h: 0.3, fontSize: 11, color: C.background2 });
  addNotes(s, [
    "[스크립트]",
    "- 목표: 국내 기축 건축물을 형태와 관계없이 외벽 화재위험 측면에서 평가하는 절차(연구 방법론).",
    "- 핵심: 외벽을 구간으로 나누면 필로티, 저층 상가 + 고층 주거, 측벽, 커튼월 같은 어떤 형태도 같은 평가식 하나로 계산됩니다.",
    "- 재료 시험(콘·TGA·FT-IR·중규모)은 평가식의 재료 입력을 '보수적 기본값'에서 '측정값'으로 바꾸는 단계로 들어갑니다.",
  ]);
}

// ================================================================================================ 2 summary
{
  const s = content("요약", "요약 — 무엇을 계산하고, 무엇이 나오나",
    "외벽을 구간으로 나눠 '연간 노출 인원'을 계산하고 자택 화재 위험과 비교해 1–5등급 —\n재료 시험은 증거 수준을 올려 등급을 확정하는 단계로 들어감");
  const cw = (CW - 0.25) / 2, ch = 2.15;
  card(s, MX, 1.8, cw, ch, "good", "평가식: 형태가 달라도 하나", [
    "R = Σ구간 [실내 화재 분출 + 외벽 하단 화원 + 하부 구간 분출]",
    "각 항 = 빈도 × 위층 도달 확률 × 그 층에서 피난 못 한 재실자",
    "필로티 · 저층 상가 · 측벽 · 커튼월 = 구간의 조합으로 입력",
  ], { size: 12 });
  card(s, MX + cw + 0.25, 1.8, cw, ch, "good", "판정: 개인 위험 ÷ 자택 화재 위험", [
    "r = (R ÷ 재실자) ÷ λ_res, 경계 1 % · 10 % · 같은 수준 · 10배 → 1–5등급",
    "표시 M(노출 1건당 10명 이상) · A(급발 연소)",
    "대책(차단재 · 확산방지구조 · 하단 보호 · 스프링클러 · 교체)도 같은 식으로 재계산",
  ], { size: 12 });
  const e0 = evP("E0"), e3 = evP("E3");
  card(s, MX, 1.8 + ch + 0.2, cw, ch, "conditional", "재료 인자: 시험으로 확정", [
    "E0 사진 · 허가 시점 → 보수값 / E1 현장 확인",
    "E2 시료: FT-IR 식별 → 콘 등급 → TGA t̂₆₀₀ · p̂_B / E3 시스템 시험 · FDS",
    `예제 B: 시료 · 시험으로 3등급 이상 확률 ${pct(pGe3(e0))} → ${pct(pGe3(e3))}`,
  ], { size: 12 });
  card(s, MX + cw + 0.25, 1.8 + ch + 0.2, cw, ch, "fair", "예제 3동(가상)", [
    `필로티 다세대(드라이비트) ${A_.label} → 차단재 + 확산방지구조 + 하단 보호 ${OPT("A", "차단재 + 확산방지구조 + 하단 보호").grade}등급`,
    `15층 아파트(중공층 PIR 준불연) ${B_.label} → 확산방지구조 ${OPT("B", "화재확산방지구조").grade}등급`,
    `주상복합(PE 심재 복합패널) ${C_.label} → 교체만 효과(${OPT("C", "외장재 불연 교체").grade}등급)`,
  ], { size: 12 });
  box(s, MX, 6.42, CW, 0.48, { fill: HEX.ink, line: HEX.ink, r: 0.08, name: "caveat" });
  text(s, [{ text: "주의  ", options: { bold: true } }, { text: "빈도 · 분기확률 · 시간은 예시값(근거 표시) — 수치는 방법 시연용이며, 국내 통계 · 시험 DB로 바꾸면 같은 절차로 갱신됩니다.", options: {} }],
    { x: MX + 0.25, y: 6.42, w: CW - 0.5, h: 0.48, fontSize: 12, color: C.background1, valign: "middle" });
  addNotes(s, [
    "[스크립트]",
    "- 평가 결과는 '이 건물에서 외벽을 통해 번지는 화재로 1년에 몇 명이 노출되는가'입니다. 재실자 1인당으로 나누고 자택 화재 위험과 비교해 등급을 매깁니다.",
    "- 재료 시험은 별도 평가가 아니라, 평가식에 들어가는 재료 인자를 정하는 단계입니다.",
  ]);
}

// ================================================================================================ 3 need
pres.addSection({ title: "배경" });
{
  const s = content("배경", "왜 필요한가 — 국내 기축 건물과 해외 기존 건물 평가",
    "국내 기축 건물의 외벽은 허가 시점 기준을 따르고, 건물별 위험을 따지는 공통 절차가 없음 —\n해외의 기존 건물용 평가도 판단(정성) 위주라 재료 시험을 계산에 쓰지 않음");
  const cw = (CW - 0.25) / 2;
  card(s, MX, 1.85, cw, 3.95, "poor", "국내", [
    "주거용 필로티 건물 약 28만 동 중 약 22만 동(78 %)에 가연성 외장재 — 정부 대책(2025년 9월 보도)의 지원 우선순위는 필로티 주차장 면적 · 1층 방화구획 여부",
    "외벽 마감재료 기준은 허가 시점마다 다름: 2019.11.7 전 신청은 '6층 이상 또는 22 m 이상'만 대상 → 그보다 낮은 건물은 가연 외장재 가능",
    "2025년 7월 광명 아파트 필로티 주차장 화재: 천장 단열재로 급격히 확산, 스프링클러 없음(보도)",
    "드라이비트 · 복합패널 연구는 있으나 기축 건물 외벽 공통 평가 절차는 찾지 못함",
  ], { size: 12, space: 6 });
  card(s, MX + cw + 0.25, 1.85, cw, 3.95, "conditional", "해외", [
    "영국 PAS 9980: 기존 공동주택 외벽 화재위험 평가(FRAEW) — 단계별 절차와 위험 수준 · 대책, 2026년 9월 개정판 · 전문가 판단 위주",
    "NFPA 연구재단 EFFECT(2018): 기존 고층 건물을 A–E로 줄 세우는 정성 도구 — 재료는 범주, 실험실 시험은 식별용",
    "NFPA 551: 화재위험성평가를 검토하는 지침 — 방법은 정하지 않고 문서 · 검토 기준만 제시",
    "EFFECT 후속 과제(13장): 재료 조합 점수를 시험 자료로 뒷받침하고 시험 DB를 갖출 것 — 이 방법에서는 시험값이 평가식에 바로 들어감",
  ], { size: 12, space: 6 });
  box(s, MX, 5.98, CW, 0.62, { fill: HEX.ink, line: HEX.ink, r: 0.08, name: "aim" });
  text(s, [{ text: "이 방법(안)  ", options: { bold: true } }, { text: "국내 기준 이력 · 통계를 쓰고, 어떤 형태든 구간으로 계산하며, 재료 시험(콘 · TGA · FT-IR · 중규모)을 평가식의 입력으로 쓰는 정량 절차", options: {} }],
    { x: MX + 0.25, y: 5.98, w: CW - 0.5, h: 0.62, fontSize: 12.5, color: C.background1, valign: "middle" });
  NOTE(s, "출처: 쿠키뉴스 · 매일신문(2025.9, 필로티 대책), 서울신문(2025.7, 광명 화재), 건축법 시행령 2019.8.6 개정(대통령령 제30030호), BSI PAS 9980, Lamont & Ingolfsson(2018).", 6.68, { fontSize: 9.5 });
  addNotes(s, [
    "[스크립트]",
    "- 기축 건물은 지을 때의 기준을 따르므로, 지금 기준으로는 가연 외장재가 허용되지 않는 건물에도 남아 있습니다. 대상이 수십만 동이라 '어디부터 무엇을 보강할지'를 정할 근거가 필요합니다.",
    "- 해외에도 기존 건물 외벽 평가 절차는 있지만, 판단 위주이고 재료 시험을 계산에 쓰지 않습니다.",
  ]);
}

// ================================================================================================ 4 process
pres.addSection({ title: "절차 · 평가식" });
{
  const s = content("절차 · 평가식", "평가 절차 — 5단계, 단계마다 같은 평가식으로 다시 계산",
    "선별에서 정밀 평가로 갈수록 재료 인자의 근거(증거 수준)만 바뀌고 식은 같음 —\n그래서 어느 단계에서 멈추든 같은 척도의 등급이 나옴");
  const steps = [
    ["0", "선별", "건축물대장 · 사진으로 여러 동을 한꺼번에", "허가일 · 층수 · 높이 · 용도 · 로드뷰", "E0 보수값", "1차 등급 · 우선순위"],
    ["1", "현장 조사", "구간 · 외벽 구성 · 필로티 · 외부 화원 확인", "현장 조사표", "E1 유형 확인", "구간 입력 · 등급"],
    ["2", "시료 시험", "구간 유형별 시료 채취", "FT-IR → 콘 → TGA", "E2 시료 측정값", "재료 인자 · 등급"],
    ["3", "정밀 평가", "등급이 높거나 경계인 건물", "중규모 · 실대형 시험, FDS", "E3 시스템 · s", "확정 등급 · 대책 설계"],
    ["4", "판정 · 대책", "등급 · 표시 · 대책별 감소량", "대책 스위치", "—", "우선순위 · 문서(NFPA 551)"],
  ];
  const gap = 0.3, w = (CW - 4 * gap) / 5, y = 1.9, h = 3.6;
  steps.forEach(([n, t, what, data, ev, outp], i) => {
    const x = MX + i * (w + gap);
    const hi = i === 2;
    box(s, x, y, w, h, { fill: hi ? "E2F3E2" : "F3F4F6", line: hi ? STATUS.good.color : "F3F4F6", lw: hi ? 1.5 : 0.5, r: 0.08, name: `step ${n}` });
    badge(s, n, x + 0.15, y + 0.15);
    text(s, t, { x: x + 0.55, y: y + 0.1, w: w - 0.6, h: 0.4, fontSize: 15, bold: true, valign: "middle" });
    const rows = [["하는 일", what], ["자료", data], ["재료 인자", ev], ["산출", outp]];
    rows.forEach(([k, v], r) => {
      text(s, k, { x: x + 0.18, y: y + 0.62 + r * 0.74, w: w - 0.3, h: 0.24, fontSize: 9.5, color: C.text2, bold: true });
      text(s, v, { x: x + 0.18, y: y + 0.85 + r * 0.74, w: w - 0.3, h: 0.5, fontSize: 11.5 });
    });
    if (i < steps.length - 1) arrow(s, x + w + 0.02, y + h / 2, x + w + gap - 0.02, y + h / 2, { width: 1.75 });
  });
  box(s, MX, 5.7, CW, 0.62, { fill: "FFFFFF", line: HEX.ink, lw: 1.5, r: 0.08, name: "same formula" });
  text(s, [{ text: "같은 평가식  ", options: { bold: true } }, { text: "R = Σ구간 [실내 화재 분출 + 하단 화원 + 하부 구간 분출] — 재료 인자(t₆₀₀ · p_B · A)의 근거만 E0 → E3로 바뀜", options: {} }],
    { x: MX + 0.25, y: 5.7, w: CW - 0.5, h: 0.62, fontSize: 13, valign: "middle", align: "center" });
  NOTE(s, "2단계 강조: 이 연구의 시험 역량(콘 · TGA–DSC · FT-IR · 3 m 중공층)이 들어가는 자리. 2 · 3단계는 0 · 1단계 등급이 높거나 판정이 갈리는 건물에만 적용.", 6.45, { h: 0.45 });
  addNotes(s, [
    "[스크립트]",
    "- 0단계는 서류와 사진만으로 많은 건물을 평가합니다. 재료는 허가 시점 기준으로 보수적으로 가정합니다.",
    "- 1단계 현장 조사에서 구간과 입력을 확정하고, 2단계에서 시료를 시험해 재료 인자를 측정값으로 바꿉니다.",
    "- 3단계는 시스템 시험과 시뮬레이션으로, 판정이 경계에 걸린 건물만 합니다. 어느 단계에서든 같은 식이라 등급이 같은 척도입니다.",
  ]);
}

// ================================================================================================ 5 formula
{
  const s = content("절차 · 평가식", "평가식 — 한 장으로",
    "외벽이 화염에 노출되는 빈도 × 위층 j개 층까지 번질 확률 × 아직 피난 못 한 사람 —\n세 가지 노출 경로를 구간마다 더함");
  box(s, MX, 1.82, CW, 2.45, { fill: "F3F4F6", line: "F3F4F6", r: 0.08, name: "formula" });
  text(s, "R = Σ구간 z ( R1 + R2 + R3 )   [외벽 화재로 1년에 노출되는 재실자, 명/년]", { x: MX + 0.25, y: 1.9, w: CW - 0.5, h: 0.42, fontSize: 17, bold: true });
  const f = [
    ["R1 실내 화재 → 창 분출", "Σ발화층 i  f_i  Σj  c(k) · [ p_e·a·(b·q_min + (1−b)·q_a) + (1 − p_e·a)·b·q_b ]"],
    ["R2 외벽 하단 화원", "Σ화원 s  F_s  Σj  c(k) · p_e2 · a_s(j) · q_s(j)        (필로티 차량, 쓰레기 · 실외기 등)"],
    ["R3 하부 구간 분출화염", "f_하부  Σj  c(k) · p_e2 · a(j) · [ b·q_min + (1−b)·q_a − b·q_b ]        (예: 저층 상가 → 고층 패널)"],
  ];
  f.forEach(([k, v], i) => {
    text(s, k, { x: MX + 0.25, y: 2.42 + i * 0.58, w: 2.6, h: 0.5, fontSize: 12.5, bold: true, valign: "middle" });
    text(s, v, { x: MX + 2.9, y: 2.42 + i * 0.58, w: CW - 3.15, h: 0.5, fontSize: 13, valign: "middle" });
  });
  const defs = [
    ["f_i 발화 빈도", "세대 · 실 화재 × (1 − p_s · 스프링클러) × p_FO × p_v × 창 계수"],
    ["a(j) 외벽 경로", "(1 − p_cb)^j · S_FS(jτ) · [p_B + (1 − p_B) · S_e(jτ)],  τ = s · t₆₀₀ · h / 3 m"],
    ["b(j) 창→창 도약", "p_L^j · (1 − p_s)^(위층 스프링클러 층 수) · S_FS(j · τ_L)"],
    ["q(t) 피난 지연", "P(RSET > t) — 필로티 화재 + 1층 방화구획 없음이면 + p_st"],
    ["c(k) 노출 재실자", "확산 띠 세대 × 세대당 인원 × 실내 유입 (층별)"],
  ];
  const dw = (CW - 4 * 0.2) / 5;
  defs.forEach(([k, v], i) => {
    const x = MX + i * (dw + 0.2);
    box(s, x, 4.42, dw, 1.42, { fill: "FFFFFF", line: "D1D5DB", r: 0.06, name: k });
    text(s, k, { x: x + 0.12, y: 4.48, w: dw - 0.24, h: 0.3, fontSize: 11.5, bold: true });
    text(s, v, { x: x + 0.12, y: 4.8, w: dw - 0.24, h: 1.0, fontSize: 10.5 });
  });
  box(s, MX, 5.98, CW, 0.75, { fill: HEX.ink, line: HEX.ink, r: 0.08, name: "judge" });
  text(s, para([
    [["판정  ", { bold: true }], ["IR = R ÷ 재실자,  r = IR ÷ λ_res(자택 화재 위험) → 1등급 < 0.01 ≤ 2 < 0.1 ≤ 3 < 1 ≤ 4 < 10 ≤ 5", {}]],
    [["기준  ", { bold: true }], ["R_ref = 같은 건물의 외벽을 모두 불연으로 바꿨을 때(창→창 도약만) → 외벽 기여 RR = R / R_ref, ΔR = R − R_ref", {}]],
  ], { bullet: false, space: 2 }), { x: MX + 0.25, y: 6.0, w: CW - 0.5, h: 0.72, fontSize: 12, color: C.background1, valign: "middle" });
  addNotes(s, [
    "[스크립트]",
    "- 이것이 제안하는 평가식입니다. 세 경로: 세대 화재가 창으로 분출하는 경우, 외벽 바로 아래 화재(필로티 차량 등), 아래 구간 창의 화염이 위 구간 외벽을 가열하는 경우.",
    "- 외벽 재료는 a(j)에만 들어갑니다: 층당 확산시간 τ(시험의 t600)와 자립 연소 p_B. 외벽이 불연이면 a = 0이고 창→창 도약 b만 남는데, 이것이 기준 위험도입니다.",
    "- 판정은 1인당 위험을 자택 화재 위험과 비교합니다.",
  ]);
}

// ================================================================================================ 6 zones
{
  const s = content("절차 · 평가식", "어떤 형태든 — 외벽을 '구간'으로 나누는 방법",
    "같은 외벽 시스템 · 같은 노출(창 유무 · 하단 화원)의 수직 구간으로 나누면 —\n필로티 · 저층 상가 + 고층 주거 · 측벽 · 커튼월이 모두 구간의 조합");
  const baseY = 6.35, topY = 2.05, H = baseY - topY;
  const fillOf = (type) => (type === "NC" ? "E5E7EB" : D.types.find((t) => t.key === type).grade === "가연" ? "F8DEDE" : "FDF0D5");
  const ZSHORT = { A1: "전면\n(필로티 위)", A2: "측 · 후면", B1: "세대 창 면", B2: "측벽", C1: "저층 상가", C2: "고층 주거", C3: "측면" };
  const drawBld = (bid, x0, wTot) => {
    const b = BLD(bid), N = b.N, sh = H / N;
    const cols = [...new Set(b.zones.map((z) => z.col))];
    const cw = cols.map((c, i) => (i === 0 ? wTot * 0.68 : wTot * 0.32));
    const cx = cols.map((c, i) => x0 + (i === 0 ? 0 : cw[0]));
    if (b.piloti) {
      s.addShape(pres.ShapeType.rect, { x: x0, y: baseY - sh, w: wTot, h: sh, fill: { color: "FFFFFF" }, line: { color: HEX.ink2, width: 1, dashType: "dash" }, objectName: "piloti" });
      text(s, "필로티 주차", { x: x0, y: baseY - sh, w: wTot, h: sh, fontSize: 10, align: "center", valign: "middle", color: C.text2 });
    }
    b.zones.forEach((z) => {
      const ci = cols.indexOf(z.col), y = baseY - z.t * sh, hh = (z.t - z.b + 1) * sh;
      s.addShape(pres.ShapeType.rect, { x: cx[ci], y, w: cw[ci], h: hh, fill: { color: fillOf(z.type) }, line: { color: HEX.ink2, width: 1 }, objectName: `zone ${z.id}` });
      text(s, `${z.id}\n${ZSHORT[z.id] || z.name}`, { x: cx[ci] + 0.04, y: y + Math.min(0.06, hh / 4), w: cw[ci] - 0.08, h: Math.min(0.62, hh), fontSize: hh > 0.6 ? 10 : 8.5, bold: true, align: "center" });
    });
    text(s, `${b.id}  ${b.name}`, { x: x0 - 0.2, y: 1.72, w: wTot + 0.4, h: 0.3, fontSize: 12.5, bold: true, align: "center" });
    text(s, `${N}층`, { x: x0 - 0.2, y: baseY + 0.03, w: wTot + 0.4, h: 0.25, fontSize: 9.5, color: C.text2, align: "center" });
    s.addShape(pres.ShapeType.line, { x: x0 - 0.15, y: baseY, w: wTot + 0.3, h: 0, line: { color: HEX.ink, width: 1.5 } });
    return { sh, cx, cw };
  };
  // A
  const a = drawBld("A", MX + 0.35, 1.7);
  text(s, "IE2 차량", { x: MX + 0.85, y: baseY - a.sh + 0.06, w: 0.9, h: 0.25, fontSize: 9.5, color: "B42318", bold: true });
  arrow(s, MX + 0.75, baseY - 0.12, MX + 0.75, baseY - a.sh - 0.25, { color: "B42318", width: 2 });
  // B
  const bx = MX + 2.75, bb = drawBld("B", bx, 1.7);
  arrow(s, bx + 0.5, baseY - 6.5 * bb.sh, bx + 0.5, baseY - 8.6 * bb.sh, { color: "B42318", width: 2 });
  text(s, "IE1", { x: bx + 0.56, y: baseY - 8.2 * bb.sh, w: 0.5, h: 0.25, fontSize: 9.5, color: "B42318", bold: true });
  // C
  const cxx = MX + 5.15, cc = drawBld("C", cxx, 1.9);
  const reachK = Math.floor(P("h_reach") / BLD("C").h) + 1;
  const yR = baseY - (reachK - 1) * cc.sh;
  s.addShape(pres.ShapeType.line, { x: cxx - 0.25, y: yR, w: 2.4, h: 0, line: { color: HEX.ink2, width: 1, dashType: "dash" } });
  text(s, `고가사다리 ${P("h_reach")} m`, { x: cxx + 0.05, y: yR - 0.25, w: 1.25, h: 0.22, fontSize: 8.5, color: HEX.ink2 });
  arrow(s, cxx + 0.6, baseY - 2.6 * cc.sh, cxx + 0.6, baseY - 5.2 * cc.sh, { color: "B42318", width: 2 });
  text(s, "IE3", { x: cxx + 0.66, y: baseY - 4.9 * cc.sh, w: 0.5, h: 0.25, fontSize: 9.5, color: "B42318", bold: true });
  // rules
  const rx = MX + 7.6, rw = CW - 7.6;
  card(s, rx, 1.8, rw, 3.35, "good", "구간 나누는 규칙", [
    "외벽 시스템이 바뀌는 높이 · 면에서 나눔(석재 저층부 / 패널 고층부)",
    "창이 있는 면과 거의 없는 면(측벽)은 나눔 — 창 계수 1 · 0.4 · 0",
    "수직으로 이어진 구간은 같은 '열': 위층 도약 · 하부 분출이 이어짐",
    "하단 화원(필로티 차량 · 쓰레기 · 실외기)은 그 위 구간에 붙임",
  ], { size: 11.5 });
  card(s, rx, 5.3, rw, 1.35, "conditional", "건물 단위 입력", [
    "층수 · 층고 · 사다리 도달 높이 · 스프링클러 설치 층 · 필로티(천장 마감, 1층 방화구획)",
  ], { size: 11.5 });
  legend(s, [{ kind: "box", color: "F8DEDE", label: "가연", w: 0.55 }, { kind: "box", color: "FDF0D5", label: "준불연", w: 0.7 }, { kind: "box", color: "E5E7EB", label: "불연", w: 0.55 }], MX + 0.2, 6.68);
  text(s, "예제 3동의 구간 · 건물마다 높이 축척이 다름", { x: MX + 3.7, y: 6.71, w: 3.7, h: 0.26, fontSize: 9.5, color: C.text2 });
  addNotes(s, [
    "[스크립트]",
    "- 구간은 '같은 외벽 시스템, 같은 노출'의 수직 띠입니다. 필로티 건물은 필로티 위 전면과 측·후면, 아파트는 창이 있는 면과 측벽, 주상복합은 저층 상가와 고층 주거 면으로 나눕니다.",
    "- 빨간 화살표가 세 경로입니다: 필로티 차량 화재(IE2), 세대 화재 분출(IE1), 저층 상가 분출이 고층 패널을 가열(IE3). 점선은 고가사다리가 닿는 높이로, 그 위는 외부 진압이 늦습니다.",
  ]);
}

// ================================================================================================ 7 F term
pres.addSection({ title: "평가식의 항" });
{
  const s = content("평가식의 항", "F 항 — 외벽이 화염에 노출되는 빈도",
    "실내 화재는 통계로(2023년 아파트 화재 2,993건), 외부 화원은 자료가 없어 예시값 —\n자료를 넣을 자리와 대체할 출처를 함께 둠");
  const rows = [
    ["IE1 실내 화재 → 창 분출", "f_i = n_u · λ_res · (1 − p_s · 스프링클러(i층)) · p_FO · p_v · 창 계수",
      `λ_res ${sci(P("lam_res"))}/세대·년 = 2023 아파트 화재 2,993건 ÷ 아파트 약 1,260만 호`, "근거 있음"],
    ["", "비주거: n_u · λ_res 대신 면적 · λ_com", `λ_com ${sci(P("lam_com"))}/m²·년 (예시) → 용도별 화재 ÷ 연면적`, "대체 필요"],
    ["", "p_s 스프링클러 작동 · 제어, p_FO 최성기 도달, p_v 창 분출", `p_s ${P("p_s")} · p_FO ${P("p_fo")} · p_v ${P("p_v")} (예시)`, "대체 필요"],
    ["IE2 필로티 차량", "F = λ_car · 주차면 · p_x(천장 가연 0.9 / 불연 0.5)", `λ_car ${sci(P("lam_car"))}/주차면·년 (예시) → 주차 중 차량 화재 통계`, "대체 필요"],
    ["IE2 기타 하단 화원", "F = λ_adj · 개소 · p_x(0.3)", `λ_adj ${sci(P("lam_adj"))}/개소·년 (예시) → 건물 외부 발화 통계`, "대체 필요"],
    ["IE3 하부 구간 분출", "F = 아래 구간 최상층의 f_i", "IE1과 같은 자료", "근거 있음"],
  ];
  const head = [TH("경로"), TH("식"), TH("값 · 출처"), TH("상태")];
  table(s, [head, ...rows.map(([a, b, c, d]) => [TD(a, { bold: true }), TD(b), TD(c, { fontSize: 10.5 }),
    TD(d, { align: "center", bold: true, fill: { color: d === "근거 있음" ? "E2F3E2" : "FDF0D5" } })])],
  { x: MX, y: 1.85, w: CW, colW: [2.3, 4.4, 4.43, 1.2], fontSize: 11, border: BORDER, rowH: [0.36, 0.56, 0.5, 0.5, 0.56, 0.5, 0.46], margin: [0.04, 0.08, 0.04, 0.08] });
  NOTE(s, "스프링클러는 층별 입력(예: 11층 이상만) — 기축 아파트는 허가 시점에 따라 일부 층만 설치된 경우가 많아 층마다 (1 − p_s)를 적용.", 5.5);
  NOTE(s, "통계 출처: 소방청 2023 화재 통계(공동주택 4,868건 · 아파트 2,993건 보도), 인구주택총조사(아파트 2020년 1,166만 호 · 2025년 약 1,330만 호).", 5.85);
  addNotes(s, ["[스크립트]", "- 빈도는 세 경로마다 따로 셉니다. 실내 화재는 국내 통계로 바로 정할 수 있고, 필로티 차량 화재 같은 외부 화원은 통계를 따로 뽑아야 합니다."]);
}

// ================================================================================================ 8 E term
{
  const s = content("평가식의 항", "E 항 — 외벽을 타고 위층에 닿을 확률: 재료 인자가 들어가는 곳",
    "외벽 재료는 층당 확산시간 τ와 자립 연소 p_B로만 들어감 — 가연 외단열은 끝까지,\n준불연 중공층은 1–3개 층, 불연 외벽은 창→창 도약만 남음");
  const lw = 5.6;
  text(s, para([
    [["외벽 경로  ", { bold: true }], ["a(j) = (1 − p_cb)^j · S_FS(jτ) · [p_B + (1 − p_B) · S_e(jτ)]", {}]],
    [["", {}], ["진압 전(S_FS)이고, 분출화염이 남아 있거나(S_e) 스스로 타면(p_B) 다음 층으로", {}]],
    [["창→창 도약  ", { bold: true }], ["b(j) = p_L^j · (1 − p_s)^(위층 스프링클러 층) · S_FS(j · τ_L)", {}]],
    [["진압 시각  ", { bold: true }], [`고가사다리 도달 범위 T_FS 중앙값 ${P("t_fs_lo")}분, 범위 밖 ${P("t_fs_hi")}분(예시)`, {}]],
    [["대책  ", { bold: true }], [`화재확산방지구조 p_cb ${P("p_cb")}/층 경계, 개구부 차단재 p_e ${P("p_e0")} → ${P("p_e1")}`, {}]],
    [["인정 규칙  ", { bold: true }], ["가연재가 바깥 면이거나 패널 심재로 이어진 유형(PE 심재 복합패널 · 샌드위치패널 · 미상 가연)은 확산방지구조 · 차단재 효과를 인정하지 않음(EFFECT와 같은 판단)", {}]],
  ], { bullet: false, space: 7 }), { x: MX, y: 1.85, w: lw, h: 4.7, fontSize: 12 });
  const keys = ["ETICS_EPS_C", "CAV_PF_IN", "CAV_PIR_SC", "CAV_PF_OUT", "NC"];
  const xs = [1, 2, 3, 4, 5, 6, 7, 8];
  const ch = { x: MX + lw + 0.25, y: 1.8, w: CW - lw - 0.25, h: 4.2, L: { x: 0.1, y: 0.05, w: 0.86, h: 0.8 }, xMin: 1, xMax: 8, xMajor: 1, yMin: 0, yMax: 1, yMajor: 0.2 };
  const series = keys.map((k) => ({ name: `reach ${k}`, ys: xs.map((j) => D.reach.find((r) => r.type === k && r.j === j).p), color: TYPE_C[k],
    style: { marker: { symbol: "circle", size: 6, fill: TYPE_C[k], line: TYPE_C[k] } } }));
  curves(s, ch, xs, series, { lineSize: 2.25, valAxisLabelFormatCode: "0%", catAxisLabelFormatCode: "0" });
  const mp = plotMap(ch);
  text(s, "발화층 위 j번째 층", { x: mp.px, y: mp.py + mp.ph + 0.3, w: mp.pw, h: 0.26, fontSize: 11, color: C.text2, align: "center" });
  vlabel(s, "그 층까지 닿을 확률", ch.x + 0.12, mp.py + mp.ph / 2, 2.2);
  const lab = { ETICS_EPS_C: "외단열 미장 EPS(가연)", CAV_PF_IN: "중공층 PF 내부용", CAV_PIR_SC: "중공층 PIR 준불연", CAV_PF_OUT: "중공층 PF 외부용", NC: "불연(도약만)" };
  legend(s, keys.slice(0, 3).map((k) => ({ label: lab[k], color: TYPE_C[k], w: 1.75 })), ch.x + 0.1, 6.12);
  legend(s, keys.slice(3).map((k) => ({ label: lab[k], color: TYPE_C[k], w: 1.75 })), ch.x + 0.1, 6.4);
  text(s, "조건: 실내 화재 분출 · 스프링클러 없음 · 사다리 도달 범위 · 대책 없음 · s = 0.5", { x: ch.x + 0.1, y: 6.7, w: ch.w, h: 0.24, fontSize: 9.5, color: C.text2 });
  addNotes(s, [
    "[스크립트]",
    "- 오른쪽 그래프가 재료의 차이가 평가식에서 드러나는 방식입니다. 가연 외단열은 분출화염에 노출되면 사실상 끝까지 번지고, 준불연 중공층은 한두 개 층, 불연 외벽은 창→창 도약만 남습니다.",
    "- PF 외부용은 확산이 느리지만 자립 연소가 많아 꼬리가 깁니다.",
  ]);
}

// ================================================================================================ 9 C term + grade
{
  const s = content("평가식의 항", "C 항과 판정 — 누가 노출되고, 몇 등급인가",
    "위층에 화재가 닿을 때 아직 피난하지 못한 재실자를 세고, 1인당 위험을 자택 화재 위험(λ_res)과\n비교해 등급을 매김 — 자택 화재 위험의 1 % · 10 % · 같은 수준 · 10배가 경계");
  const lw = 5.4;
  text(s, para([
    [["노출 재실자  ", { bold: true }], ["c(k) = 확산 띠 세대(발화 세대 열 + 이웃 1열) × 세대당 인원 × 실내 유입", {}]],
    [["피난 지연  ", { bold: true }], [`q(t) = P(RSET > t): 주거(취침) 중앙값 ${P("rset_res")}분, 비주거 ${P("rset_com")}분(예시)`, {}]],
    [["필로티 · 계단  ", { bold: true }], [`필로티 화재 + 1층 방화구획 없음 → 계단 연기로 피난 실패 p_st ${P("p_stair")}를 더함`, {}]],
    [["개인 위험  ", { bold: true }], ["IR = R ÷ 재실자(/인 · 년),  r = IR ÷ λ_res", {}]],
    [["표시  ", { bold: true }], ["M: 노출 1건당 노출 인원 C̄ ≥ 10명(대형 피해) · A: 급발(용융 · 가연 전면 화염)", {}]],
  ], { bullet: false, space: 8 }), { x: MX, y: 1.85, w: lw, h: 4.6, fontSize: 12 });
  const rx = MX + lw + 0.3, rw = CW - lw - 0.3;
  const lam = P("lam_res");
  const g = [[1, "r < 0.01", `< ${sci(lam * 0.01)}`, "자택 화재 위험의 1 % 미만"], [2, "0.01 – 0.1", `${sci(lam * 0.01)} –`, "1–10 %"],
    [3, "0.1 – 1", `${sci(lam * 0.1)} –`, "10 % – 같은 수준"], [4, "1 – 10", `${sci(lam)} –`, "같은 수준 – 10배"], [5, "≥ 10", `≥ ${sci(lam * 10)}`, "10배 이상"]];
  table(s, [[TH("등급"), TH("r = IR / λ_res"), TH("IR (/인·년)"), TH("의미")],
    ...g.map(([n, a, b, c]) => [gradeCell(`${n} ${GK[n]}`, n), TD(a, { align: "center" }), TD(b, { align: "center" }), TD(c)])],
  { x: rx, y: 1.85, w: rw, colW: [1.35, 1.45, 1.5, rw - 4.3], fontSize: 11.5, border: BORDER, rowH: [0.36, 0.46, 0.46, 0.46, 0.46, 0.46], margin: [0.04, 0.08, 0.04, 0.08] });
  text(s, para([
    `λ_res = ${sci(lam)}/세대·년(2023 아파트 화재 통계) — 내 집에서 불이 날 연간 확률`,
    "실내 화재가 주된 건물은 λ_res가 분자 · 분모에서 상쇄되어 r이 화재 빈도의 불확실성에 덜 민감함",
    "등급 경계는 제안값 — 사례 · 사회적 합의로 보정할 연구 대상",
  ], { space: 5 }), { x: rx, y: 4.85, w: rw, h: 1.6, fontSize: 11, color: C.text2 });
  addNotes(s, [
    "[스크립트]",
    "- 등급의 기준점은 '내 집에서 불이 날 위험'입니다. 외벽 때문에 노출될 위험이 그 1 %면 1·2등급 경계, 같은 수준이면 3·4등급 경계입니다.",
    "- 실내 화재가 주된 건물에서는 화재 빈도가 분자와 분모에 같이 들어가 상쇄되므로, 통계가 다소 틀려도 등급이 안정적입니다.",
  ]);
}

// ================================================================================================ 10 material factor table
pres.addSection({ title: "재료 인자" });
{
  const s = content("재료 인자", "재료 인자 — 외벽 유형별 t₆₀₀ · p_B · 급발, 시험한 것과 경계값",
    "시험한 중공층 5종은 시험값(E3), 드라이비트 · 복합패널 · 샌드위치패널은 경계값 —\n시험 자료가 없어 가연 경계값(30 s)으로 두고, 시료 · 시스템 시험이 들어오면 바뀜");
  const keys = ["CAV_PIR_SC", "CAV_PF_OUT", "CAV_PF_IN", "CAV_EPS_SC", "CAV_PIR_C", "ETICS_EPS_C", "ACP_PE", "SWP_C", "SC_UNKNOWN", "C_UNKNOWN", "NC"];
  const rows = keys.map((k) => {
    const t = D.types.find((x) => x.key === k);
    const ev = t.evidence;
    return [TD(t.label, { bold: true, fontSize: 10.5 }), TD(t.grade, { align: "center" }), TD(t.t600_s === null ? "—" : fx(t.t600_s), { align: "center" }),
      TD(t.pB === null ? "—" : f2(t.pB, 2), { align: "center" }), TD(t.A ? "A" : "—", { align: "center" }),
      TD(t.path ? (t.cb_ok ? "인정" : "불인정") : "—", { align: "center", fill: { color: t.path && !t.cb_ok ? "F8DEDE" : "FFFFFF" } }),
      TD(ev, { align: "center", bold: true, fill: { color: ev === "E3" ? "E2F3E2" : ev === "E0" ? "FDF0D5" : "F3F4F6" } }),
      TD(t.basis, { fontSize: 9 })];
  });
  table(s, [[TH("외벽 유형", { fontSize: 10.5 }), TH("등급", { fontSize: 10.5 }), TH("t₆₀₀ (s)", { fontSize: 10.5 }), TH("p_B", { fontSize: 10.5 }), TH("급발", { fontSize: 10.5 }),
    TH("확산방지 · 차단재", { fontSize: 10.5 }), TH("근거", { fontSize: 10.5 }), TH("값의 출처", { fontSize: 10.5 })], ...rows],
  { x: MX, y: 1.8, w: CW, colW: [3.0, 0.7, 0.85, 0.65, 0.6, 1.25, 0.65, 4.63], fontSize: 10.5, border: BORDER,
    rowH: [0.34, ...keys.map(() => 0.36)], margin: [0.02, 0.06, 0.02, 0.06] });
  NOTE(s, "t₆₀₀: 3 m 중공층 시험(30 kW 버너)에서 3000 mm 높이가 600 K 오르는 시간 = 층당 확산시간. p_B: 버너 소화 뒤 스스로 탈 확률(시험 횟수로 추정).", 6.3);
  NOTE(s, "외단열 미장(드라이비트) · 복합패널 · 샌드위치패널의 시스템 시험 DB를 만드는 것이 이 방법을 쓰기 위한 첫 연구 과제.", 6.6, { bold: true, color: C.text1 });
  addNotes(s, [
    "[스크립트]",
    "- 재료 인자는 외벽 유형별 표에서 시작합니다. 우리가 시험한 중공층 5종은 시험값이고, 국내 기축 건물에 많은 드라이비트와 복합패널은 아직 시스템 시험 자료가 없어 가연 경계값을 씁니다.",
    "- 그래서 이 방법의 첫 연구 과제는 기축 건물에 많은 외벽 시스템의 시험 DB입니다.",
  ]);
}

// ================================================================================================ 11 sampling (stage 2)
{
  const s = content("재료 인자", "2단계 시료 시험 — FT-IR로 식별, 콘으로 등급, TGA로 t̂₆₀₀ · p̂_B",
    "시료로 재료 인자를 측정값으로 바꿈 — 콘이 가연이면 가연 경계값,\n준불연 탄화형이면 TGA 예측식, 용융형이면 급발 A");
  const steps = [
    ["채취", "구간 유형마다 2개소 이상(개구부 둘레 · 하단 포함), 마감 + 단열재 + 접착층", "fair"],
    ["FT-IR", "고분자(EPS · PIR/PUR · PF · PE 심재) · 난연제(P=O 등) 식별 → 유형 확정", "conditional"],
    ["콘", "KS F ISO 5660-1, 50 kW/m² — 준불연 기준 통과 여부(게이트)", "fair"],
    ["TGA", "공기 · N₂, 5 °C/min — r_max,air · f₄₀₀,air", "good"],
  ];
  const w = (CW - 3 * 0.35) / 4;
  steps.forEach(([h1, b1, st], i) => {
    const x = MX + i * (w + 0.35);
    box(s, x, 1.9, w, 1.6, { fill: STATUS[st].tint, line: STATUS[st].tint, r: 0.08, name: h1 });
    badge(s, i + 1, x + 0.15, 2.02);
    text(s, h1, { x: x + 0.55, y: 1.96, w: w - 0.6, h: 0.4, fontSize: 15, bold: true, valign: "middle" });
    text(s, b1, { x: x + 0.18, y: 2.45, w: w - 0.32, h: 1.0, fontSize: 11.5 });
    if (i < 3) arrow(s, x + w + 0.03, 2.7, x + w + 0.32, 2.7, { width: 1.75 });
  });
  const out = [
    ["콘: 준불연 기준 미달", "가연 경계값(t₆₀₀ 30 s, p_B 0.5, A)", "poor"],
    ["준불연 · 탄화형(f₄₀₀,air ≥ 0.5)", "t̂₆₀₀ = 247.5 + 422.97 / r_max,air,  p̂_B = (f₄₀₀,air − 0.60)/0.10 (0.05–0.95)", "good"],
    ["준불연 · 용융형(f₄₀₀,air < 0.5)", "유형 기본값 + 급발 A (확산식 미적용)", "fair"],
  ];
  out.forEach(([a, b, st], i) => {
    const y = 3.8 + i * 0.72;
    box(s, MX, y, CW, 0.6, { fill: STATUS[st].tint, line: STATUS[st].tint, r: 0.06, name: a });
    text(s, a, { x: MX + 0.2, y, w: 3.6, h: 0.6, fontSize: 12.5, bold: true, valign: "middle" });
    text(s, b, { x: MX + 3.9, y, w: CW - 4.1, h: 0.6, fontSize: 12.5, valign: "middle" });
  });
  text(s, para([
    "TGA 예측식: 준불연 탄화형 3종 · 시험 6회로 보정, 제품을 빼고 맞춘 교차검증에서 확산등급 6/6 일치(integrated_model)",
    "가연성 PIR(PUR)은 TGA로 812 s를 예측했지만 실측 30 s — 콘 게이트가 없으면 위험을 크게 낮게 봄(multiscale_fra)",
    "EFFECT도 시료 · 실험실 시험(콘 · FT-IR · TGA)을 권하지만 식별에만 씀 → 여기서는 평가식의 입력",
  ], { space: 4 }), { x: MX, y: 5.98, w: CW, h: 0.95, fontSize: 11, color: C.text2 });
  addNotes(s, [
    "[스크립트]",
    "- 시료 시험의 순서는 식별, 등급, 지표입니다. FT-IR로 무엇인지 확인하고, 콘으로 준불연인지 가린 뒤, 준불연 탄화형이면 TGA로 층당 확산시간과 자립 연소를 예측합니다.",
    "- 가연성 PIR에서 보았듯 콘 게이트 없이 TGA만 쓰면 위험을 크게 낮게 봅니다.",
  ]);
}

// ================================================================================================ 12 stage 0
pres.addSection({ title: "단계별 절차" });
{
  const s = content("단계별 절차", "0단계 선별 — 건축물대장 · 사진 · 허가 시점 기준으로 E0 기본값",
    "허가(신청) 시점 기준으로 '가연 가능 건물'을 가려 보수값을 둠 —\n기준은 요구일 뿐 시공 사실이 아니므로 0단계 등급은 위로 치우침(현장 · 시료로 낮아짐)");
  const rows = D.history.map((h) => [TD(h.period, { bold: true, align: "center" }), TD(`${h.from || "…"} – ${h.to || "현재"}`, { align: "center", fontSize: 10.5 }),
    TD(h.scope, { fontSize: 10.5 }), TD(h.e0_rule, { fontSize: 10.5 }),
    TD(h.status, { fontSize: 10, align: "center", fill: { color: h.status.includes("필요") ? "FDF0D5" : "E2F3E2" } })]);
  table(s, [[TH("구간"), TH("허가 시점"), TH("외벽 기준 대상"), TH("E0 추정"), TH("확인")], ...rows],
    { x: MX, y: 1.85, w: 8.6, colW: [0.6, 1.5, 2.9, 2.3, 1.3], fontSize: 10.5, border: BORDER, rowH: [0.36, 0.62, 0.62, 0.72, 0.9], margin: [0.03, 0.06, 0.03, 0.06] });
  const rx = MX + 8.85, rw = CW - 8.85;
  card(s, rx, 1.85, rw, 3.25, "good", "건축물대장에서", [
    "용도 · 층수 · 높이 · 구조",
    "허가일(사용승인일 아님)",
    "세대 수 · 연면적 → 재실자",
    "필로티 여부(1층 용도)",
  ], { size: 12 });
  card(s, rx, 5.25, rw, 1.4, "conditional", "사진(로드뷰)에서", ["외벽 유형 · 구간 · 창 비율 · 필로티 주차"], { size: 12 });
  NOTE(s, "2019.11.7 시행 기준은 대통령령 제30030호로 확인 · 2010 · 2015년 경계는 법령 원문(국가법령정보센터)으로 확정 후 사용 — 엑셀 3_허가시점기준에 계산기.", 5.45, { w: 8.6, h: 0.5 });
  NOTE(s, "P3에서도 5층 · 22 m 이하는 화재확산방지구조를 두면 가연 단열재 허용 → 1단계에서 확인.", 6.05, { w: 8.6 });
  addNotes(s, ["[스크립트]", "- 0단계는 서류만으로 많은 건물을 걸러내는 단계입니다. 허가 시점에 외벽 기준이 없던 건물은 가연으로 가정하고, 이후 현장과 시료로 확인합니다."]);
}

// ================================================================================================ 13 stage 1
{
  const s = content("단계별 절차", "1단계 현장 조사 — 조사 항목이 평가식의 어느 입력이 되나",
    "현장 조사표 항목을 평가식 입력과 1:1로 맞춤 —\n구간 나누기 · 외부 화원 · 필로티 · 스프링클러 층이 등급을 가장 크게 바꿈");
  const rows = [
    ["구간 나누기", "외벽 시스템이 바뀌는 높이 · 면, 창이 있는 면 / 측벽", "구간 b–t, 열, 창 계수"],
    ["외벽 구성", "마감재 · 단열재 계열 · 공기층 유무 · 두께, 품질 표시", "외벽 유형(키), 증거 E1"],
    ["개구부", "창 크기 · 간격, 스팬드럴 높이, 발코니", "p_v · 창 계수, p_L"],
    ["층간 · 개구부 처리", "층간 불연 띠 · 캐비티 배리어, 창 둘레 불연 차단재, 슬래브 끝 틈", "확산방지구조(p_cb), 차단재(p_e), 층간 틈"],
    ["외벽 하단 · 필로티", "하단 마감, 필로티 천장 마감(가연/불연), 주차면 수, 1층 방화구획", "하단 보호(p_e2), F 필로티, p_st"],
    ["외부 화원", "쓰레기 · 적재물 · 실외기 · 전기설비 · 인접 건물", "기타 외부 화원 수"],
    ["설비 · 피난", "스프링클러 설치 층, 경보 방식, 계단 수", "스프링클러 시작층, RSET"],
    ["재실자 · 소방", "세대 수 · 용도, 고가사다리 접근 면 · 높이", "c(k), 사다리 도달 높이"],
  ];
  table(s, [[TH("항목"), TH("현장에서 볼 것"), TH("평가식 입력")], ...rows.map(([a, b, c]) => [TD(a, { bold: true }), TD(b), TD(c, { fill: { color: "F3F4F6" } })])],
    { x: MX, y: 1.85, w: CW, colW: [2.2, 6.0, 4.13], fontSize: 11.5, border: BORDER, rowH: [0.36, ...rows.map(() => 0.5)], margin: [0.04, 0.08, 0.04, 0.08] });
  NOTE(s, "엑셀 건물 시트의 노란 칸이 이 표의 입력과 같은 순서입니다(건물 B6–B13, 구간 행 20–23).", 6.45);
  addNotes(s, ["[스크립트]", "- 현장 조사표는 평가식의 입력을 채우는 양식입니다. 각 항목이 어떤 변수가 되는지 오른쪽 열에 적었습니다."]);
}

// ================================================================================================ 14 stage 3 (precise) + s chart
{
  const s = content("단계별 절차", "3단계 정밀 평가 — 시스템 시험과 시뮬레이션으로 s와 t₆₀₀ 확정",
    `경계에 걸린 건물은 확산시간 비 s가 판정을 좌우 — 예제 B는 s ≥ ${f2(D.scale.filter((x) => x.building === "B" && x.grade === 1).map((x) => x.s)[0] || 1)}이면 1등급,\n중공층 · 실대형(KS F 8414) 시험과 FDS로 s를 정하는 것이 3단계`);
  const lw = 5.5;
  card(s, MX, 1.85, lw, 2.2, "good", "시스템 시험", [
    "중규모 3 m 중공층(30 kW 20분): t₆₀₀ · 자립 연소(B)",
    "실대형 KS F 8414: 층 단위 확산 → s 보정",
    "배리어 · 차단재를 넣은 시험 → p_cb · p_e",
  ], { size: 11.5 });
  card(s, MX, 4.2, lw, 2.3, "conditional", "시뮬레이션(FDS)", [
    "중공층 시험 재현으로 검증(t₆₀₀ ±30 %, 최고온도 ±100 K)",
    "건물 외벽 3–6개 층: 층별 도달시각 → s",
    "용융 · 적하 · 박리는 재현 한계 → 시험 우선",
  ], { size: 11.5 });
  const sg = D.scale.filter((x) => x.building === "B");
  const xs = sg.map((x) => x.s);
  const ch = { x: MX + lw + 0.3, y: 1.8, w: CW - lw - 0.3, h: 4.6, L: { x: 0.13, y: 0.05, w: 0.78, h: 0.8 }, xMin: 0.25, xMax: 1.0, xMajor: 0.25, yMin: 1e-3, yMax: 1, yLog: true };
  const mp = plotMap(ch);
  gradeBands(s, mp, ch.yMin, ch.yMax);
  curves(s, ch, xs, [{ name: "B r vs s", ys: sg.map((x) => x.r), color: BC.B, style: { marker: { symbol: "circle", size: 7, fill: BC.B, line: BC.B } } }],
    { lineSize: 2.25, valAxisLogScaleBase: 10, valAxisLabelFormatCode: "0.000", catAxisLabelFormatCode: "0.00", valAxisMajorUnit: undefined });
  text(s, "실물/시험 층당 확산시간 비 s", { x: mp.px, y: mp.py + mp.ph + 0.3, w: mp.pw, h: 0.26, fontSize: 11, color: C.text2, align: "center" });
  vlabel(s, "예제 B의 r (로그)", ch.x + 0.12, mp.py + mp.ph / 2, 2.2);
  addNotes(s, [
    "[스크립트]",
    "- 3단계는 판정이 경계에 걸린 건물만 합니다. 예제 B는 s가 0.84 이상이면 1등급으로 내려갑니다. 즉 이 건물의 결론은 재료 시험을 더 하는 것보다 s를 정하는 데 달려 있습니다.",
    "- s는 실대형 시험과 FDS 해석으로 정합니다.",
  ]);
}

// ================================================================================================ 15 judgement
{
  const s = content("단계별 절차", "4단계 판정 · 대책 — 등급, 표시, 대책별 위험 감소, 문서",
    "등급(1인당 위험)과 우선순위(줄일 수 있는 노출 인원 ΔR)를 나눠 보고,\n대책마다 같은 식으로 다시 계산해 등급이 몇 단계 내려가는지 보임");
  const cw = (CW - 2 * 0.25) / 3;
  card(s, MX, 1.85, cw, 2.6, "good", "등급 · 표시", [
    "1–5등급(r = IR / λ_res)",
    "M: 노출 1건당 10명 이상",
    "A: 급발(용융 · 가연 전면 화염)",
    "표기 예: 4등급(매우 높음)·M·A, 증거 E2",
  ], { size: 12 });
  card(s, MX + cw + 0.25, 1.85, cw, 2.6, "conditional", "우선순위", [
    "ΔR = R − R_ref: 외벽을 불연으로 바꾸면 줄어드는 노출 인원",
    "같은 등급이면 ΔR가 큰 건물부터",
    "대책 비용이 있으면 ΔR ÷ 비용",
  ], { size: 12 });
  card(s, MX + 2 * (cw + 0.25), 1.85, cw, 2.6, "fair", "대책(같은 식으로 재계산)", [
    "개구부 차단재: p_e 0.9 → 0.3",
    "화재확산방지구조: 층 경계마다 p_cb 0.7",
    "하단 보호 · 필로티 개선: 천장 불연, 1층 방화구획",
    "스프링클러 전층 · 외장재 불연 교체",
  ], { size: 12 });
  const rows = [
    ["NFPA 551 7.3.3.4 정량 결과 문서", "평가식 · 변수 · 출처(이 덱 5–9쪽, 엑셀 4_공통변수)"],
    ["모델 · 자료 평가와 한계", "증거 수준(E0–E3), 예시값 표시, 재현 한계(용융 · 박리)"],
    ["불확실성", "몬테카를로 등급 확률, 민감도(무엇이 등급을 바꾸나)"],
    ["검토(8장)", "엑셀 수식 = 파이썬 계산(12_검증), 대책 스위치 재현"],
  ];
  table(s, [[TH("문서 항목"), TH("이 방법에서 채우는 내용")], ...rows.map(([a, b]) => [TD(a, { bold: true }), TD(b)])],
    { x: MX, y: 4.65, w: CW, colW: [4.0, CW - 4.0], fontSize: 11.5, border: BORDER, rowH: [0.34, 0.4, 0.4, 0.4, 0.4], margin: [0.03, 0.08, 0.03, 0.08] });
  addNotes(s, ["[스크립트]", "- 등급은 1인당 위험이라 작은 건물도 높을 수 있고, 우선순위는 줄일 수 있는 노출 인원(ΔR)으로 봅니다. 대책은 같은 식으로 다시 계산해 효과를 숫자로 보입니다."]);
}

// ================================================================================================ 16 example A
pres.addSection({ title: "예제" });
const SHORT_OPT = { "현재": "현재", "개구부 차단재": "차단재", "화재확산방지구조": "확산\n방지구조", "하단 보호·필로티 개선": "하단 보호\n·필로티",
  "차단재 + 확산방지구조 + 하단 보호": "셋 함께", "스프링클러 전층 설치": "스프링\n클러", "외장재 불연 교체": "외장재\n교체" };
function remChart(s, bid, x, y, w, h) {
  const opts = D.remediation.filter((r) => r.building === bid);
  const ch = { x, y, w, h, L: { x: 0.13, y: 0.04, w: 0.74, h: 0.76 }, xMin: -0.5, xMax: opts.length - 0.5, xMajor: 1, yMin: 1e-4, yMax: 1e2, yLog: true };
  const mp = plotMap(ch);
  gradeBands(s, mp, ch.yMin, ch.yMax);
  const xs = opts.map((o, i) => i);
  curves(s, ch, xs, [{ name: `rem ${bid}`, ys: opts.map((o) => o.r), color: BC[bid], style: { noLine: true, marker: { symbol: "circle", size: 9, fill: BC[bid], line: BC[bid] } } }],
    { valAxisLogScaleBase: 10, valAxisLabelFormatCode: "General", catAxisHidden: true, valAxisMajorUnit: undefined, catAxisCrossesAt: ch.xMin });
  const cw = mp.pw / opts.length;
  opts.forEach((o, i) => {
    text(s, SHORT_OPT[o.option] || o.option, { x: mp.X(i) - cw / 2, y: mp.py + mp.ph + 0.06, w: cw, h: 0.42, fontSize: 9.5, align: "center", color: HEX.ink });
    text(s, `${o.grade}`, { x: mp.X(i) - 0.2, y: mp.Y(o.r) - 0.36, w: 0.4, h: 0.24, fontSize: 9.5, bold: true, align: "center" });
  });
  vlabel(s, "r (로그)", x + 0.3, mp.py + mp.ph / 2, 1.6);
}
{
  const sh = share("A");
  const pil = sh.z.find((z) => z.zone === "A1").r2 / sh.R, oth = sh.z.find((z) => z.zone === "A2").r2 / sh.R;
  const s = content("예제", `예제 A — ${BLD("A").form}`,
    `${A_.label} — 위험의 ${pct(pil)}가 필로티 차량 화재에서 옴:\n차단재만으로는 ${OPT("A", "개구부 차단재").grade}등급, 셋을 함께 하면 ${OPT("A", "차단재 + 확산방지구조 + 하단 보호").grade}등급, 외장재 교체면 ${OPT("A", "외장재 불연 교체").grade}등급`);
  const lw = 4.6;
  const b = BLD("A");
  table(s, [[TH("구간"), TH("층"), TH("외벽"), TH("하단 화원")],
    ...b.zones.map((z) => [TD(`${z.id} ${z.name}`, { bold: true, fontSize: 10.5 }), TD(`${z.b}–${z.t}`, { align: "center" }), TD("EPS 외단열(가연)", { fontSize: 10.5 }),
      TD(z.ext.map((e) => (e.src === "piloti" ? `필로티 ${e.count}면` : `기타 ${e.count}`)).join(", "), { fontSize: 10.5 })])],
  { x: MX, y: 1.85, w: lw, colW: [1.55, 0.55, 1.45, 1.05], fontSize: 10.5, border: BORDER, rowH: [0.34, 0.44, 0.44], margin: [0.03, 0.05, 0.03, 0.05] });
  text(s, para([
    `스프링클러 없음 · 필로티 천장 가연 · 1층 방화구획 없음 · 허가 ${b.permit}(6층 미만 → 당시 기준 없음)`,
    `R ${sci(A_.R)}명/년 · 재실자 ${fx(A_.Nocc)}명 → IR ${sci(A_.IR)} · r ${f2(A_.r)}`,
    `노출 1건당 ${f2(A_.Cbar, 1)}명(M) · 외벽 기여 ${pct(A_.AF, 1)}`,
  ], { space: 4 }), { x: MX, y: 3.15, w: lw, h: 1.3, fontSize: 11 });
  s.addChart(pres.ChartType.bar, [
    { name: "몫", labels: ["필로티 차량(IE2)", "기타 하단 화원(IE2)", "실내 화재 분출(IE1)"], values: [pil * 100, oth * 100, sh.ie1 * 100] },
  ], { x: MX, y: 4.5, w: lw, h: 2.2, barDir: "bar", chartColors: ["4B5563"], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0\"%\"",
    dataLabelFontSize: 10.5, dataLabelColor: HEX.ink, catAxisLabelColor: HEX.ink, valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    catAxisLabelFontSize: 10.5, showLegend: false, showTitle: true, title: "위험의 구성(%)", titleFontSize: 11.5, titleColor: HEX.ink, catAxisOrientation: "maxMin",
    valAxisMaxVal: 80, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", titleFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", barGapWidthPct: 60 });
  remChart(s, "A", MX + lw + 0.25, 1.8, CW - lw - 0.25, 5.0);
  addNotes(s, [
    "[스크립트]",
    "- 5층 필로티 다세대, 드라이비트 외벽, 스프링클러 없음. 위험의 절반 이상이 필로티 차량 화재에서 와서, 창 둘레 차단재만으로는 등급이 그대로입니다.",
    "- 층간 확산방지구조, 필로티 하단 보호, 차단재를 함께 하면 2등급, 외장재 교체는 1등급입니다.",
  ]);
}

// ================================================================================================ 17 example B (evidence)
{
  const e = (l) => D.evidence.find((x) => x.level === l);
  const s = content("예제", `예제 B — ${BLD("B").form}`,
    `${B_.label}: 사진 · 현장 r ${f2(e("E0").r, 3)} → 시료 TGA ${f2(e("E2").r, 3)} → 중공층 시험 ${f2(e("E3").r, 3)} —\n시험할수록 3등급 이상일 확률이 ${pct(pGe3(evP("E0")))} → ${pct(pGe3(evP("E3")))}로 줄어 2등급이 확정됨`);
  const lw = 6.2;
  const rows = ["E0", "E1", "E2", "E3"].map((l) => {
    const x = e(l);
    return [TD(l, { bold: true, align: "center" }), TD(fx(x.t600), { align: "center" }), TD(f2(x.pB, 2), { align: "center" }), TD(f2(x.r, 3), { align: "center" }),
      gradeCell(`${x.grade} ${GK[x.grade]}`, x.grade), TD(x.note, { fontSize: 9.5 })];
  });
  table(s, [[TH("수준"), TH("t₆₀₀ (s)"), TH("p_B"), TH("r"), TH("등급"), TH("재료 인자 근거")], ...rows],
    { x: MX, y: 1.85, w: lw, colW: [0.55, 0.75, 0.55, 0.65, 0.95, 2.75], fontSize: 10.5, border: BORDER, rowH: [0.34, 0.62, 0.62, 0.62, 0.62], margin: [0.03, 0.05, 0.03, 0.05] });
  text(s, para([
    `대책: 개구부 차단재 ${OPT("B", "개구부 차단재").grade}등급 · 화재확산방지구조 ${OPT("B", "화재확산방지구조").grade}등급 · 셋 함께 ${OPT("B", "차단재 + 확산방지구조 + 하단 보호").grade}등급`,
    "스프링클러가 있어 실내 화재가 창으로 분출하는 빈도가 낮음 → 1인당 위험은 2등급",
  ], { space: 4 }), { x: MX, y: 4.8, w: lw, h: 1.0, fontSize: 11 });
  const levels = ["E0", "E1", "E2", "E3"];
  const data = [1, 2, 3, 4, 5].map((g) => ({ name: `${g}등급`, labels: levels, values: levels.map((l) => evP(l)[`P_g${g}`] * 100) }));
  s.addChart(pres.ChartType.bar, data, {
    x: MX + lw + 0.3, y: 1.8, w: CW - lw - 0.3, h: 4.9, barDir: "col", barGrouping: "stacked", chartColors: [1, 2, 3, 4, 5].map((g) => GC[g]),
    showValue: true, dataLabelPosition: "ctr", dataLabelFormatCode: '[<6]"";0', dataLabelFontSize: 10.5, dataLabelColor: HEX.ink,
    valAxisMinVal: 0, valAxisMaxVal: 100, valAxisMajorUnit: 20, valAxisLabelFormatCode: "0", catAxisLabelColor: HEX.ink, valAxisLabelColor: HEX.ink2,
    catAxisLabelFontSize: 11, valAxisLabelFontSize: 10, valGridLine: { color: HEX.grid, size: 0.75 }, catGridLine: { style: "none" }, showLegend: true, legendPos: "b",
    legendFontSize: 10.5, legendColor: HEX.ink2, showTitle: true, title: "증거 수준별 등급 확률(%, 몬테카를로)", titleFontSize: 12, titleColor: HEX.ink,
    catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", titleFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", barGapWidthPct: 55,
  });
  addNotes(s, [
    "[스크립트]",
    "- 같은 건물을 증거 수준별로 평가한 결과입니다. 사진만 있을 때는 준불연 미상 보수값을 쓰고, 시료 TGA와 중공층 시험으로 갈수록 값이 측정값으로 바뀝니다.",
    "- 오른쪽은 등급의 확률입니다. 시험할수록 3등급 이상일 확률이 25 %에서 12 %로 줄어 2등급이 확정됩니다. 이것이 재료 시험의 가치를 숫자로 보인 것입니다.",
  ]);
}

// ================================================================================================ 18 example C
{
  const sh = share("C");
  const c2 = sh.z.find((z) => z.zone === "C2"), c3 = sh.z.find((z) => z.zone === "C3");
  const s = content("예제", `예제 C — ${BLD("C").form}`,
    `${C_.label}: 위험의 ${pct(c2.r3 / sh.R)}는 저층 상가의 분출화염이 고층 패널을 가열하는 경로(IE3) —\nPE 심재 패널은 확산방지구조 · 차단재 효과를 인정하지 않아 교체만 등급을 낮춤(${OPT("C", "외장재 불연 교체").grade}등급)`);
  const lw = 4.6;
  const b = BLD("C");
  table(s, [[TH("구간"), TH("층"), TH("외벽"), TH("비고")],
    ...b.zones.map((z) => [TD(`${z.id} ${z.name.replace("(주거)", "")}`, { bold: true, fontSize: 10 }), TD(`${z.b}–${z.t}`, { align: "center" }),
      TD(z.type === "NC" ? "석재 · 커튼월(불연)" : "PE 심재 복합패널", { fontSize: 10 }),
      TD(z.use === "res" ? (z.win < 1 ? "주거 · 창 적음" : "주거") : z.gap ? "층간 틈 미흡" : "", { fontSize: 10 })])],
  { x: MX, y: 1.85, w: lw, colW: [1.5, 0.5, 1.5, 1.1], fontSize: 10, border: BORDER, rowH: [0.34, 0.42, 0.42, 0.42], margin: [0.03, 0.05, 0.03, 0.05] });
  const reachK = Math.floor(P("h_reach") / b.h) + 1;
  text(s, para([
    `스프링클러 전층 · 고가사다리 ${P("h_reach")} m → ${reachK + 1}층부터 외부 진압 늦음`,
    `R ${sci(C_.R)}명/년 · 재실자 ${fx(C_.Nocc)}명 → r ${f2(C_.r, 3)} · 노출 1건당 ${f2(C_.Cbar, 0)}명(M)`,
    `ΔR ${sci(C_.dR)}명/년 — 세 건물 중 가장 큼(우선순위 1)`,
  ], { space: 4 }), { x: MX, y: 3.6, w: lw, h: 1.2, fontSize: 11 });
  s.addChart(pres.ChartType.bar, [
    { name: "몫", labels: ["C2 세대 화재 분출(IE1)", "C2 저층 상가 분출(IE3)", "C3 측면 분출(IE1)"], values: [c2.r1 / sh.R * 100, c2.r3 / sh.R * 100, c3.r1 / sh.R * 100] },
  ], { x: MX, y: 4.85, w: lw, h: 1.95, barDir: "bar", chartColors: ["4B5563"], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0\"%\"",
    dataLabelFontSize: 10.5, dataLabelColor: HEX.ink, catAxisLabelColor: HEX.ink, valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    catAxisLabelFontSize: 10, showLegend: false, showTitle: true, title: "위험의 구성(%)", titleFontSize: 11.5, titleColor: HEX.ink, catAxisOrientation: "maxMin",
    valAxisMaxVal: 80, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", titleFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", barGapWidthPct: 60 });
  remChart(s, "C", MX + lw + 0.25, 1.8, CW - lw - 0.25, 5.0);
  addNotes(s, [
    "[스크립트]",
    "- 25층 주상복합, 고층부 PE 심재 복합패널. 위험의 3분의 1은 저층 상가 화재의 분출화염이 고층 패널 하단을 가열하는 경로입니다. 구간으로 나눠야 이 경로가 보입니다.",
    "- 가연재가 바깥 면인 패널은 층간 띠나 차단재로 끊기지 않으므로, 교체 말고는 등급이 내려가지 않습니다.",
  ]);
}

// ================================================================================================ 19 comparison
{
  const s = content("예제", "세 건물 비교 — 등급(1인당 위험)과 우선순위(줄일 수 있는 노출 인원)",
    `등급은 필로티 다세대가 가장 높고(${A_.grade}등급) 우선순위는 주상복합이 가장 높음(ΔR 최대) —\n작은 건물의 높은 개인 위험과 큰 건물의 많은 노출 인원을 함께 보기 위해 두 지표를 나눔`);
  const bl = ["A", "B", "C"];
  const rank = [...bl].sort((x, y) => SUM(y).dR - SUM(x).dR);
  const rows = bl.map((b) => {
    const x = SUM(b);
    return [TD(`${b} ${x.name}`, { bold: true }), TD(BLD(b).form, { fontSize: 10 }), gradeCell(`${x.grade} ${GK[x.grade]}${x.M ? "·M" : ""}${x.A ? "·A" : ""}`, x.grade),
      TD(f2(x.r, 3), { align: "center" }), TD(sci(x.IR), { align: "center" }), TD(f2(x.Cbar, 1), { align: "center" }), TD(pct(x.AF, 1), { align: "center" }),
      TD(sci(x.dR), { align: "center" }), TD(String(rank.indexOf(b) + 1), { align: "center", bold: true })];
  });
  table(s, [[TH("건물"), TH("형태"), TH("등급"), TH("r"), TH("IR (/인·년)"), TH("C̄ (명/건)"), TH("외벽 기여"), TH("ΔR (명/년)"), TH("우선순위")], ...rows],
    { x: MX, y: 1.85, w: CW, colW: [1.9, 3.6, 1.5, 0.8, 1.15, 0.95, 0.95, 1.05, 0.43], fontSize: 11, border: BORDER, rowH: [0.36, 0.56, 0.56, 0.56], margin: [0.03, 0.06, 0.03, 0.06] });
  const ch = { x: MX, y: 4.2, w: 7.0, h: 2.55, L: { x: 0.12, y: 0.06, w: 0.76, h: 0.74 }, xMin: -0.5, xMax: 2.5, xMajor: 1, yMin: 1e-3, yMax: 1e1, yLog: true };
  const mp = plotMap(ch);
  gradeBands(s, mp, ch.yMin, ch.yMax);
  const xs = [], ser = [];
  bl.forEach((b, i) => {
    xs.push(i);
  });
  bl.forEach((b, i) => {
    const ys = [null, null, null];
    ys[i] = SUM(b).r;
    ser.push({ name: `cmp ${b}`, ys, color: BC[b], style: { noLine: true, marker: { symbol: "circle", size: 11, fill: BC[b], line: BC[b] } } });
  });
  curves(s, ch, xs, ser, { valAxisLogScaleBase: 10, valAxisLabelFormatCode: "General", catAxisHidden: true, valAxisMajorUnit: undefined, catAxisCrossesAt: ch.xMin });
  bl.forEach((b, i) => text(s, `${b} ${SUM(b).name}`, { x: mp.X(i) - 0.8, y: mp.py + mp.ph + 0.05, w: 1.6, h: 0.26, fontSize: 10, align: "center" }));
  vlabel(s, "r (로그)", ch.x + 0.3, mp.py + mp.ph / 2, 1.6);
  text(s, para([
    "등급: 1인당 위험(작은 필로티 건물도 높게 나옴)",
    "우선순위: ΔR(외벽 교체 · 보강으로 줄일 수 있는 노출 인원)",
    "M · A 표시는 등급과 별도로 보고(대형 피해 · 급발)",
  ], { space: 6 }), { x: MX + 7.3, y: 4.35, w: CW - 7.3, h: 2.3, fontSize: 11.5 });
  addNotes(s, ["[스크립트]", "- 등급과 우선순위를 나눈 이유입니다. 필로티 다세대는 1인당 위험이 가장 높고, 주상복합은 줄일 수 있는 노출 인원이 가장 많습니다."]);
}

// ================================================================================================ 20 uncertainty
pres.addSection({ title: "불확실성 · 연구 과제" });
{
  const s = content("불확실성 · 연구 과제", `불확실성과 민감도 — 무엇이 등급을 바꾸나 (몬테카를로 ${D.n_mc}회)`,
    "등급을 좌우하는 자료가 형태마다 다름 — 필로티는 외부 화원 빈도, 중층 준불연은 s,\n고층 가연은 스프링클러 · 최성기 확률을 먼저 모아야 함");
  const bl = ["A", "B", "C"];
  const data = [1, 2, 3, 4, 5].map((g) => ({ name: `${g}등급`, labels: bl.map((b) => `${b} ${SUM(b).name}`), values: bl.map((b) => UNC(b)[`P_g${g}`] * 100) }));
  s.addChart(pres.ChartType.bar, data, {
    x: MX, y: 1.8, w: 6.2, h: 4.6, barDir: "bar", barGrouping: "stacked", chartColors: [1, 2, 3, 4, 5].map((g) => GC[g]),
    showValue: true, dataLabelPosition: "ctr", dataLabelFormatCode: '[<6]"";0', dataLabelFontSize: 10.5, dataLabelColor: HEX.ink,
    valAxisMinVal: 0, valAxisMaxVal: 100, valAxisMajorUnit: 20, valAxisLabelFormatCode: "0", catAxisLabelColor: HEX.ink, valAxisLabelColor: HEX.ink2,
    catAxisLabelFontSize: 11, valAxisLabelFontSize: 10, valGridLine: { color: HEX.grid, size: 0.75 }, catGridLine: { style: "none" }, showLegend: true, legendPos: "b",
    legendFontSize: 10.5, legendColor: HEX.ink2, showTitle: true, title: "등급 확률(%)", titleFontSize: 12, titleColor: HEX.ink, catAxisOrientation: "maxMin",
    catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", titleFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", barGapWidthPct: 45,
  });
  const rx = MX + 6.45, rw = CW - 6.45;
  const rows = [];
  bl.forEach((b) => {
    D.sensitivity.filter((x) => x.building === b).slice(0, 3).forEach((x, i) => {
      rows.push([TD(i === 0 ? `${b} ${SUM(b).name}` : "", { bold: true, fontSize: 10.5 }), TD(x.symbol, { align: "center", fontSize: 10.5 }), TD(x.item, { fontSize: 10 }),
        TD(f2(x.swing, 2), { align: "center", fontSize: 10.5 })]);
    });
  });
  table(s, [[TH("건물", { fontSize: 10.5 }), TH("변수", { fontSize: 10.5 }), TH("내용", { fontSize: 10.5 }), TH("폭(log)", { fontSize: 10.5 })], ...rows],
    { x: rx, y: 1.85, w: rw, colW: [1.45, 0.95, rw - 3.2, 0.8], fontSize: 10.5, border: BORDER, rowH: [0.34, ...rows.map(() => 0.42)], margin: [0.03, 0.05, 0.03, 0.05] });
  NOTE(s, "폭(log): 변수를 하한 → 상한으로 바꿀 때 r이 변하는 자릿수(log₁₀). 엑셀 11_불확실성에 전체 표.", 6.5, { x: rx, w: rw, h: 0.4 });
  addNotes(s, [
    "[스크립트]",
    "- 왼쪽은 입력의 불확실성을 모두 반영한 등급 확률입니다. 세 건물 모두 한 등급에 70 % 이상이 몰립니다.",
    "- 오른쪽이 연구 과제의 우선순위입니다. 필로티 건물은 차량 화재 통계, 중층 준불연 아파트는 s, 고층 가연 건물은 스프링클러 효과와 최성기 도달 확률이 등급을 좌우합니다.",
  ]);
}

// ================================================================================================ 21 research needs
{
  const s = content("불확실성 · 연구 과제", "연구로 채울 자료와 검증 계획",
    "절차 · 식은 지금 쓸 수 있고, 등급의 정확도는 국내 자료로 높임 —\n외벽 시스템 시험 DB와 외부 화원 통계가 1순위, 과거 화재 사례로 등급을 검증");
  const rows = [
    ["외벽 시스템 시험 DB", "드라이비트 · 복합패널 · 샌드위치패널의 중규모 시험(t₆₀₀ · B)", "재료 인자(E0 경계값 대체)", "1"],
    ["외부 화원 통계", "필로티 주차 차량 화재 · 건물 외부 발화 빈도(국가화재정보시스템)", "λ_car · λ_adj (필로티 등급)", "1"],
    ["실물 보정 s", "KS F 8414 실대형 + FDS 건물 외벽 해석", "τ = s · t₆₀₀ (중층 준불연 등급)", "1"],
    ["대책 효과", "배리어 · 차단재를 넣은 시험", "p_cb · p_e", "2"],
    ["실내 화재 분기확률", "스프링클러 작동 · 제어, 최성기 · 창 분출 비율", "p_s · p_FO · p_v", "2"],
    ["피난 · 재실", "용도 · 경보 방식별 RSET, 야간 재실", "q(t), c(k)", "2"],
    ["기준 이력 확정", "외벽 마감재료 · 스프링클러 기준의 시행일(법령 원문)", "0단계 E0 규칙", "1"],
    ["사례 검증", "과거 외벽 화재(의정부 2015 등)의 확산 층수와 등급 비교, 시범 평가 10–20동", "등급 경계 보정", "1"],
  ];
  table(s, [[TH("자료"), TH("방법"), TH("쓰는 곳"), TH("순위")],
    ...rows.map(([a, b, c, d]) => [TD(a, { bold: true }), TD(b), TD(c, { fontSize: 10.5 }), TD(d, { align: "center", bold: true, fill: { color: d === "1" ? "E2F3E2" : "F3F4F6" } })])],
  { x: MX, y: 1.85, w: CW, colW: [2.4, 5.6, 3.5, 0.83], fontSize: 11, border: BORDER, rowH: [0.36, ...rows.map(() => 0.5)], margin: [0.04, 0.08, 0.04, 0.08] });
  addNotes(s, ["[스크립트]", "- 순위는 등급을 가장 크게 바꾸는 자료부터 정했습니다. 외벽 시스템 시험 DB, 외부 화원 통계, 실물 보정 s, 그리고 과거 사례로 등급 경계를 검증하는 것이 1순위입니다."]);
}

// ================================================================================================ 22 limitations
{
  const s = content("불확실성 · 연구 과제", "한계와 해석 주의",
    "평가식과 절차는 형태와 관계없이 쓸 수 있지만, 수치는 예시값과 일부 경계값 위에 있음 —\n등급은 '현재 자료로 본 위험'이고 자료가 늘면 같은 식으로 갱신");
  const cw = (CW - 0.25) / 2;
  card(s, MX, 1.85, cw, 4.6, "poor", "한계", [
    "빈도 · 분기확률 · 시간 · s는 예시값(근거 표시), 실내 화재 빈도만 통계",
    "시험 자료 없는 외벽 시스템(드라이비트 · 복합패널 · 샌드위치패널)은 가연 경계값(30 s)",
    "수직 확산만 다룸: 수평 확산 · 모서리 · 낙하물 · 연기의 실내 유입은 단순화",
    "구간 경계에서 외벽 경로는 끊긴다고 봄(아래 구간 분출은 IE3로 반영)",
    "등급 경계(자택 화재 위험의 1 % · 10 % · 1배 · 10배)는 제안값",
  ], { size: 12, space: 6 });
  card(s, MX + cw + 0.25, 1.85, cw, 4.6, "good", "해석 주의", [
    "등급은 증거 수준과 함께 표기(예: 3등급 · E0) — E0 등급은 보수적으로 높게 나옴",
    "가연 외장재 건물에서 차단재만으로 등급이 안 내려가는 것은 하단 화원 · 패널 심재 때문 — 구성 몫을 함께 보고",
    "등급(1인당)과 우선순위(ΔR)는 다른 질문에 대한 답",
    "엑셀 수식과 파이썬 계산은 같음(상대차 10⁻¹¹ 수준) — 값의 근거는 4_공통변수의 근거 열",
  ], { size: 12, space: 6 });
  addNotes(s, ["[스크립트]", "- 방법은 지금 쓸 수 있고, 수치의 정확도는 자료에 달려 있습니다. 특히 시험 자료가 없는 외벽 시스템은 보수적으로 평가됩니다."]);
}

// ================================================================================================ 23 appendix
pres.addSection({ title: "부록" });
{
  const s = content("부록", "부록 — 주요 변수의 기본값과 재현 방법", "모든 수치는 ebfra.py로 재현되고, 엑셀 계산서가 같은 값을 셀 수식으로 계산함");
  const keys = ["s", "lam_res", "lam_car", "lam_adj", "p_s", "p_fo", "p_v", "p_e0", "p_e1", "p_cb", "p_L", "tau_L", "t_fs_lo", "t_fs_hi", "rset_res", "p_stair"];
  const rows = keys.map((k) => {
    const p = PAR(k);
    const v = (x) => (x === null || x === undefined ? "—" : Math.abs(x) < 0.01 && x !== 0 ? sci(x) : String(x));
    return [TD(p.symbol, { align: "center", fontSize: 10 }), TD(p.item, { fontSize: 9.5 }), TD(v(p.best), { align: "center", fontSize: 10 }),
      TD(`${v(p.low)} – ${v(p.high)}`, { align: "center", fontSize: 9.5 })];
  });
  table(s, [[TH("기호", { fontSize: 10 }), TH("변수", { fontSize: 10 }), TH("값", { fontSize: 10 }), TH("범위", { fontSize: 10 })], ...rows],
    { x: MX, y: 1.8, w: 7.6, colW: [0.95, 4.35, 1.0, 1.3], fontSize: 10, border: BORDER, rowH: [0.3, ...keys.map(() => 0.29)], margin: [0.02, 0.05, 0.02, 0.05] });
  const rx = MX + 7.85, rw = CW - 7.85;
  box(s, rx, 1.8, rw, 4.95, { fill: "F3F4F6", line: "F3F4F6", r: 0.08, name: "reproduce" });
  text(s, para([
    [["재현 (existing_facade_fra/)", { bold: true }]],
    [["python ebfra.py", { fontFace: "Consolas" }], [` → results/, figures/, report/ebfra_data.json (몬테카를로 ${D.n_mc}회)`, {}]],
    [["python report/build_ebfra_xlsx.py", { fontFace: "Consolas" }], [" → 계산서(수식) → recalc", {}]],
    [["node report/build_ebfra_deck.js", { fontFace: "Consolas" }], [" → 이 덱", {}]],
    [["입력", { bold: true }]],
    [["inputs/parameters.csv(변수 · 근거), facade_types.csv(외벽 유형), code_history.csv(허가 시점 기준), buildings.json(예제 건물 · 구간)", {}]],
    [["자료 출처", { bold: true }]],
    [["중공층 시험 · TGA(cavity_fire_index, tga_dsc, integrated_model, multiscale_fra), 건물 모델(facade_fra), 소방청 2023 화재 통계, 인구주택총조사, BSI PAS 9980, Lamont & Ingolfsson(2018), NFPA 551(2022)", {}]],
  ], { bullet: false, space: 6 }), { x: rx + 0.2, y: 1.9, w: rw - 0.4, h: 4.75, fontSize: 10.5 });
  addNotes(s, ["[스크립트]", "- 변수의 값과 범위, 재현 방법입니다. 엑셀 4_공통변수에 모든 변수의 근거가 있습니다."]);
}

// pptxgenjs writes the X axis of a scatter chart (its first <c:valAx>) with the Y axis number format: set it for the
// charts whose X axis is shown
const JSZip = require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));
const X_AXES = [{ series: "reach NC", fmt: "0" }, { series: "B r vs s", fmt: "0.00" }];
async function fixX(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  for (const name of Object.keys(zip.files).filter((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n))) {
    let x = await zip.file(name).async("string");
    const ax = X_AXES.find((a) => x.includes(`<c:v>${a.series}</c:v>`));
    if (!ax) continue;
    let first = true;
    x = x.replace(/<c:valAx>[\s\S]*?<\/c:valAx>/g, (blk) => {
      if (!first) return blk;
      first = false;
      return blk.replace(/<c:numFmt formatCode="[^"]*" sourceLinked="[01]"\/>/, `<c:numFmt formatCode="${ax.fmt}" sourceLinked="0"/>`);
    });
    zip.file(name, x);
  }
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await finish(OUT);
  await fixX(OUT);
  console.log("wrote", OUT);
})();
