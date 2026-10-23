#!/usr/bin/env node
// 외벽 정량적 화재위험성평가(QFRA) 절차 덱 (pptxgenjs, native charts): NFPA 551 5.5를 외벽에 적용하는 12단계와
// 중공층 시험으로 계산한 예제. Reads qfra_data.json (../qfra.py) and verification.json (check_xlsx.py).
//
//   python ../qfra.py && python build_qfra_xlsx.py && python <xlsx skill>/scripts/recalc.py 외벽_QFRA_계산서.xlsx
//   python check_xlsx.py && NODE_PATH=<folder holding pptxgenjs> node build_qfra_deck.js [out.pptx]
"use strict";
const fs = require("fs");
const path = require("path");
const { makeDeck } = require("../../tga_dsc/report/deck_lib");

const D = JSON.parse(fs.readFileSync(path.join(__dirname, "qfra_data.json"), "utf8"));
const VER = JSON.parse(fs.readFileSync(path.join(__dirname, "verification.json"), "utf8"));
const OUT = process.argv[2] || path.join(__dirname, "외벽_정량적_화재위험성평가_절차.pptx");

const SUBS = {
  R_S: ["R", "S"], R_P: ["R", "P"], R_t: ["R", "t"], RR_S: ["RR", "S"], F_i: ["F", "i"], C_i: ["C", "i"], C_ij: ["C", "ij"],
  "F₁": ["F", "1"], "F₂": ["F", "2"], "λ₁": ["λ", "1"], "λ₂": ["λ", "2"], p_s: ["p", "s"], p_FO: ["p", "FO"], p_v: ["p", "v"],
  p_x: ["p", "x"], p_e0: ["p", "e0"], p_e1: ["p", "e1"], p_e: ["p", "e"], p_cb: ["p", "cb"], p_L: ["p", "L"], p_B: ["p", "B"],
  p_d: ["p", "d"], T_FS: ["T", "FS"], T_stop: ["T", "stop"], "τ_F": ["τ", "F"], "τ_L": ["τ", "L"], n_F: ["n", "F"], n_L: ["n", "L"],
  w_j: ["w", "j"], a_j: ["a", "j"], b_j: ["b", "j"], t_e: ["t", "e"], t_ref: ["t", "ref"], n_a: ["n", "a"], n_u: ["n", "u"],
  t_ig: ["t", "ig"], t_sp: ["t", "sp"], r_max: ["r", "max"], C_L: ["C", "L"], "C₁": ["C", "1"], "C₂": ["C", "2"], t_b: ["t", "b"],
};
const {
  pres, C, HEX, STATUS, MX, CW, LANG, content: content0, text: text0, vlabel, table, paras, badge, legend, chip, plotMap, curves, finish,
} = makeDeck({
  title: "외벽 화재위험성평가의 정량적 방법과 절차",
  subject: "NFPA 551(2022) 5.5 정량적 화재위험성평가를 외벽(중공층 외단열)에 적용하는 절차, EFFECT 요인 연계, 중공층 시험 예제",
  footer: "KICT 중공층 외단열 · 외벽 정량적 화재위험성평가(NFPA 551 5.5) 절차와 예제",
  themeName: "KICT facade QFRA",
  subs: SUBS,
});

// ------------------------------------------------------------------------------------------------ data access
const MATS = D.materials;
const ORDER = ["NC", ...MATS];
const LABEL = D.label;
const MAT = { NC: "8A8F98", EPS_SC: "E87BA4", PIR_C: "6250D6", PIR_SC: "008300", PF_out: "2A78D6", PF_in: "E34948" };
const SHORT = { NC: "불연\n(기준안)", EPS_SC: "EPS\n준불연", PIR_C: "PIR(PUR)\n가연성", PIR_SC: "PIR(PUR)\n준불연", PF_out: "PF\n외부용", PF_in: "PF\n내부용" };
const PRM = Object.fromEntries(D.params.map((p) => [p.key, p]));
const PT = (m, b, g) => D.point.find((r) => r.material === m && r.building === b && r.mitigation === g);
const MC = (m, b, g) => D.mc.find((r) => r.material === m && r.building === b && r.mitigation === g);
const CO = (m, ie, bar) => D.conditional.find((r) => r.material === m && r.ie === ie && r.barrier === bar);
const EX = (m, b, g) => D.exceedance.filter((r) => r.material === m && r.building === b && r.mitigation === g).sort((a, b2) => a.n - b2.n);
const SE = (g, m, k) => D.sensitivity.find((r) => r.mitigation === g && r.material === m && r.parameter === k);
const TESTS = (m) => D.tests.filter((t) => t.material === m && t.used);
const VERDICT = { 충족: "good", "판단 유보": "fair", 불충족: "poor" };
const best = (k) => PRM[k].best;

// ------------------------------------------------------------------------------------------------ helpers
// "1×10⁻⁴" written with Unicode superscripts becomes runs with a real superscript (the glyphs fall back to another
// font); "\n" becomes a line break run, paragraph options stay on the first piece and breakLine on the last
const SUP = { "⁻": "−", "⁰": "0", "¹": "1", "²": "2", "³": "3", "⁴": "4", "⁵": "5", "⁶": "6", "⁷": "7", "⁸": "8", "⁹": "9" };
const SUP_RE = /[⁻⁰¹²³⁴⁵⁶⁷⁸⁹]/;
function supSplit(str, o) {
  const { breakLine, ...rest } = o;
  const out = [], lines = str.split("\n");
  lines.forEach((line, li) => {
    line.split(/([⁻⁰¹²³⁴⁵⁶⁷⁸⁹]+)/).filter((p) => p !== "").forEach((p) => {
      const sup = /^[⁻⁰¹²³⁴⁵⁶⁷⁸⁹]+$/.test(p), opt = { ...rest };
      if (out.length > 0) { delete opt.bullet; delete opt.paraSpaceAfter; }
      if (sup) opt.superscript = true;
      out.push({ text: sup ? [...p].map((c) => SUP[c]).join("") : p, options: opt });
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
// 3.7×10⁻¹³ as a Unicode string (S_ turns it into runs) for titles built from data
const SUPD = ["⁰", "¹", "²", "³", "⁴", "⁵", "⁶", "⁷", "⁸", "⁹"];
function supStr(v, d = 2) {
  let e = Math.floor(Math.log10(Math.abs(v)));
  let m = v / 10 ** e;
  if (Math.abs(Number(m.toFixed(d - 1))) >= 10) { m /= 10; e += 1; }
  return `${m.toFixed(d - 1)}×10${e < 0 ? "⁻" : ""}${String(Math.abs(e)).split("").map((c) => SUPD[+c]).join("")}`;
}
function tint(hex, t) {
  return [0, 2, 4].map((i) => Math.round(255 + (parseInt(hex.slice(i, i + 2), 16) - 255) * (1 - t)).toString(16).padStart(2, "0")).join("").toUpperCase();
}
// 1.26×10⁻⁴ as runs with a real superscript (Unicode superscripts fall back to another font)
function sci(v, d = 2, o = {}) {
  if (v === null || v === undefined || Number.isNaN(v)) return [{ text: "—", options: { ...o } }];
  if (v === 0) return [{ text: "0", options: { ...o } }];
  let e = Math.floor(Math.log10(Math.abs(v)));
  let m = v / 10 ** e;
  if (Math.abs(Number(m.toFixed(d - 1))) >= 10) { m /= 10; e += 1; }
  return [{ text: `${m.toFixed(d - 1)}×10`, options: { ...o } }, { text: String(e).replace("-", "−"), options: { ...o, superscript: true } }];
}
const sciPlain = (v, d = 2) => (v === 0 ? "0" : Number(v).toExponential(d - 1).replace("e", "E"));
const runs = (...parts) => parts.flatMap((p) => (typeof p === "string" ? [{ text: p, options: {} }] : Array.isArray(p) ? p : [p]));
const withOpt = (rs, o) => rs.map((r) => ({ text: r.text, options: { ...r.options, ...o } }));
const pct = (v, d = 0) => `${(100 * v).toFixed(d)} %`;
const f2 = (v, d = 2) => Number(v).toFixed(d);

const BORDER = { type: "solid", pt: 0.5, color: "D1D5DB" };
const TH = (t, o = {}) => ({ text: S_(t), options: { bold: true, color: "FFFFFF", fill: { color: HEX.ink }, fontSize: 11, align: "center", valign: "middle", ...o } });
const TD = (t, o = {}) => ({ text: S_(t), options: { fontSize: 11, color: HEX.ink, valign: "middle", ...o } });
const NOTE = (s, t, y, o = {}) => text(s, t, { x: MX, y, w: CW, h: 0.3, fontSize: 10.5, color: C.text2, ...o });

function arrow(s, x1, y1, x2, y2, o = {}) {
  s.addShape(pres.ShapeType.line, {
    x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.max(Math.abs(x2 - x1), 0.001), h: Math.max(Math.abs(y2 - y1), 0.001),
    flipH: x2 < x1, flipV: y2 < y1, line: { color: o.color || HEX.ink2, width: o.width || 1.25, endArrowType: o.head === false ? undefined : "triangle", dashType: o.dash ? "dash" : "solid" },
  });
}
function box(s, x, y, w, h, o = {}) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: o.r === undefined ? 0.06 : o.r, fill: { color: o.fill || "FFFFFF" },
    line: { color: o.line || "D1D5DB", width: o.lw || 0.75 }, objectName: o.name });
}
function card(s, x, y, w, h, st, head, items, o = {}) {
  box(s, x, y, w, h, { fill: STATUS[st].tint, line: STATUS[st].tint, r: 0.08, name: `card ${head}` });
  s.addShape(pres.ShapeType.ellipse, { x: x + 0.2, y: y + 0.21, w: 0.2, h: 0.2, fill: { color: STATUS[st].color }, line: { color: STATUS[st].color, width: 0.5 } });
  text(s, head, { x: x + 0.5, y: y + 0.1, w: w - 0.65, h: 0.42, fontSize: o.headSize || 15, bold: true, valign: "middle" });
  text(s, para(items, { space: 4 }), { x: x + 0.22, y: y + 0.58, w: w - 0.42, h: h - 0.7, fontSize: o.size || 12 });
}
function verdictChip(s, x, y, w, v) { chip(s, x, y, w, VERDICT[v] || "poor", v); }
// diverging tint for a correlation (brown + / teal -), white at 0
function heat(v) {
  const t = Math.min(1, Math.abs(v) / 0.8), c = v >= 0 ? [0xb3, 0x58, 0x06] : [0x2d, 0x7f, 0x86];
  const mix = c.map((k) => Math.round(255 + (k - 255) * t));
  return mix.map((k) => k.toString(16).padStart(2, "0")).join("").toUpperCase();
}
function addNotes(s, lines) { s.addNotes(lines.join("\n")); }
// paras() items are strings or [text, options] tuples; runs() gives {text, options} objects -> convert
const tup = (it) => (Array.isArray(it) && it.length && it[0] && typeof it[0] === "object" && !Array.isArray(it[0]) ? it.map((r) => [r.text, r.options]) : it);
const para = (items, o) => paras(items.map((it) => tup(S_(it))), o);

// ================================================================================================ 1 title
pres.addSection({ title: "결론" });
{
  const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "결론" });
  s.addText("외벽 화재위험성평가 · NFPA 551 정량적 방법 · 2026-10-23", { placeholder: "kicker", lang: LANG });
  s.addText("외벽(중공층 외단열) 화재위험성평가의\n정량적 방법과 평가 절차", { placeholder: "title", lang: LANG });
  s.addText("NFPA 551(2022) 5.5 정량적 방법을 외벽에 적용하는 12단계 · EFFECT 요인 연계 · 중공층 시험 자료로 계산한 예제",
    { placeholder: "subtitle", lang: LANG });
  const facts = [
    ["R = Σ F·C", "사건수목 × 층간 확산 모델", runs("NFPA 551 식 5.5.1.2 · 5.5.1.3")],
    ["12단계", "준비 4 · 분석 6 · 문서화·검토 2", runs("단계마다 조항 · 산출물 · 서술 항목")],
    ["엑셀 = 파이썬", `점추정 ${VER.n_cases}경우 재현 일치`, runs("최대 상대차 ", sci(VER.max_rel_diff, 2), " (8.2.1.1)")],
  ];
  facts.forEach(([big, lab, sub], i) => {
    const x = 0.8 + i * 3.95;
    s.addShape(pres.ShapeType.roundRect, { x, y: 4.55, w: 3.7, h: 1.65, rectRadius: 0.08, fill: { color: C.background1, transparency: 90 }, line: { color: C.background2, width: 0.75 } });
    text(s, big, { x: x + 0.25, y: 4.7, w: 3.3, h: 0.6, fontSize: 26, bold: true, color: C.background1 });
    text(s, lab, { x: x + 0.25, y: 5.32, w: 3.3, h: 0.32, fontSize: 13, bold: true, color: C.background1 });
    text(s, withOpt(sub, { color: C.background2 }), { x: x + 0.25, y: 5.66, w: 3.3, h: 0.32, fontSize: 11, color: C.background2 });
  });
  text(s, "함께 드리는 파일: 외벽_QFRA_계산서.xlsx (수식 계산서 — 노란 칸을 바꾸면 다시 계산) · 외벽_QFRA_서술양식.docx (NFPA 551 7장 보고서 작성 양식과 예문)",
    { x: 0.8, y: 6.65, w: 11.7, h: 0.3, fontSize: 11, color: C.background2 });
  addNotes(s, [
    "[스크립트]",
    "- 질문: NFPA 551의 화재위험성평가 중 정량적 방법을 외벽에 쓰고 싶은데, 정량 평가를 어떻게 서술하고 어떤 절차로 하는지가 없다.",
    "- 답: NFPA 551은 '평가를 검토하는 지침'이라 방법을 정하지 않습니다. 그래서 외벽에 맞춘 절차(12단계), 모델(사건수목 + 층간 확산), 서술 양식을 만들었고, 중공층 시험 자료로 끝까지 계산해 보였습니다.",
    "- 순서: 결론(2) → NFPA 551·EFFECT에서 무엇을 가져오나(3–5) → 12단계 절차(6–12) → 예제 결과(13–15) → 불확실성·검증(16–17) → 문서화·검토(18–20) → 한계(21) → 부록.",
    "",
    "[질문 대비]",
    "- EFFECT 사용자 가이드 원문은 이번에 받지 못해 공개 자료 기준으로 연계했습니다(5쪽).",
  ]);
}

// ================================================================================================ 2 conclusion
{
  const s = content("결론", "결론 — 무엇을 어떻게 하면 되나",
    "NFPA 551은 정량 평가의 ‘요건’만 주므로, 외벽은 사건수목(빈도) × 층간 확산 모델(결과)로\n구성하고 중공층 시험 t₆₀₀을 층당 확산시간으로 넣어 R = Σ F·C를 계산함");
  const cw = (CW - 0.25) / 2, ch = 2.2;
  card(s, MX, 1.8, cw, ch, "fair", "NFPA 551 = ‘평가를 검토하는 법’을 주는 지침", [
    "관할기관(AHJ)이 화재위험성평가를 검토·승인할 때 쓰는 지침 (1.1)",
    "방법은 정하지 않음: “does not mandate the methods” (1.2)",
    "정량법 요건: R = Σ C_i F_i (5.5.1.2), 시나리오 군집, 불확실성(4.5), 문서(7.3.3.4), 검토(8.2·8.3)",
    "→ 외벽용 시나리오·모델·입력값·서술은 평가자가 구성해야 함 (이 자료가 그 예)",
  ]);
  card(s, MX + cw + 0.25, 1.8, cw, ch, "good", "외벽 정량 평가 = 사건수목 × 층간 확산 모델", [
    "빈도 F: 세대 화재 → 창 분출화염(IE1), 외부 발화원 → 외벽 하부(IE2)",
    "결과 C: 발화층 위로 번지는 층 수 n, 피난 전에 화염이 닿는 재실자",
    "재료 입력: t₆₀₀ = 중공층 화재가 한 층(3 m)을 오르는 시간, 플래그 B = 자립 연소",
    "산출: R = Σ F·C, 초과빈도 F(n), 기준안 대비 RR_S, 기준 충족 신뢰수준",
  ]);
  const c3 = MC("PF_in", "B1", "M1");
  card(s, MX, 1.8 + ch + 0.2, cw, ch, "poor", "예제 결과 (15층 공동주택, 예시 입력)", [
    "대책 없음: PIR(PUR) 가연성·PF 내부용 ‘불충족’, 다른 준불연 3종 ‘판단 유보’",
    `차단재 + 배리어: 준불연 4종 모두 ‘충족’ (신뢰 ${pct(Math.min(...MATS.filter((m) => m !== "PIR_C").map((m) => MC(m, "B1", "M1").conf_meets_F2)))} 이상)`,
    "가연성 등급은 대책 효과 불인정(근거 시험 없음) → ‘불충족’ 유지",
    runs("예시 기준: 발화층 위 2개 층 이상 확산 F(n ≥ 2) ≤ ", sci(best("crit_f2"), 1), "/동·년, 신뢰 90 %"),
  ], { size: 11.5 });
  card(s, MX + cw + 0.25, 1.8 + ch + 0.2, cw, ch, "conditional", "판정을 바꾸는 자료 (민감도 상위)", [
    "s (3 m 시험 → 실물 층당 시간): 실대형 시험(KS F 8414)으로 보정",
    "λ₂ · p_x (외부 발화원): 국가화재정보시스템 발화 장소(필로티·주차장) 통계",
    "p_cb · p_e1 (배리어·차단재 효과): 배리어를 넣은 중공층 시험",
    "T_FS (진압 개시): 관할 소방서 출동·전개 시간 자료",
  ]);
  box(s, MX, 6.42, CW, 0.48, { fill: HEX.ink, line: HEX.ink, r: 0.08, name: "caveat" });
  text(s, runs({ text: "주의  ", options: { bold: true } }, "빈도·분기확률·시간은 절차를 보이기 위한 예시값입니다. 측정값은 t₆₀₀과 플래그 B뿐 → 국내 통계·실대형 시험으로 바꾼 뒤 결론으로 쓰십시오."),
    { x: MX + 0.25, y: 6.42, w: CW - 0.5, h: 0.48, fontSize: 12, color: C.background1, valign: "middle" });
  addNotes(s, [
    "[스크립트]",
    "- 왼쪽 위: NFPA 551 1.2절이 직접 '방법을 정하지 않는다'고 씁니다. 그래서 '서술 방법·절차가 없다'는 느낌이 맞습니다. 대신 정량 평가가 갖춰야 할 요소(5.5)와 문서(7.3.3.4), 검토(8장)는 구체적입니다.",
    "- 오른쪽 위: 외벽에 맞춘 구성. 빈도는 사건수목, 결과는 '몇 개 층이 더 타는가'. 재료는 중공층 시험의 t600 하나로 들어갑니다(앞서 제안한 FRI와 같은 생각).",
    "- 아래: 예시 입력으로 계산하면 대책이 없을 때 가연성 PIR과 PF 내부용은 예시 기준을 넘고, 차단재·배리어를 넣으면 준불연은 기준 안으로 들어옵니다. 결론을 좌우하는 것은 s(시험→실물), 외부 발화원 빈도, 배리어 효과입니다.",
  ]);
}

// ================================================================================================ 3 NFPA 551 categories
pres.addSection({ title: "NFPA 551 · EFFECT" });
{
  const s = content("NFPA 551 · EFFECT", "NFPA 551 표 5.1.2.1 — 화재위험성평가 방법의 다섯 범주",
    "정량적 방법은 ‘가능성(빈도)과 결과를 모두 수치로’ 결합하는 방법 — 지금의 R = P × S는\nP·S가 모두 결과 쪽 지표(확산 속도·고온 지속)라 빈도 F가 빠져 있음");
  const cats = [
    ["정성적 방법", "정성", "정성", "시나리오별 결과·상대적 가능성 표, 대책 효과", "What-if, 위험 매트릭스, 위험 지수, 화재안전개념 트리"],
    ["준정량적\n가능성 방법", "정량", "정성", "화재 유형·방호 조건별 발생 빈도", "손실 통계 분석, 단독 사건수목"],
    ["준정량적\n결과 방법", "정성", "정량", "결정론적 화재 모델 결과 + 정성적 가능성", "선정한 도전적 시나리오의 구획화재 모델"],
    ["정량적 방법", "정량", "정량", "손실 기대값, 플래시오버 확률, 다른 층 사망 확률, F–N 곡선, 개인·사회 위험", "사건수목 + 화재 모델 (원전 노심손상 화재 PRA 등)"],
    ["비용편익 방법", "정량", "정량 + 비용", "위험 감소 수준별 비용, 최적 방호 수준", "확률·결과·비용 통합 계산 모델"],
  ];
  const n = cats.length, gap = 0.16, w = (CW - gap * (n - 1)) / n, y = 1.85, h = 3.55;
  cats.forEach(([name, lik, con, out, ex], i) => {
    const x = MX + i * (w + gap);
    const target = i === 3, now = i === 0 || i === 2;
    box(s, x, y, w, h, { fill: target ? "E2F3E2" : "F3F4F6", line: target ? STATUS.good.color : "F3F4F6", lw: target ? 2 : 0.5, name: `category ${i + 1}` });
    text(s, name, { x: x + 0.15, y: y + 0.12, w: w - 0.3, h: 0.62, fontSize: 14, bold: true, valign: "middle" });
    [["가능성", lik], ["결과", con]].forEach(([k, v], j) => {
      const yy = y + 0.85 + j * 0.36;
      text(s, k, { x: x + 0.15, y: yy, w: 0.75, h: 0.3, fontSize: 11, color: C.text2, valign: "middle" });
      box(s, x + 0.9, yy + 0.02, w - 1.05, 0.27, { fill: v.startsWith("정량") ? HEX.ink : "FFFFFF", line: HEX.ink, r: 0.05 });
      text(s, v, { x: x + 0.9, y: yy + 0.02, w: w - 1.05, h: 0.27, fontSize: 11, bold: true, align: "center", valign: "middle", color: v.startsWith("정량") ? C.background1 : C.text1 });
    });
    text(s, "출력", { x: x + 0.15, y: y + 1.66, w: w - 0.3, h: 0.25, fontSize: 10.5, bold: true, color: C.text2 });
    text(s, out, { x: x + 0.15, y: y + 1.9, w: w - 0.3, h: 0.85, fontSize: 11 });
    text(s, "예", { x: x + 0.15, y: y + 2.75, w: w - 0.3, h: 0.25, fontSize: 10.5, bold: true, color: C.text2 });
    text(s, ex, { x: x + 0.15, y: y + 2.98, w: w - 0.3, h: 0.55, fontSize: 10.5, color: C.text2 });
    if (now) chip(s, x + 0.15, y + h + 0.12, w - 0.3, "fair", i === 0 ? "현재: R = P×S" : "현재: 중공층 시험");
    if (target) chip(s, x + 0.15, y + h + 0.12, w - 0.3, "good", "목표: 이 자료의 절차");
  });
  box(s, MX, 6.05, CW, 0.82, { fill: "FFFFFF", line: "D1D5DB", r: 0.08, name: "definition" });
  text(s, para([
    [["NFPA 551 3.3.14 위험 = 확률과 결과의 쌍. ", { bold: true }], ["R = P × S의 P는 확산 등급(속도)이고 S는 고온 지속시간이라 둘 다 ‘결과’입니다 → 그대로 두고, 사건이 얼마나 자주 일어나는지(F)를 사건수목으로 더합니다.", {}]],
    [["정량적 방법의 출력(표 5.1.2.1) 중 외벽에 맞는 것: ", { bold: true }], ["(3) 다른 층의 피해 확률, (4) 빈도–규모 곡선(F–N), (7) 개인위험도 → 이 예제의 초과빈도 F(n)·R_P·IR", {}]],
  ], { space: 2 }), { x: MX + 0.2, y: 6.08, w: CW - 0.4, h: 0.78, fontSize: 11.5 });
  addNotes(s, [
    "[스크립트]",
    "- 표 5.1.2.1의 다섯 범주입니다. 가르는 기준은 가능성과 결과를 정성으로 다루는지 정량으로 다루는지입니다.",
    "- 지금까지 한 일: 중공층 시험은 결과를 수치로 주므로 '준정량적 결과' 쪽이고, R = P×S는 등급을 곱한 위험 지수라 정성 범주(위험 지수, 5.2.4)에 가깝습니다.",
    "- R = P×S의 P는 이름이 '가능성'처럼 보이지만 실제로는 확산 속도 등급입니다. NFPA 551 정의(3.3.14)로는 위험은 확률과 결과의 쌍이므로, 정량법으로 가려면 빈도를 더해야 합니다.",
  ]);
}

// ================================================================================================ 4 gap table
{
  const s = content("NFPA 551 · EFFECT", "NFPA 551이 주는 것과 외벽 평가에 비어 있는 것",
    "NFPA 551은 정량법의 ‘형식’(식·군집·불확실성·문서·검토)은 정해 두었지만 외벽용 시나리오,\n분기 구조, 결과 모델, 기준값은 비워 둠 → 이 자료에서 채운 내용");
  const rows = [
    ["위험의 정의", "R_t = Σ C_i F_i (식 5.5.1.2), 결과가 여럿이면 Σ Σ C_ij F_i (5.5.1.3)", "외벽 결과 지표 C → 발화층 위 확산 층 수 n, 노출 인원, (선택) 사망·재산"],
    ["시나리오", "대표성 있는 시나리오 군집, 제외 근거 문서화 (5.1.1.5, 5.5.1.1, 7.2.2.6)", "외벽 초기사건 → IE1 세대 화재 → 창 분출, IE2 외부 발화원(필로티·차량·쓰레기) — EFFECT 발화원 항목"],
    ["빈도", "통계, 사건수목, 경로 확률 = 분기 확률의 곱 (5.3, 5.3.1.2.2)", "분기 구조와 값 → 스프링클러·최성기·창 분출·중공층 노출 분기, 값은 예시 + 대체할 통계 명시"],
    ["결과", "결정론적 화재 모델, 적용 한계, 검증 (5.4, 5.4.4.1, 5.4.9)", "외벽 확산 모델 → t₆₀₀ = 층당 확산시간, 진압 T_FS·배리어 p_cb·창→창 도약과 경쟁"],
    ["허용기준", "지표 형태(상대·절대), 사전협의에서 결정 (4.4.3, A.5.1.2)", "기준값 → 예시: F(n ≥ 2) ≤ 1×10⁻⁴/동·년, 신뢰수준 90 % (협의로 확정)"],
    ["불확실성", "우연적·인식론적 구분, 전파, 민감도 (4.5, 5.5.6, A.5.5.6)", "방법 → 2중 몬테카를로(인식론적 2,000 × 우연적 400), 순위상관 민감도"],
    ["문서 · 검토", "개념보고서(7.2), 정량 결과 문서(7.3.3.4), 유지관리(7.4), 검증·확인(8.2), 검토 질문(8.3)", "서술 방법 → 서술양식.docx(항목별 쓸 내용·예문), 계산서 8_문서화점검 시트"],
  ];
  const head = [TH("요소"), TH("NFPA 551이 주는 것 (조항)"), TH("외벽에서 비어 있던 것 → 이 자료에서 채운 것")];
  const body = rows.map(([a, b, c]) => [TD(a, { bold: true, align: "center" }), TD(b), TD(c, { fill: { color: "F3F7F2" } })]);
  table(s, [head, ...body], { x: MX, y: 1.85, w: CW, colW: [1.55, 5.0, 5.783], fontSize: 11, border: BORDER, rowH: [0.36, ...rows.map(() => 0.6)], margin: [0.04, 0.08, 0.04, 0.08] });
  NOTE(s, "NFPA 551 1.2: “This guide does not mandate the methods for use in demonstrating acceptable risk; rather, it describes the technical review process and documentation …”", 6.55);
  addNotes(s, [
    "[스크립트]",
    "- 왼쪽 열이 NFPA 551이 실제로 주는 것, 오른쪽 열이 외벽에 적용하려고 이 자료에서 채운 것입니다.",
    "- 핵심은 세 가지 빈칸입니다: 외벽 시나리오 목록, 분기 구조와 확률값, 외벽 확산을 계산하는 결과 모델. 기준값은 원래 협의 대상이라 예시만 둡니다.",
  ]);
}

// ================================================================================================ 5 EFFECT
{
  const s = content("NFPA 551 · EFFECT", "EFFECT (Exterior Façade Fire Evaluation Comparison Tool)와의 연계",
    "EFFECT는 가연성 외벽 건물의 우선순위를 정하는 2단계 ‘정성’ 도구 — 그 요인들을 정량 평가의\n초기사건 · 분기확률 · 결과 입력으로 옮기면 선별(EFFECT) → 정량(QFRA)으로 이어짐");
  const lx = MX, lw = 5.25;
  box(s, lx, 1.85, lw, 4.25, { fill: "F3F4F6", line: "F3F4F6", r: 0.08, name: "effect overview" });
  text(s, "EFFECT 개요 (공개 자료 기준)", { x: lx + 0.2, y: 1.95, w: lw - 0.4, h: 0.36, fontSize: 14, bold: true });
  text(s, para([
    "NFPA 후원, Arup이 방법론 개발, FPRF가 검토 주관, Jensen Hughes 동료 검토 (2018)",
    "목적: 가연성 외벽 고층 건물(18 m 초과, 주거·업무) 포트폴리오에서 점검·개선 우선순위",
    "Tier 1: 짧은 객관식 문항 — 외장재·단열재 가연성, 스프링클러, 발화원, 경보설비 형식 → 건물 순위",
    "Tier 2: 우선 건물 상세 평가 — 현장조사, 준공도면, 유지관리 기록, 시료 채취, 실험실 시험",
    "발화원: 실내 화재의 창 분출, 발코니, 차량, 쓰레기 용기",
    "성격: 정성적 위험 순위 (NFPA 551 표 5.1.2.1의 정성 범주)",
  ], { space: 5 }), { x: lx + 0.2, y: 2.35, w: lw - 0.4, h: 3.7, fontSize: 12 });
  const rx = MX + lw + 0.25, rw = CW - lw - 0.25;
  const map = [
    ["외장재·단열재 가연성", "결과 모델의 재료 입력", "t₆₀₀ → τ_F, 플래그 B → p_B, 등급(c/sc)"],
    ["발화원 (실내→창, 발코니, 차량, 쓰레기)", "초기사건과 빈도", "IE1: λ₁·p_FO·p_v  /  IE2: λ₂·p_x"],
    ["건물 특성 (높이·용도·재실)", "결과의 크기", "N, n_u, 재실자, 확산 띠 세대 n_a"],
    ["방화설비 (스프링클러·경보)", "분기확률 · 피난 시간", "p_s, RSET"],
    ["Tier 2 현장조사·시료·시험", "인식론적 불확실성 축소", "p_e (차단재), p_cb (배리어), s (시험)"],
  ];
  table(s, [[TH("EFFECT 요인군"), TH("정량 평가에서의 역할"), TH("파라미터")], ...map.map(([a, b, c]) => [TD(a, { bold: true }), TD(b), TD(c)])],
    { x: rx, y: 1.85, w: rw, colW: [2.45, 2.0, 2.633], fontSize: 11, border: BORDER, rowH: [0.36, 0.62, 0.62, 0.5, 0.5, 0.62], margin: [0.04, 0.08, 0.04, 0.08] });
  const flow = [["EFFECT Tier 1", "정성 선별: 포트폴리오 순위"], ["EFFECT Tier 2", "상세 자료: 현장·시료·시험"], ["정량 평가 (이 절차)", "R = Σ F·C, 기준 판정"]];
  const fw = (rw - 2 * 0.35) / 3;
  flow.forEach(([a, b], i) => {
    const x = rx + i * (fw + 0.35), y = 5.25;
    box(s, x, y, fw, 0.85, { fill: i === 2 ? "E2F3E2" : "FFFFFF", line: i === 2 ? STATUS.good.color : "9CA3AF", lw: 1, name: `flow ${i + 1}` });
    text(s, a, { x: x + 0.1, y: y + 0.06, w: fw - 0.2, h: 0.34, fontSize: 12.5, bold: true, align: "center", valign: "middle" });
    text(s, b, { x: x + 0.1, y: y + 0.42, w: fw - 0.2, h: 0.36, fontSize: 10.5, color: C.text2, align: "center", valign: "middle" });
    if (i < 2) arrow(s, x + fw + 0.04, y + 0.42, x + fw + 0.31, y + 0.42);
  });
  box(s, MX, 6.27, CW, 0.6, { fill: STATUS.conditional.tint, line: STATUS.conditional.tint, r: 0.08, name: "effect note" });
  text(s, runs({ text: "EFFECT 사용자 가이드 원문은 이번에 받지 못했습니다.  ", options: { bold: true } },
    "위 내용은 NFPA·Arup의 공개 발표 기준입니다. 원문(EFFECT_Users Guide_final)을 주시면 문항·선택지·가중치 단위로 이 대응표와 계산서 입력을 맞춰 갱신합니다."),
    { x: MX + 0.2, y: 6.27, w: CW - 0.4, h: 0.6, fontSize: 11.5, valign: "middle" });
  addNotes(s, [
    "[스크립트]",
    "- EFFECT는 NFPA가 그렌펠 이후 내놓은 외벽 위험 순위 도구로, Tier 1은 짧은 설문으로 포트폴리오 순위를, Tier 2는 현장 자료로 상세 평가를 합니다. 방법 자체는 정성적입니다.",
    "- 그래서 EFFECT를 '대체'하는 것이 아니라, EFFECT가 보는 요인을 정량 평가의 입력 자리로 옮기는 대응표를 만들었습니다(오른쪽).",
    "- 흐름: EFFECT로 대상 건물·대안을 고르고(선별), 그 대상에 이 정량 절차를 적용합니다. NFPA 551 5.1.3.1도 선별 목적이면 정성, 확정 평가면 정량을 권합니다.",
    "",
    "[질문 대비]",
    "- 원문 미수령: 업로드 폴더에 파일이 없었고, 이 환경의 네트워크 정책상 nfpa.org 등에 접속할 수 없었습니다. 원문을 다시 올려 주시면 반영합니다.",
  ]);
}

// ================================================================================================ 6 procedure
pres.addSection({ title: "12단계 절차" });
{
  const s = content("12단계 절차", "외벽 정량적 화재위험성평가 — 12단계 절차 (NFPA 551 그림 4.3 흐름을 외벽에 구체화)",
    "준비 → 분석 → 문서화·검토의 12단계 — 단계마다 조항과 산출물을 정해 두면\n그대로 보고서 목차와 검토 대응표가 됨");
  const bands = [
    ["준비", 1.85, 1.2, [
      ["1", "사전협의 · 문제 정의", "4.3, 4.4.1", "목적 · 대상 · 범위 · 이해관계자"],
      ["2", "위험요소 정의", "4.4.2", "노출 대상 · 자극 · 전달 · 반응"],
      ["3", "허용기준 설정", "4.4.3, A.5.1.2", "지표 · 기준값 · 신뢰수준"],
      ["4", "방법 선정", "4.4.4, 5.1.3, 표 5.1.2.1", "범주 · 모델 · 적합성 근거"],
    ]],
    ["분석", 3.27, 1.62, [
      ["5", "시나리오 선정", "5.1.1.5, 5.5.1.1", "초기사건 · 군집 · 제외 근거"],
      ["6", "빈도 분석", "5.3", "사건수목 · 경로별 F"],
      ["7", "결과 분석", "5.4, 5.4.9", "층간 확산 C (시험·모델)"],
      ["8", "위험도 산정", "5.5.1", "R = Σ F·C · 초과빈도"],
      ["9", "불확실성 분석", "4.5, 5.5.6", "분포 · 신뢰수준 · 민감도"],
      ["10", "검증 · 확인", "5.5.9, 8.2", "재현 계산 · 시험 비교"],
    ]],
    ["문서화\n· 검토", 5.11, 1.05, [
      ["11", "문서화", "7.2, 7.3.3.4, 7.4", "개념보고서 · 정량 결과 문서 · 유지관리 매뉴얼"],
      ["12", "검토 · 승인", "8.1, 8.3, 4.3.2.3", "검토 질문 9개 대응 · 승인 또는 보완"],
    ]],
  ];
  const bx = MX + 1.15, bw = CW - 1.15, gap = 0.16;
  bands.forEach(([name, y, h, steps]) => {
    box(s, MX, y, 1.0, h, { fill: HEX.ink, line: HEX.ink, r: 0.06, name: `band ${name}` });
    text(s, name, { x: MX, y, w: 1.0, h, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle" });
    const w = (bw - gap * (steps.length - 1)) / steps.length;
    steps.forEach(([n, t, ref, out], i) => {
      const x = bx + i * (w + gap);
      box(s, x, y, w, h, { fill: "FFFFFF", line: "9CA3AF", lw: 0.75, name: `step ${n}` });
      badge(s, n, x + 0.1, y + 0.1);
      text(s, t, { x: x + 0.47, y: y + 0.06, w: w - 0.55, h: 0.4, fontSize: 12.5, bold: true, valign: "middle" });
      text(s, ref, { x: x + 0.12, y: y + 0.5, w: w - 0.22, h: 0.27, fontSize: 10.5, color: C.text2 });
      text(s, "→ " + out, { x: x + 0.12, y: y + 0.77, w: w - 0.22, h: h - 0.82, fontSize: 11 });
      if (i < steps.length - 1) arrow(s, x + w + 0.01, y + 0.28, x + w + gap - 0.01, y + 0.28, { width: 1 });
    });
  });
  arrow(s, MX + 0.5, 3.07, MX + 0.5, 3.25);
  arrow(s, MX + 0.5, 4.91, MX + 0.5, 5.09);
  NOTE(s, "번호는 이 덱의 7–20쪽, 서술양식.docx의 절, 계산서(xlsx)의 시트 번호와 같은 흐름입니다. 결과가 기준을 넘으면 4·7단계로 돌아가 대책(차단재·배리어 등)을 넣고 다시 계산합니다.", 6.3, { h: 0.5 });
  addNotes(s, [
    "[스크립트]",
    "- NFPA 551 그림 4.3의 검토 흐름(범위·개발 → FRA 수행 → 검토·평가)을 외벽 정량 평가의 작업 단계로 풀어 쓴 것입니다.",
    "- 각 칸의 아래 줄이 그 단계의 산출물입니다. 산출물을 모으면 7장의 문서 요건이 채워지도록 짰습니다.",
    "- 기준을 못 넘으면 대책을 넣고 7단계(결과)부터 다시 계산하는 반복 구조입니다.",
  ]);
}

// ================================================================================================ 7 steps 1-4
{
  const s = content("12단계 절차", "단계 1–4 · 준비 — 외벽에서 정할 것과 예제에서 정한 것",
    "준비 단계의 산출물은 ‘무엇을, 어디까지, 어떤 지표와 기준으로’ — 특히 허용기준은 분석 전에\n관할기관과 합의해야 함(4.4.3.4)");
  const rows = [
    ["1 문제 정의\n4.3, 4.4.1", "평가 목적(대안 비교 · 기존 건물 우선순위 · 동등성 입증), 대상 건물과 외벽 시스템, 평가 범위와 제외 범위, 이해관계자(4.2)",
      `중공층 외단열 5종 × 대책 2안 비교 · ${D.params.find((p) => p.key === "N").best}층 공동주택(세대 ${best("N") * best("units")}, 재실 ${D.occupants}명) · 사용 단계의 외벽 경로만(실내 연기 확산·공사 중 화재 제외)`],
    ["2 위험요소\n4.4.2", "노출 대상(4.4.2.1) · 화재 자극(열·연기·가스, 4.4.2.2) · 전달 현상(4.4.2.3) · 대상의 반응과 기준(4.4.2.4)",
      "발화층 위 재실자 · 화염과 열(낙하물 포함) · 창 분출 → 외장·중공층 → 수직 확산 → 창으로 재유입 · 피난 완료(RSET) 전에 화염이 닿는가"],
    ["3 허용기준\n4.4.3, A.5.1.2", "지표(확산 범위 · 개인위험도 · 기준안 대비), 기준값, 불확실성을 판정에 넣는 방식(신뢰수준). 규정·성능기준·합의값·표준 중 근거(4.4.3.5)",
      runs("주 지표: F(n ≥ 2) ≤ ", sci(best("crit_f2"), 1), "/동·년 (A.5.1.2(6) 확산 범위), 신뢰수준 90 % — 예시값 · 보조: 개인위험도 IR(HSE 10", { text: "−6", options: { superscript: true } }, "·10", { text: "−4", options: { superscript: true } }, " 참고), 불연 기준안 대비 RR_S")],
    ["4 방법\n4.4.4, 5.1.3", "범주(표 5.1.2.1), 모델과 입력 자료, 적용 범위, 선택 근거, 사용할 소프트웨어(6.2.4.2)",
      "정량적 방법 = 사건수목 + 층간 확산 결과 모델 + 2중 몬테카를로 · 근거: 중공층 시험(층당 시간)과 화재 통계(빈도)를 직접 결합 · 파이썬 qfra.py, 엑셀 계산서"],
  ];
  const body = rows.map(([a, b, c]) => [TD(a, { bold: true, align: "center" }), TD(b), TD(c, { fill: { color: "F3F7F2" } })]);
  table(s, [[TH("단계"), TH("외벽에서 정할 것"), TH("예제에서 정한 것")], ...body],
    { x: MX, y: 1.85, w: CW, colW: [1.7, 5.0, 5.633], fontSize: 11.5, border: BORDER, rowH: [0.36, 1.0, 1.0, 1.15, 1.0], margin: [0.05, 0.1, 0.05, 0.1] });
  NOTE(s, "허용기준은 예시값입니다. 실제 기준은 규정·관할기관 협의로 정하고(4.4.3.3), 개념보고서(7.2.2.4)에 근거와 함께 적습니다.", 6.55);
  addNotes(s, [
    "[스크립트]",
    "- 1단계는 목적을 하나로 정하는 것이 중요합니다. '재료 대안 비교'인지 '기존 건물 우선순위'인지에 따라 기준안과 지표가 달라집니다.",
    "- 3단계 허용기준은 분석 전에 정합니다. 결과를 보고 기준을 고르면 검토에서 받아들여지지 않습니다. 이 예제는 A.5.1.2의 '확산 범위' 지표를 주 지표로 썼습니다.",
    "",
    "[질문 대비]",
    "- HSE R2P2(2001): 개인위험도 10^-6/년 이하는 '일반적으로 허용', 10^-4/년은 일반인 '허용 한계'.",
  ]);
}

// ================================================================================================ 8 step 5 scenarios
{
  const s = content("12단계 절차", "단계 5 · 시나리오 선정과 군집 (5.1.1.5, 5.5.1.1, 7.2.2.6)",
    "외벽을 화염에 노출시키는 초기사건 두 가지로 군집하고, 발화층·방호 상태·재료 반복시험으로\n세분 — 제외한 시나리오는 이유와 함께 문서에 남김");
  const iw = 4.1;
  const ies = [
    ["IE1", "세대 화재 → 창 분출화염", ["발화층 1–15층 균등", "스프링클러 · 최성기 · 창 파손 분기를 거쳐 외벽 노출", "창→창 도약 확산이 함께 일어남", "EFFECT: 실내 화재의 창 분출"]],
    ["IE2", "외부 발화원 → 외벽 하부", ["필로티·주차 차량, 쓰레기, 외부 설비", "1층에서 시작, 세대 스프링클러와 무관", "EFFECT: 차량 · 쓰레기 용기 · 발코니", "국내 사례: 의정부(2015) · 제천(2017) — 1층 필로티 발화 후 외벽 확산"]],
  ];
  ies.forEach(([tag, name, items], i) => {
    const y = 1.85 + i * 2.25;
    box(s, MX, y, iw, 2.1, { fill: "F3F4F6", line: "F3F4F6", r: 0.08, name: tag });
    box(s, MX + 0.18, y + 0.16, 0.7, 0.36, { fill: HEX.ink, line: HEX.ink, r: 0.05 });
    text(s, tag, { x: MX + 0.18, y: y + 0.16, w: 0.7, h: 0.36, fontSize: 13, bold: true, color: C.background1, align: "center", valign: "middle" });
    text(s, name, { x: MX + 1.0, y: y + 0.13, w: iw - 1.1, h: 0.42, fontSize: 14, bold: true, valign: "middle" });
    text(s, para(items, { space: 3 }), { x: MX + 0.2, y: y + 0.62, w: iw - 0.35, h: 1.45, fontSize: 11.5 });
  });
  const mx2 = MX + iw + 0.25, mw = 3.6;
  box(s, mx2, 1.85, mw, 4.35, { fill: "FFFFFF", line: "D1D5DB", r: 0.08, name: "clusters" });
  text(s, "군집 기준 (5.5.1.4)", { x: mx2 + 0.2, y: 1.95, w: mw - 0.4, h: 0.36, fontSize: 14, bold: true });
  text(s, para([
    "초기사건 IE1 · IE2",
    "발화층: IE1은 15개 층 균등(가중치 w_j = (N − j)/N), IE2는 1층",
    "방호 상태: 스프링클러(B1/B0), 개구부 차단재·층간 배리어(M0/M1)",
    "재료: 반복시험 하나하나를 같은 무게의 사건으로(우연적 변동)",
    "군집 빈도 = 군집에 든 시나리오 빈도의 합 (5.1.1.5.3)",
  ], { space: 5 }), { x: mx2 + 0.2, y: 2.38, w: mw - 0.4, h: 3.75, fontSize: 11.5 });
  const rx = mx2 + mw + 0.25, rw = CW - iw - mw - 0.5;
  const excl = [
    ["공사 중 화재", "사용 단계 평가 범위 밖 → 별도 평가"],
    ["방화 · 다중 발화", "비현실적으로 가혹 (5.5.1.1)"],
    ["발코니 화재", "IE3로 추가 가능(발화층 균등, 스프링클러 무관) — 예제 생략"],
    ["인접 건물 복사열 · 강풍", "별도 군집으로 추가 가능"],
    ["실내 연기 확산", "외벽 경로 밖 → 실내 FRA와 결합"],
  ];
  table(s, [[TH("제외 · 보류한 시나리오"), TH("근거")], ...excl.map(([a, b]) => [TD(a, { bold: true, fontSize: 11 }), TD(b, { fontSize: 11 })])],
    { x: rx, y: 1.85, w: rw, colW: [1.6, rw - 1.6], fontSize: 11, border: BORDER, rowH: [0.36, 0.72, 0.72, 0.85, 0.72, 0.72], margin: [0.04, 0.08, 0.04, 0.08] });
  NOTE(s, "시나리오를 너무 줄이면 위험을 낮게 봅니다(5.5.8). 제외는 ‘비현실적으로 가혹’ 또는 ‘충분히 드묾’일 때만, 근거를 적어서(7.2.2.6).", 6.4, { h: 0.45 });
  addNotes(s, [
    "[스크립트]",
    "- 외벽 위험의 출발점은 '외벽이 화염에 노출되는 사건'입니다. EFFECT가 보는 발화원도 실내 화재의 창 분출과 외부 화원(차량·쓰레기·발코니)으로 나뉩니다.",
    "- IE2는 국내에서 의정부·제천처럼 1층 필로티에서 시작해 외벽을 타고 올라간 사례가 있어 따로 둡니다. 세대 스프링클러로 막을 수 없는 경로라는 점이 중요합니다.",
    "- 제외 목록은 반드시 문서에 남깁니다. 검토자는 '왜 이 시나리오가 없는가'를 묻습니다.",
  ]);
}

// ================================================================================================ 9 step 6 event tree
{
  const s = content("12단계 절차", "단계 6 · 빈도 분석 — 사건수목 (5.3, 경로 확률 = 분기 확률의 곱)",
    `외벽이 화염에 노출되는 빈도: IE1 ${supStr(PT("PF_in", "B1", "M0").F_IE1_ext)}, IE2 ${supStr(PT("PF_in", "B1", "M0").F_IE2_ext)} /동·년 (예시 입력)\n— 세대 스프링클러는 IE1만 줄이므로 외부 발화원이 더 큰 몫을 차지함`);
  const p = (k) => best(k);
  const U = best("N") * best("units");
  const f0 = best("lam1") * U, f1 = f0 * (1 - p("p_spr")), f2 = f1 * p("p_fo"), f3 = f2 * p("p_vent");
  const g0 = best("lam2"), g1 = g0 * p("p_x2");
  const nw = 2.05, nh = 0.98, gx = 0.33, xs = [0, 1, 2, 3, 4].map((i) => MX + i * (nw + gx));
  function node(x, y, w, h, title, formula, value, o = {}) {
    box(s, x, y, w, h, { fill: o.fill || "FFFFFF", line: o.line || "9CA3AF", lw: o.lw || 0.75, name: `node ${title}` });
    text(s, title, { x: x + 0.08, y: y + 0.04, w: w - 0.16, h: 0.3, fontSize: 11.5, bold: true, align: "center", valign: "middle" });
    text(s, formula, { x: x + 0.08, y: y + 0.34, w: w - 0.16, h: 0.27, fontSize: 10.5, color: C.text2, align: "center", valign: "middle" });
    text(s, value, { x: x + 0.08, y: y + 0.62, w: w - 0.16, h: 0.32, fontSize: 13, bold: true, align: "center", valign: "middle" });
  }
  const yr = (v) => runs(sci(v, 2), "/년");
  // IE1
  const y1 = 2.25;
  text(s, "IE1", { x: MX, y: 1.8, w: 0.6, h: 0.3, fontSize: 13, bold: true, color: C.text2 });
  node(xs[0], y1, nw, nh, "세대 화재", "λ₁ × 세대 수 U", yr(f0), { fill: "F3F4F6" });
  node(xs[1], y1, nw, nh, "스프링클러 제어 실패", "× (1 − p_s)", yr(f1));
  node(xs[2], y1, nw, nh, "최성기 도달", "× p_FO", yr(f2));
  node(xs[3], y1, nw, nh, "창 분출 = F₁", "× p_v", yr(f3), { line: HEX.ink, lw: 1.5 });
  [[0, 1, "제어 → 실내 국한"], [1, 2, "최성기 전 소화"], [2, 3, "창 유지 → 실내"]].forEach(([a, b, lab]) => {
    arrow(s, xs[a] + nw, y1 + nh / 2, xs[b], y1 + nh / 2);
    text(s, lab, { x: xs[a] + nw - 0.4, y: y1 + nh + 0.05, w: gx + 0.8, h: 0.25, fontSize: 9.5, color: C.text2, align: "center" });
  });
  const sw = CW - 4 * (nw + gx);
  node(xs[4], 1.8, sw, 0.9, "중공층 노출 → 외벽 + 도약", "× p_e0", yr(f3 * p("p_exp0")), { fill: "FCE6DC", line: "EC835A" });
  node(xs[4], 2.85, sw, 0.9, "노출 안 됨 → 도약만", "× (1 − p_e0)", yr(f3 * (1 - p("p_exp0"))), { fill: "F3F4F6" });
  arrow(s, xs[3] + nw, y1 + nh / 2, xs[4], 2.25);
  arrow(s, xs[3] + nw, y1 + nh / 2, xs[4], 3.3);
  // IE2
  const y2 = 4.55;
  text(s, "IE2", { x: MX, y: 4.1, w: 0.6, h: 0.3, fontSize: 13, bold: true, color: C.text2 });
  node(xs[0], y2, nw, nh, "외부 발화원 화재", "λ₂", yr(g0), { fill: "F3F4F6" });
  node(xs[1], y2, nw, nh, "외벽 화염 노출 = F₂", "× p_x", yr(g1), { line: HEX.ink, lw: 1.5 });
  arrow(s, xs[0] + nw, y2 + nh / 2, xs[1], y2 + nh / 2);
  text(s, "떨어져 있음 → 영향 없음", { x: xs[0] + nw - 0.4, y: y2 + nh + 0.05, w: gx + 0.8, h: 0.25, fontSize: 9.5, color: C.text2, align: "center" });
  node(xs[2], 4.1, nw, 0.9, "중공층 노출 → 외벽", "× p_e0", yr(g1 * p("p_exp0")), { fill: "FCE6DC", line: "EC835A" });
  node(xs[2], 5.15, nw, 0.9, "노출 안 됨 → 확산 없음", "× (1 − p_e0)", yr(g1 * (1 - p("p_exp0"))), { fill: "F3F4F6" });
  arrow(s, xs[1] + nw, y2 + nh / 2, xs[2], 4.55);
  arrow(s, xs[1] + nw, y2 + nh / 2, xs[2], 5.6);
  // notes panel
  const nx = xs[3], nwid = CW - 3 * (nw + gx);
  box(s, nx, 4.1, nwid, 1.95, { fill: "FFFFFF", line: "D1D5DB", r: 0.08, name: "tree notes" });
  text(s, para([
    runs({ text: "예시 입력: ", options: { bold: true } }, "λ₁ = ", sci(best("lam1"), 1), "/세대·년, U = 60, p_s = 0.90, p_FO = 0.15, p_v = 0.70, λ₂ = ", sci(best("lam2"), 1), "/동·년, p_x = 0.30, p_e0 = 0.90"),
    runs({ text: "B0 (스프링클러 없음): ", options: { bold: true } }, "F₁ = ", sci(f3 / (1 - p("p_spr")), 2), "/년 (10배)"),
    runs({ text: "M1 (개구부 차단재): ", options: { bold: true } }, "p_e0 0.90 → p_e1 0.30, 준불연만 인정"),
    runs({ text: "가정: ", options: { bold: true } }, "분기는 서로 독립(5.3.1.2.2) — 독립이 아니면 조건부 확률로 넣음"),
  ], { space: 3 }), { x: nx + 0.18, y: 4.17, w: nwid - 0.3, h: 1.85, fontSize: 11 });
  NOTE(s, "빨간 칸(중공층 노출)이 다음 단계 결과 모델의 ‘외벽 경로’로, 회색 칸은 ‘도약만’ 또는 ‘확산 없음’으로 이어집니다. 값은 계산서 3_사건수목과 같습니다.", 6.38, { h: 0.45 });
  addNotes(s, [
    "[스크립트]",
    "- 위가 IE1, 아래가 IE2입니다. 왼쪽에서 오른쪽으로 분기 확률을 곱하면 각 끝 칸의 연간 빈도가 됩니다(5.3.1.2.2).",
    "- 스프링클러가 있는 건물에서는 IE1이 1/10로 줄어 외벽 노출 빈도가 IE2보다 작아집니다. 그래서 외부 발화원(필로티·주차)이 외벽 위험의 큰 몫이 됩니다. 단, 이것은 예시 입력의 결과이므로 국내 통계로 확인해야 합니다.",
    "- 빨간 칸만 외벽(중공층) 경로로 넘어가고, 나머지는 창→창 도약만 일어나거나 확산이 없습니다.",
  ]);
}

// ================================================================================================ 10 step 7 model
{
  const s = content("12단계 절차", "단계 7 · 결과 분석 — 층간 확산 모델 (중공층 t₆₀₀ = 층당 확산시간)",
    "중공층 화재는 τ_F = s·t₆₀₀마다 한 층씩 오르고 진압·배리어·화원 소멸에서 멈춤\n— 창→창 도약과 겨뤄, 먼저 닿은 층에 피난 전 재실자가 있으면 ‘노출’");
  // schematic: 6 storeys, origin at the 2nd from the bottom
  const bx = MX + 0.1, bw = 2.3, sh = 0.7, top = 1.95, nS = 6, cav = 0.2;
  const yOf = (i) => top + (nS - 1 - i) * sh;            // i = 0 bottom storey
  const labels = ["k − 1", "k (발화층)", "k + 1", "k + 2", "k + 3", "k + 4"];
  for (let i = 0; i < nS; i++) {
    const y = yOf(i);
    s.addShape(pres.ShapeType.rect, { x: bx, y, w: bw, h: sh, fill: { color: i === 1 ? "FCE6DC" : "FFFFFF" }, line: { color: "9CA3AF", width: 0.75 } });
    s.addShape(pres.ShapeType.rect, { x: bx + bw - 0.62, y: y + 0.18, w: 0.42, h: 0.34, fill: { color: "DCEBFA" }, line: { color: "9CA3AF", width: 0.5 } });
    text(s, labels[i], { x: bx + 0.1, y: y + 0.18, w: 1.4, h: 0.34, fontSize: 10.5, color: C.text2, valign: "middle" });
  }
  // cavity strip and cladding
  s.addShape(pres.ShapeType.rect, { x: bx + bw, y: top, w: cav, h: nS * sh, fill: { color: "FDF0D5" }, line: { color: "FAB219", width: 0.75 }, objectName: "cavity" });
  s.addShape(pres.ShapeType.line, { x: bx + bw + cav + 0.04, y: top, w: 0, h: nS * sh, line: { color: HEX.ink2, width: 2 } });
  // barriers at storey boundaries
  for (let i = 1; i < nS; i++) s.addShape(pres.ShapeType.rect, { x: bx + bw, y: yOf(i) + sh - 0.03, w: cav, h: 0.06, fill: { color: HEX.ink }, line: { color: HEX.ink, width: 0.5 } });
  // flame at origin window
  s.addShape(pres.ShapeType.triangle, { x: bx + bw + 0.22, y: yOf(1) + 0.02, w: 0.42, h: 0.62, fill: { color: "EB6834" }, line: { color: "EB6834", width: 0.5 }, objectName: "flame" });
  // cavity climb arrow
  arrow(s, bx + bw + cav / 2, yOf(1) + 0.4, bx + bw + cav / 2, yOf(4) + 0.15, { color: "D03B3B", width: 2.5 });
  // leapfrog arrow (window k -> window k+1)
  arrow(s, bx + bw - 0.2, yOf(1) + 0.18, bx + bw - 0.2, yOf(2) + 0.5, { color: "2A78D6", width: 1.75, dash: true });
  // labels on the right of the façade
  const lx = bx + bw + cav + 0.6, lw = 2.35;
  const lab = (y, t, color) => text(s, t, { x: lx, y, w: lw, h: 0.62, fontSize: 10.5, color: color || C.text1 });
  lab(yOf(4) - 0.05, "중공층 경로: 층당 τ_F = s·t₆₀₀", "B13232");
  lab(yOf(3) + 0.05, "층 경계 배리어: 막을 확률 p_cb (M1)");
  lab(yOf(2) + 0.1, "창→창 도약: 층당 p_L, τ_L", "1F5FAE");
  lab(yOf(1) + 0.08, "분출화염 (IE1) 또는 외부 화원 (IE2)", "B35806");
  lab(yOf(0) + 0.05, "진압 T_FS · 피난 RSET와 경쟁");
  // equations
  const ex = MX + 5.75, ew = CW - 5.75;
  box(s, ex, 1.85, ew, 4.45, { fill: "F3F4F6", line: "F3F4F6", r: 0.08, name: "equations" });
  const eq = [
    [["층당 확산시간  ", { bold: true }], ["τ_F = s · t₆₀₀ · h / 3 m   (s: 실물/시험 보정 0.25–1, 가정)", {}]],
    [["중공층 경로가 오르는 층 수  ", { bold: true }], ["n_F = 내림[ min(T_FS, T_stop) ÷ τ_F ],  T_stop = t_e (비자립) 또는 ∞ (자립: 플래그 B)", {}]],
    [["배리어  ", { bold: true }], ["a_j = (1 − p_cb)", {}], ["j", { superscript: true }], ["  (j ≤ n_F일 때, 아니면 0) — 층 경계마다 독립으로 막음", {}]],
    [["창→창 도약  ", { bold: true }], ["b_j = p_L", {}], ["j", { superscript: true }], ["  (j ≤ n_L일 때), n_L = 내림[T_FS ÷ τ_L]", {}]],
    [["j개 층 이상 확산할 확률  ", { bold: true }], ["P(n ≥ j) = w_j [1 − (1 − a_j)(1 − b_j)],  w_j = (N − j)/N (IE1) 또는 1 (IE2)", {}]],
    [["노출 인원  ", { bold: true }], ["층 j에 먼저 닿은 경로의 도달시각 < RSET → 그 층의 n_a × 세대당 재실자", {}]],
    [["기존 제안과의 관계  ", { bold: true }], ["FRI = t_ref / t₆₀₀(기준시간 동안 오르는 층 수) → t_ref 대신 진압 시각 T_FS와 확률적으로 경쟁", {}]],
  ];
  text(s, para(eq, { space: 7 }), { x: ex + 0.22, y: 1.98, w: ew - 0.4, h: 4.25, fontSize: 12 });
  NOTE(s, "결과는 ‘외벽 노출 사건 1회당 기대값’입니다. 같은 식이 계산서 4_결과모델에 셀 수식(SUMPRODUCT)으로 들어 있어 값을 바꿔 가며 확인할 수 있습니다.", 6.42, { h: 0.45 });
  addNotes(s, [
    "[스크립트]",
    "- 왼쪽 그림: 발화층 k의 분출화염(또는 1층 외부 화원)이 중공층에 들어가면, 중공층 화재는 한 층을 τ_F마다 오릅니다. τ_F는 3 m 중공층 시험의 t600에 실물 보정 계수 s를 곱한 값입니다.",
    "- 멈추는 조건은 셋: 소방 진압(T_FS), 층 경계의 배리어(p_cb), 그리고 자립 연소가 안 되는 재료는 분출화염이 끝나는 시각(T_stop)입니다. 플래그 B가 있는 재료는 화원이 꺼져도 계속 오릅니다.",
    "- 창→창 도약은 불연 외벽에서도 일어나는 기본 경로라 따로 계산하고, 둘 중 먼저 닿는 쪽으로 노출을 셉니다.",
    "",
    "[질문 대비]",
    "- s를 0.25–1로 둔 이유: 30 kW 버너보다 실제 창 분출화염이 크므로 실물에서는 더 빠를 수 있다는 가정입니다. 가장 큰 불확실성이고, 실대형 시험으로 보정해야 합니다(16쪽).",
  ]);
}

// ================================================================================================ 11 tests -> inputs
{
  const s = content("12단계 절차", "단계 7 · 시험 자료가 결과 모델에 들어가는 방식",
    "재료는 t₆₀₀(층당 확산시간)과 플래그 B(자립 연소)로만 들어감 — 외벽 노출 1회당\n확산 층 수: PIR(PUR) 가연성 7.0, PF 내부용 2.6, 다른 준불연 0.7–0.9");
  const sv = best("s");
  const head = [TH("재료"), TH("등급"), TH("시험 t₆₀₀ (s)"), TH(`τ_F (s), s = ${sv}`), TH("자립 연소 k/n → p_B"),
    TH("E[확산 층 수]\n배리어 없음 / 있음"), TH("P(n ≥ 2)"), TH("E[노출 인원]")];
  const body = ORDER.map((m) => {
    if (m === "NC") {
      const c = CO("NC", "IE1", false);
      return [TD(LABEL.NC, { bold: true }), TD("불연", { align: "center" }), TD("—", { align: "center" }), TD("중공층 경로 없음", { align: "center" }),
        TD("—", { align: "center" }), TD(`${f2(c.E_storeys)} (도약만)`, { align: "center" }), TD(f2(c.P_ge2), { align: "center" }), TD(f2(c.E_persons, 1), { align: "center" })];
    }
    const ts = TESTS(m), fb = D.flag_b[m];
    const t600 = ts.map((t) => (t.t600_s === null ? "미도달" : t.t600_s.toFixed(0))).join(" / ");
    const tau = ts.map((t) => (t.t600_s === null ? "∞" : (sv * t.t600_s).toFixed(0))).join(" / ");
    const c0 = CO(m, "IE1", false), c1 = CO(m, "IE1", true);
    const noCredit = D.grade[m] === "c";
    return [TD(LABEL[m], { bold: true, color: HEX.ink }), TD(D.grade[m], { align: "center" }), TD(t600, { align: "center" }), TD(tau, { align: "center" }),
      TD(`${fb.k}/${fb.n} → ${f2(fb.p_best, 3)}`, { align: "center" }),
      TD(`${f2(c0.E_storeys)} / ${noCredit ? "인정 안 함" : f2(c1.E_storeys)}`, { align: "center", bold: true }), TD(f2(c0.P_ge2), { align: "center" }), TD(f2(c0.E_persons, 1), { align: "center" })];
  });
  body.push([TD(LABEL.EPS_C, { bold: true, color: HEX.ink2 }), TD("c", { align: "center" }), TD("기록 240 / 262 s", { align: "center", color: HEX.ink2 }),
    TD("산정 불가", { align: "center", color: HEX.ink2 }), TD("—", { align: "center" }), TD("정량평가 제외 — 30분 기록으로 재시험 필요 (6.2.1.3 자료 부족)", { colspan: 3, color: HEX.ink2 })]);
  table(s, [head, ...body], { x: MX, y: 1.85, w: CW, colW: [2.05, 0.75, 1.75, 1.7, 1.75, 1.95, 1.0, 1.383], fontSize: 11.5, border: BORDER,
    rowH: [0.6, ...ORDER.map(() => 0.42), 0.42], margin: [0.03, 0.06, 0.03, 0.06] });
  ORDER.forEach((m, i) => s.addShape(pres.ShapeType.rect, { x: MX - 0.14, y: 1.85 + 0.6 + i * 0.42 + 0.06, w: 0.08, h: 0.3, fill: { color: MAT[m] }, line: { color: MAT[m], width: 0.5 } }));
  text(s, para([
    "E[확산 층 수]·P(n ≥ 2)·E[노출 인원]: IE1(세대 화재) 외벽 노출 1회당, 점추정(T_FS = RSET = 15분, 분출 지속 20분, τ_L = 12분), 발화층 15개 층 균등, 노출 인원은 층당 5명(2세대 × 2.5명) 기준",
    "p_B = (k + ½)/(n + 1): 반복 2–3회에서 0/2를 ‘절대 안 됨’으로, 2/2를 ‘항상’으로 단정하지 않기 위한 추정(Jeffreys) · EPS 준불연 #2는 30분 안에 3000 mm 600 K에 이르지 않아 ‘다음 층으로 오르지 않는 사건’으로 넣음",
  ], { space: 4 }), { x: MX, y: 5.55, w: CW, h: 0.95, fontSize: 11, color: C.text2 });
  NOTE(s, "가연성 등급의 ‘배리어 있음’은 인정 안 함: 3 m 중공층이 30 s 만에 전면 화염이 된 재료에 배리어가 듣는다는 시험 근거가 없음(5.1.4.5 신뢰도·효과).", 6.5);
  addNotes(s, [
    "[스크립트]",
    "- 재료가 모델에 들어가는 통로는 t600과 플래그 B 두 개뿐입니다. τ_F는 t600에 s를 곱한 값으로, s = 0.5이면 PF 내부용은 층당 약 4분, PIR(PUR) 가연성은 15초입니다.",
    "- 그래서 외벽이 노출된 한 번의 사건에서 PIR(PUR) 가연성은 평균 7개 층(15층 건물에서 발화층 위 전체), PF 내부용은 2.6개 층이 번집니다. 나머지 준불연은 진압 전에 한 층 정도입니다.",
    "- 배리어(M1)를 넣으면 준불연은 0.3–0.5층으로 줄어듭니다. 가연성은 근거가 없으므로 효과를 주지 않았습니다.",
  ]);
}

// ================================================================================================ 12 step 8 risk calc
{
  const s = content("12단계 절차", "단계 8 · 위험도 산정 — R = Σ F·C를 한 재료로 손계산 (8.2.1.2 단계별 확인)",
    `PF 준불연 내부용(스프링클러 설치, 대책 없음): R_S = ${supStr(PT("PF_in", "B1", "M0").R_S)}층/동·년,\nF(n ≥ 2) = ${supStr(PT("PF_in", "B1", "M0").F_n2)}/동·년 → 예시 기준의 ${Math.round(PT("PF_in", "B1", "M0").F_n2 / best("crit_f2"))}배, ${pct(PT("PF_in", "B1", "M0").share_IE2)}가 외부 발화원(IE2)`);
  const r = PT("PF_in", "B1", "M0");
  const c1 = CO("PF_in", "IE1", false), cl = CO("NC", "IE1", false), c2 = CO("PF_in", "IE2", false);
  const pe = best("p_exp0");
  const t1 = r.F_IE1_ext * (pe * c1.E_storeys + (1 - pe) * cl.E_storeys), t2 = r.F_IE2_ext * pe * c2.E_storeys;
  const lw = 7.35;
  box(s, MX, 1.85, lw, 4.5, { fill: "F3F4F6", line: "F3F4F6", r: 0.08, name: "hand calculation" });
  const steps = [
    [runs({ text: "① F₁ ", options: { bold: true } }, "= λ₁ · U · (1 − p_s) · p_FO · p_v = ", sci(best("lam1"), 2), " × 60 × 0.10 × 0.15 × 0.70 = ", sci(r.F_IE1_ext, 3), " /년")],
    [runs({ text: "② F₂ ", options: { bold: true } }, "= λ₂ · p_x = ", sci(best("lam2"), 2), " × 0.30 = ", sci(r.F_IE2_ext, 3), " /년")],
    [runs({ text: "③ C ", options: { bold: true } }, `(노출 1회당 확산 층 수, 4_결과모델): IE1 외벽+도약 ${f2(c1.E_storeys, 3)} · 도약만 ${f2(cl.E_storeys, 3)} · IE2 외벽 ${f2(c2.E_storeys, 3)}`)],
    [runs({ text: "④ R_S ", options: { bold: true } }, "= F₁ [p_e C₁ + (1 − p_e) C_L] + F₂ p_e C₂")],
    [runs("      = ", sci(r.F_IE1_ext, 3), ` × (0.9 × ${f2(c1.E_storeys, 3)} + 0.1 × ${f2(cl.E_storeys, 3)}) + `, sci(r.F_IE2_ext, 2), ` × 0.9 × ${f2(c2.E_storeys, 3)}`)],
    [runs("      = ", sci(t1, 3), " + ", sci(t2, 3), " = ", { text: "", options: {} }, withOpt(sci(r.R_S, 3), { bold: true }), { text: " 층/동·년", options: { bold: true } })],
    [runs({ text: "⑤ F(n ≥ 2) ", options: { bold: true } }, "= ", sci(r.F_IE1_ext, 3), ` × 0.9 × ${f2(c1.P_ge2, 3)} + `, sci(r.F_IE2_ext, 2), ` × 0.9 × ${f2(c2.P_ge2, 3)} = `, withOpt(sci(r.F_n2, 3), { bold: true }), " /년")],
    [runs({ text: "⑥ 판정(점추정) ", options: { bold: true } }, withOpt(sci(r.F_n2, 2), {}), " > ", sci(best("crit_f2"), 1), " → 불충족 · IE2 비중 ", pct(r.share_IE2))],
  ];
  text(s, para(steps.map((x) => x[0]), { bullet: false, space: 6 }), { x: MX + 0.25, y: 1.98, w: lw - 0.45, h: 4.3, fontSize: 12.5 });
  const rx = MX + lw + 0.25, rw = CW - lw - 0.25;
  text(s, "전체 재료 · 점추정 (B1 · M0)", { x: rx, y: 1.85, w: rw, h: 0.3, fontSize: 13, bold: true });
  const rows = ORDER.map((m) => {
    const q = PT(m, "B1", "M0");
    return [TD(LABEL[m].replace("(기준안)", ""), { bold: true, fontSize: 10.5 }), TD(sci(q.R_S, 2), { align: "center", fontSize: 10.5 }),
      TD(sci(q.F_n2, 2), { align: "center", fontSize: 10.5 }), TD(f2(q.RR_S, 0), { align: "center", fontSize: 10.5 }), TD(pct(q.share_IE2), { align: "center", fontSize: 10.5 })];
  });
  table(s, [[TH("재료", { fontSize: 10.5 }), TH("R_S", { fontSize: 10.5 }), TH("F(n ≥ 2)", { fontSize: 10.5 }), TH("RR_S", { fontSize: 10.5 }), TH("IE2", { fontSize: 10.5 })], ...rows],
    { x: rx, y: 2.2, w: rw, colW: [1.7, 1.05, 1.05, 0.6, 0.583], fontSize: 10.5, border: BORDER, rowH: [0.34, ...ORDER.map(() => 0.42)], margin: [0.03, 0.05, 0.03, 0.05] });
  text(s, para([
    "RR_S = R_S ÷ 불연 기준안의 R_S (도약만, IE2 확산 없음)",
    "F(n ≥ 2) = 0인 재료도 MC에서는 0이 아님 → 15쪽",
  ], { space: 3 }), { x: rx, y: 5.5, w: rw, h: 0.85, fontSize: 10.5, color: C.text2 });
  NOTE(s, "같은 값이 계산서 5_위험도(수식)와 파이썬 point_estimate.csv에 있고, 7_검증 시트에서 두 값이 일치함을 확인합니다(17쪽).", 6.5);
  addNotes(s, [
    "[스크립트]",
    "- NFPA 551 8.2.1.2(각 계산 단계 확인)를 직접 보이는 쪽입니다. 사건수목에서 F를, 결과 모델에서 C를 가져와 R = Σ F·C를 손으로 계산했습니다.",
    "- PF 내부용은 점추정으로 R_S 1.8e-3층/년, 2개 층 이상 확산 빈도 6.4e-4/년으로 예시 기준 1e-4를 넘습니다. 기여의 83 %가 IE2(외부 발화원)입니다.",
    "- 오른쪽 표에서 PIR(PUR) 준불연·PF 외부용·EPS 준불연은 점추정 F(n≥2)가 0입니다. 진압 시각을 중앙값 15분 하나로 놓았기 때문이며, 분포를 넣으면 0이 아닙니다(15쪽).",
  ]);
}

// ================================================================================================ 13 results R_S
pres.addSection({ title: "예제 결과" });
{
  const s = content("예제 결과", "예제 결과 ① — 외벽 확산 위험도 R_S와 90 % 구간 (2중 몬테카를로, 스프링클러 설치)",
    "순위: PIR(PUR) 가연성 ≫ PF 내부용 > PF 외부용 ≈ PIR(PUR) 준불연 > EPS 준불연\n— 차단재 + 배리어(M1)로 준불연 R_S는 1/9–1/19, 가연성은 효과 불인정");
  const ch = { x: MX, y: 1.8, w: 8.35, h: 4.6, L: { x: 0.11, y: 0.04, w: 0.87, h: 0.8 }, xMin: -0.5, xMax: ORDER.length - 0.5, xMajor: 1, yMin: 1e-6, yMax: 1e-1, yLog: true };
  const xs = [], series = [];
  const npts = ORDER.length * 2 * 3;
  ORDER.forEach((m, i) => {
    ["M0", "M1"].forEach((g, k) => {
      const r = MC(m, "B1", g);
      const x = i + (k === 0 ? -0.17 : 0.17), base = xs.length;
      xs.push(x, x, x);
      const ys = (pairs) => { const a = new Array(npts).fill(null); pairs.forEach(([j, v]) => (a[base + j] = v)); return a; };
      series.push({ name: `${m} ${g} 90%`, ys: ys([[0, r.R_S_p05], [1, r.R_S_p95]]), color: MAT[m], style: k === 0 ? { noMarker: true } : { noMarker: true, dash: "sysDash" } });
      series.push({ name: `${m} ${g} 평균`, ys: ys([[2, r.R_S_mean]]), color: MAT[m],
        style: { noLine: true, marker: k === 0 ? { symbol: "circle", size: 9, fill: MAT[m], line: MAT[m] } : { symbol: "square", size: 8, fill: tint(MAT[m], 0.55), line: MAT[m], lineW: 22225 } } });
    });
  });
  curves(s, ch, xs, series, { lineSize: 2.25, valAxisLogScaleBase: 10, valAxisLabelFormatCode: "0E+0", catAxisHidden: true, valAxisMajorUnit: undefined });
  const mp = plotMap(ch);
  ORDER.forEach((m, i) => text(s, SHORT[m], { x: mp.X(i) - 0.6, y: mp.py + mp.ph + 0.06, w: 1.2, h: 0.5, fontSize: 10.5, align: "center", color: HEX.ink }));
  vlabel(s, "R_S (층/동·년, 로그; 1E-3 = 10⁻³)", ch.x + 0.17, mp.py + mp.ph / 2, 3.2);
  const ly = 6.5;
  s.addShape(pres.ShapeType.ellipse, { x: ch.x + 0.95, y: ly + 0.06, w: 0.16, h: 0.16, fill: { color: HEX.ink2 }, line: { color: HEX.ink2, width: 0.5 } });
  text(s, "M0 대책 없음 (평균, 세로선 = 90 % 구간)", { x: ch.x + 1.18, y: ly, w: 3.4, h: 0.28, fontSize: 11, color: C.text2, valign: "middle" });
  s.addShape(pres.ShapeType.rect, { x: ch.x + 4.75, y: ly + 0.065, w: 0.15, h: 0.15, fill: { color: tint(HEX.ink2, 0.55) }, line: { color: HEX.ink2, width: 1.5 } });
  text(s, "M1 차단재 + 배리어 (점선 = 90 % 구간)", { x: ch.x + 4.98, y: ly, w: 3.3, h: 0.28, fontSize: 11, color: C.text2, valign: "middle" });
  const rx = MX + 8.6, rw = CW - 8.6;
  const rows = ORDER.map((m) => {
    const a = MC(m, "B1", "M0"), b = MC(m, "B1", "M1");
    return [TD(LABEL[m].replace(" (기준안)", "").replace("(기준안)", ""), { bold: true, fontSize: 10 }), TD(sci(a.R_S_mean, 2), { align: "center", fontSize: 10 }),
      TD(sci(b.R_S_mean, 2), { align: "center", fontSize: 10 }), TD(m === "NC" ? "—" : `÷${(a.R_S_mean / b.R_S_mean).toFixed(1)}`, { align: "center", fontSize: 10 })];
  });
  table(s, [[TH("재료", { fontSize: 10 }), TH("M0 평균", { fontSize: 10 }), TH("M1 평균", { fontSize: 10 }), TH("감소", { fontSize: 10 })], ...rows],
    { x: rx, y: 1.85, w: rw, colW: [1.38, 0.86, 0.86, 0.633], fontSize: 10, border: BORDER, rowH: [0.32, ...ORDER.map(() => 0.38)], margin: [0.03, 0.04, 0.03, 0.04] });
  text(s, para([
    "PF 외부용은 자립 연소(B 2/2) 때문에 진압이 늦은 사건에서 계속 올라가 평균이 PIR(PUR) 준불연보다 큼 — 점추정(중앙값)에서는 둘이 같았음",
    "90 % 구간 폭이 10–30배 → 판정은 평균 하나가 아니라 ‘기준을 만족하는 비율’로 함(15쪽)",
  ], { space: 5 }), { x: rx, y: 4.6, w: rw, h: 1.8, fontSize: 11 });
  addNotes(s, [
    "[스크립트]",
    "- 점이 평균, 세로선이 인식론적 90 % 구간입니다. 왼쪽 진한 원이 대책 없음(M0), 오른쪽 빈 네모가 차단재 + 배리어(M1)입니다.",
    "- 순위는 기존 R = P×S와 같은 방향입니다. 다른 점은 PF 외부용: 자립 연소 플래그 B가 '위험도'에 숫자로 들어가, 진압이 늦은 사건에서 계속 올라가는 몫이 반영됩니다.",
    "- 구간이 넓은 것은 입력이 예시값이고 범위를 넓게 잡았기 때문입니다. 그래서 판정은 신뢰수준으로 합니다.",
  ]);
}

// ================================================================================================ 14 exceedance
{
  const s = content("예제 결과", "예제 결과 ② — 초과빈도 곡선 F(발화층 위 확산 ≥ n) (MC 평균, 스프링클러 설치)",
    "n ≥ 2에서 PIR(PUR) 가연성·PF 내부용은 예시 기준의 6–9배, 다른 준불연은 기준 근처\n— 차단재 + 배리어(M1)에서 준불연은 0.8–2.9×10⁻⁵/동·년으로 기준 아래");
  const ns = EX("NC", "B1", "M0").map((r) => r.n);
  const pan = [["M0", "M0 · 대책 없음"], ["M1", "M1 · 개구부 차단재 + 층간 배리어"]];
  const pw = (CW - 0.35) / 2;
  pan.forEach(([g, title], k) => {
    const ch = { x: MX + k * (pw + 0.35), y: 1.85, w: pw, h: 4.3, L: { x: 0.15, y: 0.1, w: 0.81, h: 0.74 }, xMin: 1, xMax: ns[ns.length - 1], xMajor: 1, yMin: 1e-8, yMax: 1e-2, yLog: true, title };
    const series = ORDER.map((m) => ({
      name: `${g} ${m}`, ys: EX(m, "B1", g).map((r) => (r.F_mean >= 1e-8 ? r.F_mean : null)), color: MAT[m], style: m === "NC" ? { dash: "dash" } : undefined,
    }));
    series.push({ name: `${g} 기준`, ys: ns.map((n) => (n === 2 ? best("crit_f2") : null)), color: HEX.ink, style: { noLine: true, marker: { symbol: "diamond", size: 10, fill: HEX.ink, line: HEX.ink } } });
    curves(s, ch, ns, series, { lineSize: 2, valAxisLogScaleBase: 10, valAxisLabelFormatCode: "0E+0", catAxisLabelFormatCode: "0", valAxisMajorUnit: undefined, titleFontSize: 12 });
    const mp = plotMap(ch);
    text(s, "발화층 위로 번진 층 수 n", { x: mp.px, y: mp.py + mp.ph + 0.28, w: mp.pw, h: 0.28, fontSize: 11, color: C.text2, align: "center" });
    vlabel(s, "F (/동·년, 로그)", ch.x + 0.17, mp.py + mp.ph / 2, 2.0);
    if (k === 1) text(s, "PIR(PUR) 가연성: M1 = M0\n(대책 효과 불인정)", { x: mp.px + mp.pw - 2.3, y: mp.Y(4e-4) - 0.62, w: 2.25, h: 0.5, fontSize: 10, color: C.text2, align: "right" });
  });
  const LEG = { NC: "불연 기준안", EPS_SC: "EPS 준불연", PIR_C: "PIR 가연성", PIR_SC: "PIR 준불연", PF_out: "PF 외부용", PF_in: "PF 내부용" };
  legend(s, [...ORDER.map((m) => ({ label: LEG[m], color: MAT[m], w: 1.08, dash: m === "NC" })), { kind: "dot", color: HEX.ink, label: "예시 기준 (n = 2)", w: 1.35 }], MX, 6.3);
  NOTE(s, "곡선은 인식론적 평균. 90 % 구간은 계산서 6_불확실성(MC)·figures/fig1_exceedance.png. 1×10⁻⁸ 아래는 그리지 않음.", 6.65);
  addNotes(s, [
    "[스크립트]",
    "- 가로축은 발화층 위로 번진 층 수 n, 세로축은 그 이상 번지는 사건의 연간 빈도입니다. 정량법 출력의 'F–N 곡선'(표 5.1.2.1의 (4))을 외벽 확산 범위로 그린 것입니다.",
    "- 왼쪽(대책 없음): 가연성 PIR은 n과 상관없이 거의 수평입니다. 한 번 들어가면 건물 끝까지 갑니다. PF 내부용도 천천히 떨어집니다.",
    "- 오른쪽(대책 있음): 준불연은 배리어가 층마다 막아 곡선이 급하게 떨어집니다.",
  ]);
}

// ================================================================================================ 15 decision table
{
  const s = content("예제 결과", "예제 결과 ③ — 허용기준 판정: 점추정과 불확실성을 넣은 판정",
    "대책 없음: PIR(PUR) 가연성·PF 내부용 ‘불충족’, 준불연 3종 ‘판단 유보’(충족 41–55 %)\n— 차단재 + 배리어로 준불연 4종 ‘충족’(96–100 %), 가연성은 ‘불충족’");
  const head = [TH("재료"), TH("대책"), TH("F(n ≥ 2) 점추정"), TH("F(n ≥ 2) MC 평균 [5 – 95 %]"), TH("기준 충족 비율"), TH("판정 (신뢰 90 %)")];
  const rows = [];
  ORDER.forEach((m) => ["M0", "M1"].forEach((g) => {
    const q = PT(m, "B1", g), r = MC(m, "B1", g);
    rows.push({ m, g, cells: [TD(g === "M0" ? LABEL[m] : "", { bold: true }), TD(g, { align: "center" }), TD(sci(q.F_n2, 2), { align: "center" }),
      TD(runs(sci(r.F_n2_mean, 2), "  [", sci(r.F_n2_p05, 1), " – ", sci(r.F_n2_p95, 1), "]"), { align: "center" }),
      TD(pct(r.conf_meets_F2, 1), { align: "center", bold: true }), TD("", {})], verdict: r.verdict });
  }));
  const y0 = 1.8, rh = 0.33, hh = 0.34;
  table(s, [head, ...rows.map((r) => r.cells)], { x: MX, y: y0, w: CW, colW: [2.4, 0.8, 1.75, 3.6, 1.65, 2.133], fontSize: 10.5, border: BORDER, rowH: [hh, ...rows.map(() => rh)], margin: [0.02, 0.08, 0.02, 0.08] });
  rows.forEach((r, i) => {
    verdictChip(s, MX + 2.4 + 0.8 + 1.75 + 3.6 + 1.65 + 0.3, y0 + hh + i * rh + 0.015, 1.5, r.verdict);
    if (r.g === "M0") s.addShape(pres.ShapeType.rect, { x: MX - 0.14, y: y0 + hh + i * rh + 0.05, w: 0.08, h: rh * 2 - 0.1, fill: { color: MAT[r.m] }, line: { color: MAT[r.m], width: 0.5 } });
  });
  text(s, para([
    runs({ text: "판정 규칙(예시): ", options: { bold: true } }, "인식론적 표본 2,000개 중 F(n ≥ 2) ≤ 1×10⁻⁴인 비율이 90 % 이상 → 충족, 10 % 이하 → 불충족, 그 사이 → 판단 유보(자료를 더 모아야 결론이 나는 상태)"),
    runs({ text: "점추정의 함정: ", options: { bold: true } }, "PIR(PUR) 준불연·PF 외부용·EPS 준불연은 점추정 0인데 MC 평균 1.6–2.5×10⁻⁴ — 진압이 늦어지는 꼬리 사건을 중앙값 하나로는 볼 수 없음(A.5.5.6)"),
  ], { space: 4 }), { x: MX, y: y0 + hh + rows.length * rh + 0.12, w: CW, h: 0.8, fontSize: 11 });
  addNotes(s, [
    "[스크립트]",
    "- 같은 기준을 점추정과 몬테카를로로 판정한 표입니다. 점추정으로는 준불연 3종이 '충족'처럼 보이지만, 분포를 넣으면 기준을 만족하는 비율이 41–55 %로 '판단 유보'입니다.",
    "- 판단 유보는 실패가 아니라 '이 입력 수준으로는 결론을 못 낸다'는 뜻이고, 무엇을 측정해야 결론이 나는지는 다음 쪽 민감도가 알려 줍니다.",
    "",
    "[질문 대비]",
    "- 90 %라는 신뢰수준과 1e-4라는 기준은 모두 예시입니다. 관할기관과 정할 값입니다(4.4.3.4).",
  ]);
}

// ================================================================================================ 16 uncertainty
pres.addSection({ title: "불확실성 · 검증" });
{
  const s = content("불확실성 · 검증", "단계 9 · 불확실성과 민감도 — 우연적·인식론적 분리 (4.5, 5.5.6, A.5.5.6)",
    "결론을 가장 크게 흔드는 입력은 s(3 m 시험 → 실물)와 외부 발화원 빈도 λ₂\n— 대책(M1)을 넣으면 배리어 효과 p_cb와 차단재 효과 p_e1이 커짐");
  const lw = 4.55;
  const kinds = [
    ["인식론적 (바깥 루프 2,000)", "더 알면 줄어드는 불확실성", "λ₁, λ₂, p_s, p_FO, p_v, p_x, p_e0, p_e1, p_cb, p_L, s, T_FS·RSET·τ_L의 중앙값, p_B ~ Beta(k + ½, n − k + ½)", "conditional"],
    ["우연적 (안쪽 루프 400, LHS)", "화재마다 달라지는 변동", "T_FS·RSET·τ_L (로그정규), 분출 지속 t_e (균등), 어떤 반복시험처럼 거동하는지, 자립 연소 B", "good"],
  ];
  kinds.forEach(([h, sub, body, st], i) => {
    const y = 1.85 + i * 1.75;
    box(s, MX, y, lw, 1.6, { fill: STATUS[st].tint, line: STATUS[st].tint, r: 0.08, name: h });
    text(s, h, { x: MX + 0.2, y: y + 0.08, w: lw - 0.4, h: 0.32, fontSize: 13.5, bold: true });
    text(s, sub, { x: MX + 0.2, y: y + 0.4, w: lw - 0.4, h: 0.26, fontSize: 11, color: C.text2 });
    text(s, body, { x: MX + 0.2, y: y + 0.68, w: lw - 0.4, h: 0.88, fontSize: 11 });
  });
  text(s, para([
    "위험도 = 우연적 변동에 대한 기대값(안쪽), 그 위험도의 분포 = 인식론적 불확실성(바깥)",
    "모든 재료가 같은 표본을 써서(공통 난수) 재료 간 비교가 정밀함",
  ], { space: 3 }), { x: MX, y: 5.4, w: lw, h: 0.9, fontSize: 11, color: C.text2 });
  // heat table
  const keys = ["s", "lam2", "p_x2", "t_fs", "p_cb", "p_exp1", "lam1", "p_fo", "p_spr", "p_B"];
  const sym = { s: "s", lam2: "λ₂", p_x2: "p_x", t_fs: "T_FS", p_cb: "p_cb", p_exp1: "p_e1", lam1: "λ₁", p_fo: "p_FO", p_spr: "p_s", p_B: "p_B" };
  const rx = MX + lw + 0.3, rw = CW - lw - 0.3;
  const short = { EPS_SC: "EPS\n준불연", PIR_C: "PIR\n가연성", PIR_SC: "PIR\n준불연", PF_out: "PF\n외부용", PF_in: "PF\n내부용" };
  const head1 = [TH("입력", { rowspan: 2 }), TH("M0 대책 없음", { colspan: 5 }), TH("M1 차단재 + 배리어", { colspan: 5 })];
  const head2 = [...MATS.map((m) => TH(short[m], { fontSize: 9.5 })), ...MATS.map((m) => TH(short[m], { fontSize: 9.5 }))];
  const body = keys.map((k) => [TD(sym[k], { bold: true, align: "center" }), ...["M0", "M1"].flatMap((g) => MATS.map((m) => {
    const v = SE(g, m, k).rho_R_S;
    return TD((v >= 0 ? "+" : "−") + Math.abs(v).toFixed(2), { align: "center", fontSize: 10, fill: { color: heat(v) }, color: Math.abs(v) > 0.55 ? "FFFFFF" : HEX.ink });
  }))]);
  table(s, [head1, head2, ...body], { x: rx, y: 1.85, w: rw, colW: [0.88, ...new Array(10).fill((rw - 0.88) / 10)], fontSize: 10, border: BORDER,
    rowH: [0.3, 0.46, ...keys.map(() => 0.3)], margin: [0.02, 0.03, 0.02, 0.03] });
  text(s, "값: log R_S와의 순위상관 ρ (주황 +: 클수록 위험 증가, 청록 −: 클수록 감소)", { x: rx, y: 5.68, w: rw, h: 0.26, fontSize: 10, color: C.text2 });
  box(s, rx, 5.98, rw, 0.88, { fill: "FFFFFF", line: HEX.ink, lw: 1, r: 0.08, name: "priorities" });
  text(s, para([
    [["자료 보강 우선순위  ", { bold: true }], ["① s → 실대형 시험(KS F 8414) 2–3종  ② λ₂·p_x → 국가화재정보시스템 발화 장소 통계", {}]],
    [["", {}], ["③ p_cb·p_e1 → 배리어·차단재를 넣은 중공층 시험  ④ T_FS → 관할 소방서 출동·전개 시간", {}]],
  ], { bullet: false, space: 2 }), { x: rx + 0.15, y: 6.02, w: rw - 0.3, h: 0.8, fontSize: 11, valign: "middle" });
  addNotes(s, [
    "[스크립트]",
    "- NFPA 551 A.5.5.6은 우연적(무작위) 불확실성과 인식론적(지식 부족) 불확실성을 구분하라고 합니다. 2중 몬테카를로가 그 구분을 그대로 구현합니다.",
    "- 오른쪽 표는 어떤 입력을 알게 되면 결론이 가장 많이 바뀌는지를 보여 줍니다. 대책이 없을 때는 s와 λ2, 대책이 있을 때는 p_cb와 p_e1이 큽니다.",
    "- 그래서 다음 시험은 '배리어를 넣은 중공층 시험'과 '실대형 시험'이 가장 값어치가 큽니다.",
  ]);
}

// ================================================================================================ 17 V&V
{
  const s = content("불확실성 · 검증", "단계 10 · 검증과 확인 (5.5.9, 8.2) — 계산이 맞는가, 모델이 현실을 대표하는가",
    `검증: 엑셀 수식과 파이썬이 독립적으로 같은 값 (최대 상대차 ${supStr(VER.max_rel_diff)})\n확인: ‘3 m 시험 = 1개 층’ 외삽은 미확인 → 실대형 시험·PyroSim 비교가 다음 단계`);
  const cw = (CW - 0.25) / 2;
  box(s, MX, 1.85, cw, 4.55, { fill: STATUS.good.tint, line: STATUS.good.tint, r: 0.08, name: "verification" });
  text(s, "검증 (Verification, 8.2.1) — 완료", { x: MX + 0.2, y: 1.95, w: cw - 0.4, h: 0.36, fontSize: 14.5, bold: true });
  text(s, para([
    [["8.2.1.1 대체 계산 재현  ", { bold: true }], [`엑셀 수식 ${VER.n_formulas.toLocaleString()}개 vs 파이썬 — ${VER.n_cases}경우 × 5지표 ${VER.n_values}개 값, 최대 상대차 `, {}], ...sci(VER.max_rel_diff, 2).map((r) => [r.text, r.options])],
    [["8.2.1.2 단계별 확인  ", { bold: true }], ["PF 내부용 손계산(12쪽) = 계산서 5_위험도 = point_estimate.csv", {}]],
    [["몬테카를로 내부 확인  ", { bold: true }], ["우연적 변동을 0으로 두면 MC 결과가 점추정과 같음(상대차 0, qfra.py의 assert)", {}]],
    [["표본 수 확인  ", { bold: true }], ["안쪽 표본 1,000개(무작위)와 400개(LHS)의 R_S 평균 차이 1.5 % 이내", {}]],
  ], { space: 7 }), { x: MX + 0.2, y: 2.4, w: cw - 0.4, h: 3.9, fontSize: 12 });
  const rx = MX + cw + 0.25;
  box(s, rx, 1.85, cw, 4.55, { fill: STATUS.fair.tint, line: STATUS.fair.tint, r: 0.08, name: "validation" });
  text(s, "확인 (Validation, 8.2.2) — 계획", { x: rx + 0.2, y: 1.95, w: cw - 0.4, h: 0.36, fontSize: 14.5, bold: true });
  text(s, para([
    [["시험과 비교 (8.2.2.2)  ", { bold: true }], ["재료 입력이 시험값 자체라 3 m 안에서는 일치. ‘한 층씩 τ_F로 오른다’는 외삽은 2개 층 이상 실대형 시험(KS F 8414)으로 확인", {}]],
    [["대체 모델 비교 (8.2.2.1)  ", { bold: true }], ["PyroSim(FDS)으로 분출화염 → 중공층 확산을 계산해 τ_F와 s 비교(면재·마감층 포함 모델, 앞선 PyroSim 자료)", {}]],
    [["TGA 예측식의 범위  ", { bold: true }], ["미시험 준불연 제품의 t₆₀₀ 추정에 쓸 수 있음(교차검증 P등급 6/6, 오차 −14 – +29 %); 가연성은 불가(예측 777 s → 실측 30 s)", {}]],
    [["실제 화재와 비교  ", { bold: true }], ["국내 외벽 화재의 층간 확산 시간·층 수와 정성 비교", {}]],
  ], { space: 7 }), { x: rx + 0.2, y: 2.4, w: cw - 0.4, h: 3.9, fontSize: 12 });
  NOTE(s, "NFPA 551 5.5.9: 정량법은 ‘제대로 구성되면 실제 위험을 정확히 대표’한다고 받아들여지지만, 결과·빈도 추정에 쓴 모델의 확인 근거는 함께 제출해야 함.", 6.5);
  addNotes(s, [
    "[스크립트]",
    "- 검증은 '식을 제대로 풀었는가', 확인은 '식이 현실을 대표하는가'입니다(3.3.18–3.3.19).",
    "- 검증은 끝났습니다: 엑셀과 파이썬을 따로 짜서 120개 값이 수치 오차 수준에서 같습니다. 검토자가 엑셀만 열어도 같은 숫자가 나옵니다.",
    "- 확인은 남았습니다: 3 m 시험을 한 개 층으로 보고 층마다 같은 시간으로 오른다고 본 가정이 가장 약한 고리이며, 실대형 시험이나 PyroSim 계산으로 확인해야 합니다.",
  ]);
}

// ================================================================================================ 18 documentation
pres.addSection({ title: "문서화 · 검토" });
{
  const s = content("문서화 · 검토", "단계 11 · 문서화 — 무엇을 어떻게 서술하나 (7.2 개념보고서, 7.3.3.4 정량 결과 문서)",
    "개념보고서는 ‘분석 전에 합의할 것’ 8항목, 정량 결과 문서는 ‘분석에서 나온 것’ 8항목 —\n서술양식.docx에 항목마다 쓸 내용 · 작성 방법 · 예문을 넣어 둠");
  const left = [
    ["7.2.2.1 참여자", "평가자·검토자의 학력, FRA 경력, 자격(1.4)"],
    ["7.2.2.2 범위·목적", "목적, 건물·외벽 시스템, 사용 단계, 외벽 경로만"],
    ["7.2.2.3 목표", "발화층 위 재실자 보호, 확산 층 수 제한"],
    ["7.2.2.4 허용기준", "F(n ≥ 2) ≤ 기준값, 신뢰수준, 근거"],
    ["7.2.2.5 위험요소", "노출 대상 · 자극 · 전달 · 반응(4.4.2)"],
    ["7.2.2.6 시나리오", "IE1 · IE2, 군집 근거, 제외와 이유"],
    ["7.2.2.7 방법", "사건수목 + 층간 확산 모델 + 2중 MC, 적합성"],
    ["7.2.2.8 자료 출처", "통계·시험·가정값, 예시값 표시"],
  ];
  const right = [
    ["7.3.3.4.1 빈도·확률", "경로별 F와 기간(/동·년), 사건수목"],
    ["7.3.3.4.2 결과", "군집별 C와 대표성(반복시험 가중)"],
    ["7.3.3.4.3 위험도", "R = Σ F·C, 시나리오 대표성"],
    ["7.3.3.4.4 불확실성", "구분·전파·민감도·신뢰수준"],
    ["7.3.3.4.5 SW·모델 평가", "버전, 검증 결과, 적용 범위"],
    ["7.3.3.4.6 한계", "3 m → 1개 층 외삽, 예시 입력 등"],
    ["7.3.3.4.7 결론", "기준 대비 판정, 목적 달성 정도"],
    ["7.3.3.4.8 참고문헌", "도면·시험성적서·통계·기준(판/연도)"],
  ];
  const cw = (CW - 0.3) / 2;
  [[left, "개념보고서 (7.2) — 분석 전, 관할기관과 합의", "conditional"], [right, "정량 결과 문서 (7.3.3.4) — 분석 후 제출", "good"]].forEach(([rows, title, st], k) => {
    const x = MX + k * (cw + 0.3);
    box(s, x, 1.85, cw, 0.42, { fill: STATUS[st].tint, line: STATUS[st].tint, r: 0.06 });
    text(s, title, { x: x + 0.15, y: 1.85, w: cw - 0.3, h: 0.42, fontSize: 13, bold: true, valign: "middle" });
    table(s, rows.map(([a, b]) => [TD(a, { bold: true, fontSize: 11 }), TD(b, { fontSize: 11 })]),
      { x, y: 2.35, w: cw, colW: [2.15, cw - 2.15], fontSize: 11, border: BORDER, rowH: rows.map(() => 0.44), margin: [0.03, 0.08, 0.03, 0.08] });
  });
  box(s, MX, 6.0, CW, 0.85, { fill: "F3F4F6", line: "F3F4F6", r: 0.08, name: "example sentence" });
  text(s, para([
    [["예문 (7.3.3.4.3)  ", { bold: true }], ["“PF 준불연 내부용 외단열을 적용한 15층 공동주택의 외벽 확산 위험도는 1.8×10⁻³층/동·년(점추정)이며, 이 중 83 %는 외부 발화원(IE2)에서 기인한다. 사건수목의 두 초기사건은 외벽을 화염에 노출시키는 모든 사용 단계 시나리오를 포함하도록 선정하였다(5.5.1.1).”", {}]],
  ], { bullet: false }), { x: MX + 0.2, y: 6.03, w: CW - 0.4, h: 0.8, fontSize: 11, valign: "middle" });
  addNotes(s, [
    "[스크립트]",
    "- '서술 방법이 없다'에 대한 직접 답입니다. NFPA 551 7장이 무엇을 써야 하는지는 항목으로 정해 두었고, 이를 외벽 정량 평가에 맞게 풀었습니다.",
    "- 왼쪽은 분석 전에 관할기관과 합의하는 개념보고서, 오른쪽은 분석 결과를 담는 정량 결과 문서입니다. 항목별 예문은 서술양식.docx에 있습니다.",
  ]);
}

// ================================================================================================ 19 O&M
{
  const s = content("문서화 · 검토", "단계 11 · 운영 · 유지관리 매뉴얼 (7.4)과 이행 (7.5) — 평가가 유효하려면 지켜야 할 조건",
    "외벽 평가는 평가 때의 재료·차단재·배리어·발화원 관리가 그대로일 때만 유효\n— 지킬 조건, 점검, 변경 관리, 재평가 계기를 매뉴얼에 적음");
  const items = [
    ["7.4.2 한계·가정 목록", "conditional", ["외장재·단열재가 평가한 제품(시험 시편과 같은 구성)일 것", "개구부 차단재·층간 배리어가 설계대로 시공되고 관통부가 없을 것", "필로티·외벽 하부에 차량·쓰레기 등 가연물을 두지 않을 것(IE2 빈도 가정)", "세대 스프링클러가 작동 상태일 것(IE1 분기)"]],
    ["7.4.3 변경 관리", "fair", ["외장 교체·리모델링·창호 교체", "설비 관통부 추가, 발코니 확장·용도 변경", "→ 위 변경 전 영향 평가, 필요하면 재평가", "변경 이력을 평가 문서에 기록"]],
    ["7.4.4 점검·유지관리", "good", ["배리어·차단재: 관통부·시공 하자 정기 점검", "스프링클러: 작동 시험(신뢰도 p_s 근거)", "필로티 주차·적치물 점검", "점검 결과로 p_s·p_e·p_cb 가정을 갱신"]],
    ["7.5 이행", "poor", ["점검 기록 제출 주기와 담당", "미이행 시 조치(관할 규정 연계)", "사후 발견(제품 리콜, 시공 하자) 시 재평가", "재평가 계기를 계약·관리 규약에 반영"]],
  ];
  const cw = (CW - 3 * 0.2) / 4;
  items.forEach(([h, st, list], i) => card(s, MX + i * (cw + 0.2), 1.85, cw, 4.45, st, h, list, { size: 11.5, headSize: 13.5 }));
  NOTE(s, "NFPA 551 8.3 검토 질문 (7)–(9): 운영·유지관리 매뉴얼, 변경 관리 지침, 점검·시험·유지관리 프로그램이 있는가.", 6.5);
  addNotes(s, [
    "[스크립트]",
    "- 외벽 평가 결과는 조건부입니다. 배리어가 뚫리거나 필로티에 차량을 세우면 가정한 확률이 바뀝니다.",
    "- 그래서 7.4는 어떤 조건을 지켜야 하고, 무엇이 바뀌면 다시 평가하는지를 매뉴얼로 남기라고 합니다. 검토 질문 9개 중 3개가 여기에 해당합니다.",
  ]);
}

// ================================================================================================ 20 review
{
  const s = content("문서화 · 검토", "단계 12 · 검토와 승인 (8.1, 8.3) — 검토 질문 9개와 답하는 위치",
    "관할기관은 직접 검토하거나 제3자(동료 검토·계약 검토)에 맡김\n— 검토 질문 9개의 답 위치를 미리 표로 붙이면 검토가 빨라짐");
  const lw = 3.6;
  box(s, MX, 1.85, lw, 4.5, { fill: "F3F4F6", line: "F3F4F6", r: 0.08, name: "review approaches" });
  text(s, "검토 방식 (8.1)", { x: MX + 0.2, y: 1.95, w: lw - 0.4, h: 0.34, fontSize: 14, bold: true });
  text(s, para([
    [["직접 검토  ", { bold: true }], ["관할기관이 자원이 있을 때 문서를 직접 검토(8.1.1)", {}]],
    [["동료 검토  ", { bold: true }], ["제3자가 의견서 제출, 판단은 관할기관(8.1.2.1). 검토자는 평가에 관여하지 않은 같은 수준의 전문가", {}]],
    [["계약 검토  ", { bold: true }], ["제3자가 승인·보완·반려까지 결정(8.1.2.2)", {}]],
    [["검토 기법 (8.2)  ", { bold: true }], ["검증: 대체 계산·단계 확인·표본 감사 / 확인: 대체 모델·시험 비교·실증", {}]],
  ], { space: 6 }), { x: MX + 0.2, y: 2.35, w: lw - 0.4, h: 3.95, fontSize: 11.5 });
  const qs = [
    ["(1)", "평가 목적이 명시되었는가", "서술양식 1.2 · 이 덱 7쪽"],
    ["(2)", "평가 범위가 명시되었는가", "서술양식 1.2"],
    ["(3)", "방법과 적합성 근거가 있는가", "서술양식 1.7 · 이 덱 3·10쪽"],
    ["(4)", "한계가 명시되었는가", "서술양식 2.6 · 이 덱 21쪽"],
    ["(5)", "분석 결과가 포함되었는가", "계산서 5_위험도 · 6_불확실성(MC)"],
    ["(6)", "결론이 포함되었는가", "서술양식 2.7 · 이 덱 15쪽"],
    ["(7)", "운영·유지관리 매뉴얼이 있는가", "서술양식 3.1 · 이 덱 19쪽"],
    ["(8)", "변경 관리 지침이 있는가", "서술양식 3.2"],
    ["(9)", "점검·시험·유지관리 프로그램이 있는가", "서술양식 3.3"],
  ];
  const rx = MX + lw + 0.25, rw = CW - lw - 0.25;
  table(s, [[TH("8.3"), TH("검토 질문"), TH("답하는 위치")], ...qs.map(([a, b, c]) => [TD(a, { align: "center", bold: true }), TD(b), TD(c, { fill: { color: "F3F7F2" } })])],
    { x: rx, y: 1.85, w: rw, colW: [0.65, 4.0, rw - 4.65], fontSize: 11.5, border: BORDER, rowH: [0.36, ...qs.map(() => 0.44)], margin: [0.03, 0.08, 0.03, 0.08] });
  NOTE(s, "계산서 8_문서화점검 시트에 7.2 · 7.3.3.4 · 7.4 · 8.3 항목과 상태(완료/작성 중/미작성/해당 없음)를 기록할 수 있습니다.", 6.5);
  addNotes(s, [
    "[스크립트]",
    "- 8.3의 질문 9개는 검토자가 반드시 묻는 것들입니다. 각 질문의 답이 문서 어디에 있는지 표로 붙이면 검토 기간이 줄어듭니다.",
    "- 동료 검토자는 평가에 관여하지 않은 사람이어야 합니다(8.1.2.1).",
  ]);
}

// ================================================================================================ 21 limitations
{
  const s = content("문서화 · 검토", "한계 (7.3.3.4.6)와 다음 단계",
    "가장 약한 고리는 ‘3 m 중규모 시험 = 1개 층, 층마다 같은 시간’이라는 외삽과 예시 입력값 —\n실대형 시험 · 화재 통계 · 배리어 시험이 들어오면 같은 절차로 결론이 확정됨");
  const cw = (CW - 0.25) / 2;
  card(s, MX, 1.85, cw, 4.5, "poor", "한계", [
    "3 m 중공층 시험을 ‘1개 층 확산’으로 외삽, 층마다 같은 τ_F (s = 0.25–1 가정)",
    "빈도·분기확률·시간은 예시값 — 국내 통계 미반영",
    "분기 독립 가정, 수평 확산·바람·낙하물 미고려",
    "노출 = 피난 완료 전 화염 도달(연기 영향 제외), 사망 환산 p_d는 가정",
    "반복시험 2–3회 → t₆₀₀·p_B 표본이 작음, EPS 가연성은 자료 없음",
  ], { size: 12 });
  card(s, MX + cw + 0.25, 1.85, cw, 4.5, "good", "다음 단계", [
    "EFFECT 사용자 가이드 원문 → 문항·가중치 대응표와 입력 갱신",
    "국가화재정보시스템: 공동주택 화재 · 발화 장소(필로티·주차장) 빈도 → λ₁, λ₂",
    "실대형 시험(KS F 8414) 2–3종 → s 보정(가장 큰 불확실성)",
    "배리어·차단재를 넣은 중공층 시험 → p_cb, p_e1 (배리어 재착화 시간 t_b)",
    "EPS 가연성 재시험(30분 기록) · 필요하면 비용편익(5.6)으로 대책 비교",
  ], { size: 12 });
  NOTE(s, "입력을 바꾸면: 계산서(xlsx) 노란 칸 → 점추정 즉시 갱신 / qfra.py의 PARAMS → MC·민감도·그림 갱신.", 6.5);
  addNotes(s, [
    "[스크립트]",
    "- 한계는 숨기지 않고 문서에 씁니다(7.3.3.4.6). 이 예제의 결론은 입력값에 따라 바뀔 수 있습니다.",
    "- 다음 단계는 민감도 순서와 같습니다: 실대형 시험으로 s를, 화재 통계로 외부 발화원 빈도를, 배리어 시험으로 대책 효과를 확정합니다.",
  ]);
}

// ================================================================================================ 22 appendix params
pres.addSection({ title: "부록" });
{
  const s = content("부록", "부록 A — 기호와 입력값 (예시값과 대체할 자료)",
    "측정값은 t₆₀₀·플래그 B뿐, 나머지는 예시값 — 대체할 자료를 함께 적어 둠 (6.2.1.3)");
  const keys = ["lam1", "lam2", "p_spr", "p_fo", "p_vent", "p_x2", "p_exp0", "p_exp1", "p_cb", "p_L", "tau_L", "t_fs", "rset", "s", "p_d", "crit_f2"];
  const val = (v, k) => (Math.abs(v) < 0.01 && v !== 0 ? sci(v, 2) : runs(String(v)));
  const rows = keys.map((k) => {
    const p = PRM[k];
    const rng = p.low === null || p.low === undefined ? runs("고정") : runs(val(p.low, k), " – ", val(p.high, k));
    return [TD(p.symbol.replace("₁", "₁"), { bold: true, align: "center", fontSize: 10 }), TD(p.item, { fontSize: 10 }), TD(p.unit, { align: "center", fontSize: 10 }),
      TD(val(p.best, k), { align: "center", fontSize: 10, bold: true }), TD(rng, { align: "center", fontSize: 10 }), TD(p.nfpa551, { align: "center", fontSize: 9.5 }), TD(p.basis, { fontSize: 9.5 })];
  });
  table(s, [[TH("기호", { fontSize: 10 }), TH("항목", { fontSize: 10 }), TH("단위", { fontSize: 10 }), TH("값", { fontSize: 10 }), TH("범위", { fontSize: 10 }), TH("NFPA 551", { fontSize: 10 }), TH("근거 · 대체할 자료", { fontSize: 10 })], ...rows],
    { x: MX, y: 1.55, w: CW, colW: [0.62, 3.05, 0.75, 0.9, 1.45, 1.0, 4.563], fontSize: 10, border: BORDER, rowH: [0.3, ...rows.map(() => 0.295)], margin: [0.01, 0.05, 0.01, 0.05] });
  addNotes(s, ["[스크립트]", "- 계산서 1_입력과 같은 표입니다. 범위는 몬테카를로의 인식론적 범위이고, 분포는 확률은 삼각, 빈도·s는 로그균등입니다."]);
}

// ================================================================================================ 23 appendix sources
{
  const s = content("부록", "부록 B — 근거 자료와 재현 방법", "이 자료의 수치는 아래 스크립트로 다시 만들어지며, 원자료(시험 엑셀) 없이 결과 CSV만으로 계산됨");
  const cw = (CW - 0.25) / 2;
  text(s, para([
    [["NFPA 551 (2022) ", { bold: true }], ["Guide for the Evaluation of Fire Risk Assessments — 1.1–1.5, 3.3, 4.2–4.5, 5.1–5.6, 표 5.1.2.1, 6.2, 7.2–7.5, 8.1–8.3, Annex A(A.5.1.2, A.5.2.5, A.5.5.6)", {}]],
    [["EFFECT ", { bold: true }], ["NFPA·Arup, 2018 — Tier 1·2 구성, 발화원·요인 항목(NFPA·Arup 공개 발표, 업계 보도 기준). 사용자 가이드 원문 미반영", {}]],
    [["HSE R2P2 (2001) ", { bold: true }], ["Reducing risks, protecting people — 개인위험도 10⁻⁶/년(일반적으로 허용), 10⁻⁴/년(일반인 허용 한계) 참고값", {}]],
    [["KS F 8414 ", { bold: true }], ["외벽 실물모형 화재시험(BS 8414 계열) — s 보정과 모델 확인용", {}]],
    [["이 저장소 ", { bold: true }], ["cavity_fire_index(α₆₀₀, FRI, R = P×S), integrated_model(t₆₀₀ = t_ig + t_sp, 플래그 B, TGA 예측식), tga_dsc(PyroSim 입력)", {}]],
  ], { space: 7 }), { x: MX, y: 1.85, w: cw, h: 4.6, fontSize: 11.5 });
  box(s, MX + cw + 0.25, 1.85, cw, 4.6, { fill: "F3F4F6", line: "F3F4F6", r: 0.08, name: "reproduce" });
  text(s, para([
    [["재현 (facade_fra/)", { bold: true }]],
    [["python qfra.py", { fontFace: "Consolas" }], [" → inputs/, results/, figures/, report/qfra_data.json", {}]],
    [["python report/build_qfra_xlsx.py", { fontFace: "Consolas" }], [" → 외벽_QFRA_계산서.xlsx", {}]],
    [["recalc → python report/check_xlsx.py", { fontFace: "Consolas" }], [" → verification.json (엑셀 = 파이썬)", {}]],
    [["node report/build_qfra_deck.js", { fontFace: "Consolas" }], [" → 이 덱", {}]],
    [["node report/build_qfra_docx.js", { fontFace: "Consolas" }], [" → 외벽_QFRA_서술양식.docx", {}]],
    [[`MC: 인식론적 ${D.n_outer.toLocaleString()} × 우연적 ${D.n_inner} (LHS), seed ${D.seed}`, {}]],
  ], { bullet: false, space: 7 }), { x: MX + cw + 0.45, y: 1.98, w: cw - 0.4, h: 4.4, fontSize: 11.5 });
  addNotes(s, ["[스크립트]", "- 모든 수치는 스크립트로 다시 만들어집니다. 입력을 바꿀 때는 qfra.py의 PARAMS와 계산서 1_입력을 같이 바꾸고 check_xlsx.py로 일치를 확인합니다."]);
}

// pptxgenjs writes the X axis of a scatter chart with the Y axis number format; the exceedance charts (log Y in
// 0E+0) need plain integers on X: rewrite the first <c:valAx> (the X axis) of those charts
const JSZip = require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));
async function integerX(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  for (const name of Object.keys(zip.files).filter((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n))) {
    let x = await zip.file(name).async("string");
    if (!/<c:v>M[01] 기준<\/c:v>/.test(x)) continue;
    let first = true;
    x = x.replace(/<c:valAx>[\s\S]*?<\/c:valAx>/g, (blk) => {
      if (!first) return blk;
      first = false;
      return blk.replace(/<c:numFmt formatCode="[^"]*" sourceLinked="1"\/>/, '<c:numFmt formatCode="0" sourceLinked="0"/>');
    });
    zip.file(name, x);
  }
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await finish(OUT);
  await integerX(OUT);
  console.log("wrote", OUT);
})();
