#!/usr/bin/env node
// 외벽 화재위험성 다척도 평가체계 덱 (pptxgenjs, native charts): 소형시료(콘·TGA–DSC·FT-IR) → 중규모 → 시뮬레이션 → 건물 위험도,
// EFFECT(Lamont & Ingolfsson 2018) 검토 포함. Reads ms_data.json (../multiscale.py).
//
//   python ../multiscale.py && NODE_PATH=<folder holding pptxgenjs> node build_ms_deck.js [out.pptx]
"use strict";
const fs = require("fs");
const path = require("path");
const { makeDeck } = require("../../tga_dsc/report/deck_lib");

const D = JSON.parse(fs.readFileSync(path.join(__dirname, "ms_data.json"), "utf8"));
const OUT = process.argv[2] || path.join(__dirname, "외벽_다척도_화재위험성평가_체계.pptx");

const SUBS = {
  R_S: ["R", "S"], p_B: ["p", "B"], p_cb: ["p", "cb"], p_e1: ["p", "e1"], T_FS: ["T", "FS"], "τ_F": ["τ", "F"], t_ig: ["t", "ig"],
  t_sp: ["t", "sp"], "t_ig,sys": ["t", "ig,sys"], C_sys: ["C", "sys"], "r_max,air": ["r", "max,air"], "f₄₀₀,air": ["f", "400,air"],
  "T_onset,air": ["T", "onset,air"], THR_600: ["THR", "600"], t_pHRR: ["t", "pHRR"], k_f: ["k", "f"], t_b: ["t", "b"], "σ_rep": ["σ", "rep"],
  "μ₀": ["μ", "0"], "a₀": ["a", "0"], "b₀": ["b", "0"], "n₀": ["n", "0"], H600: ["H", "600"], D600: ["D", "600"], s_meet: ["s", "meet"],
  s_fail: ["s", "fail"], "t̂₆₀₀": ["t̂", "600"], "p̂_B": ["p̂", "B"], "t̂_TGA": ["t̂", "TGA"], "μ_post": ["μ", "post"], "σ_post": ["σ", "post"],
  p_e: ["p", "e"], C_L: ["C", "L"], "C₁": ["C", "1"], "C₂": ["C", "2"], "F₁": ["F", "1"], "F₂": ["F", "2"], "A₁₄₁₀": ["A", "1410"],
  "A₁₅₉₅": ["A", "1595"], "A_P=O": ["A", "P=O"], "A_기준": ["A", "기준"], "A_OH": ["A", "OH"], "NH₃": ["NH", "3"], "CO₂": ["CO", "2"],
  M_ref: ["M", "ref"], "FIGRA_콘": ["FIGRA", "콘"], T_j: ["T", "j"], s_j: ["s", "j"], "ΔT_j": ["ΔT", "j"],
};
const {
  pres, C, HEX, STATUS, MX, CW, LANG, content: content0, text: text0, vlabel, table, paras, badge, legend, chip, plotMap, curves, finish,
} = makeDeck({
  title: "외벽 화재위험성 다척도 평가체계",
  subject: "소형시료(콘·TGA–DSC·FT-IR), 중규모 성능시험, 시뮬레이션을 연계한 외벽 화재위험성 평가체계와 EFFECT 검토",
  footer: "KICT 중공층 외단열 · 외벽 화재위험성 다척도 평가체계 (소형시료 · 중규모 · 시뮬레이션)",
  themeName: "KICT multi-scale FRA",
  subs: SUBS,
});

// ------------------------------------------------------------------------------------------------ data
const MATS = D.materials, CHAR = D.char, LABEL = D.label;
const ORDER = ["NC", ...MATS];
const MAT = { NC: "8A8F98", EPS_C: "EDA100", EPS_SC: "E87BA4", PIR_C: "6250D6", PIR_SC: "008300", PF_out: "2A78D6", PF_in: "E34948" };
const SHORT = { NC: "불연\n(기준안)", EPS_C: "EPS\n가연성", EPS_SC: "EPS\n준불연", PIR_C: "PIR(PUR)\n가연성", PIR_SC: "PIR(PUR)\n준불연", PF_out: "PF\n외부용", PF_in: "PF\n내부용" };
const RISK = (mode, m, g = "M0") => D.risk.find((r) => r.mode === mode && r.material === m && r.mitigation === g);
const CONS = (m) => D.consistency.find((r) => r.material === m);
const FUS = (m) => D.fusion.find((r) => r.material === m);
const PROF = (m) => D.profile.find((r) => r.material === m);
const DEC = (m, g = "M0") => D.decomposition.find((r) => r.material === m && r.mitigation === g);
const THR = (m, g = "M0") => D.voi_thresholds.find((r) => r.material === m && r.mitigation === g);
const W = D.width[0];
const ST = D.stats;
const VERDICT = { 충족: "good", "판단 유보": "fair", 불충족: "poor" };

// ------------------------------------------------------------------------------------------------ helpers
// "1×10⁻⁴" written with Unicode superscripts becomes runs with a real superscript; "\n" becomes a line break run
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
const SUPD = ["⁰", "¹", "²", "³", "⁴", "⁵", "⁶", "⁷", "⁸", "⁹"];
function supStr(v, d = 2) {
  if (v === 0) return "0";
  let e = Math.floor(Math.log10(Math.abs(v)));
  let m = v / 10 ** e;
  if (Math.abs(Number(m.toFixed(d - 1))) >= 10) { m /= 10; e += 1; }
  return `${m.toFixed(d - 1)}×10${e < 0 ? "⁻" : ""}${String(Math.abs(e)).split("").map((c) => SUPD[+c]).join("")}`;
}
const pct = (v, d = 0) => `${(100 * v).toFixed(d)} %`;
const mn = (v, d = 2) => `${v < 0 ? "−" : ""}${Math.abs(v).toFixed(d)}`;   // typographic minus
const f2 = (v, d = 2) => Number(v).toFixed(d);
const fx = (v, d = 0) => (v === null || v === undefined ? "—" : Number(v).toFixed(d));
function tint(hex, t) {
  return [0, 2, 4].map((i) => Math.round(255 + (parseInt(hex.slice(i, i + 2), 16) - 255) * (1 - t)).toString(16).padStart(2, "0")).join("").toUpperCase();
}
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
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: o.r === undefined ? 0.06 : o.r, fill: { color: o.fill || "FFFFFF" },
    line: { color: o.line || "D1D5DB", width: o.lw || 0.75 }, objectName: o.name });
}
function card(s, x, y, w, h, st, head, items, o = {}) {
  box(s, x, y, w, h, { fill: STATUS[st].tint, line: STATUS[st].tint, r: 0.08, name: `card ${head}` });
  s.addShape(pres.ShapeType.ellipse, { x: x + 0.2, y: y + 0.21, w: 0.2, h: 0.2, fill: { color: STATUS[st].color }, line: { color: STATUS[st].color, width: 0.5 } });
  text(s, head, { x: x + 0.5, y: y + 0.1, w: w - 0.65, h: 0.42, fontSize: o.headSize || 15, bold: true, valign: "middle" });
  text(s, para(items, { space: 4 }), { x: x + 0.22, y: y + 0.58, w: w - 0.42, h: h - 0.7, fontSize: o.size || 12 });
}
const statusChip = (s, x, y, w, kind) => {
  const map = { 측정: ["good", "측정"], 미측정: ["poor", "미측정"], 미수행: ["poor", "미수행"], 보정: ["good", "보정·검증"], 잠정: ["fair", "잠정"],
    제안: ["conditional", "제안"], 가정: ["endpoint", "가정"], 규정: ["fair", "규정 기준"] };
  const [st, lab] = map[kind];
  chip(s, x, y, w, st, lab);
};

// ================================================================================================ 1 title
pres.addSection({ title: "결론" });
{
  const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "결론" });
  s.addText("외벽 화재위험성평가 · 다척도 연계 체계 · 2026-10-23", { placeholder: "kicker", lang: LANG });
  s.addText("소형시료 · 중규모 · 시뮬레이션을 잇는\n외벽 화재위험성 평가체계", { placeholder: "title", lang: LANG });
  s.addText("콘칼로리미터·TGA–DSC·FT-IR → 3 m 중공층 시험 → FDS → 건물 위험도 — 척도마다 위험도를 내고, 같은 단위로 결합",
    { placeholder: "subtitle", lang: LANG });
  const facts = [
    ["3척도 × 6차원", "소형시료 · 중규모 · 시뮬레이션", "착화·확산·발열·지속·급발·유해"],
    [`${f2(W.ratio90_L0, 1)}배 → ${f2(W["ratio90_L1+L2"], 1)}배`, "재료 t₆₀₀ 불확실성(90 % 폭)", "등급만 → TGA → 중규모 → 통합"],
    [`z = ${f2(CONS("PIR_C").z_spread, 1)}`, "게이트 없는 TGA의 가연성 오판", "콘·FT-IR 게이트가 필요한 근거"],
  ];
  facts.forEach(([big, lab, sub], i) => {
    const x = 0.8 + i * 3.95;
    s.addShape(pres.ShapeType.roundRect, { x, y: 4.55, w: 3.7, h: 1.65, rectRadius: 0.08, fill: { color: C.background1, transparency: 90 }, line: { color: C.background2, width: 0.75 } });
    text(s, big, { x: x + 0.25, y: 4.7, w: 3.3, h: 0.6, fontSize: 26, bold: true, color: C.background1 });
    text(s, lab, { x: x + 0.25, y: 5.32, w: 3.3, h: 0.32, fontSize: 13, bold: true, color: C.background1 });
    text(s, sub, { x: x + 0.25, y: 5.66, w: 3.3, h: 0.32, fontSize: 11, color: C.background2 });
  });
  text(s, "함께 드리는 파일: 다척도_위험성평가_계산서.xlsx (콘·FT-IR·시뮬레이션 입력 양식, TGA·중공층 자료, 베이즈 결합·정합성 수식, 위험도 결과)",
    { x: 0.8, y: 6.65, w: 11.7, h: 0.3, fontSize: 11, color: C.background2 });
  addNotes(s, [
    "[스크립트]",
    "- 요청: NFPA 551·EFFECT 같은 틀과 별개로, 소형시료(콘·TGA–DSC·FT-IR), 중규모 성능시험, 시뮬레이션 각각을 통한 위험성평가를 하고 이를 독자 체계로 묶고 싶다. TGA–DSC도 재료 요소로 정량 반영.",
    "- 답: 세 척도가 각자 위험도를 내되, 결과를 모두 '중규모 등가값'(층당 확산시간 t600, 자립 연소 확률 p_B)으로 바꿔 같은 건물 위험도 모델에 넣습니다. 그래서 척도끼리 비교·결합·검사가 됩니다.",
    "- 순서: 결론(2) → EFFECT 검토(3–4) → 체계(5–6) → 소형시료(7–11) → 결합·정합성(12–13) → 척도별 위험도(14–16) → 시뮬레이션(17–18) → 증거 수준·계획·한계(19–21).",
  ]);
}

// ================================================================================================ 2 conclusion
{
  const s = content("결론", "결론 — 무엇을 만들었고, 지금 무엇을 말할 수 있나",
    "척도마다 위험도를 내고 '중규모 등가값'(t₆₀₀ · p_B)으로 결합하는 독자 체계 — TGA–DSC는\n확산 · 지속 · 급발 인자로 위험도에 정량 반영되고, 콘 · FT-IR이 게이트 역할을 맡음");
  const cw = (CW - 0.25) / 2, ch = 2.15;
  const mfs = MATS.map((m) => RISK("L2", m).MF_p50);
  card(s, MX, 1.8, cw, ch, "conditional", "EFFECT 검토: 재료는 범주로만 다룸", [
    "Tier 1 가연 여부, Tier 2 외장재 × 단열재 × 공기층 → Low/Medium/High",
    "보고서 스스로 '정성적'(통계·시험자료 부족) — 콘·FT-IR·TGA는 재료 식별에만 사용",
    `이 시험 5종에 적용하면 모두 연료 High → 판정 D(동일)`,
    `실제: 시험 t₆₀₀ 30–${fx(PROF("PF_out").t600_obs)} s, 재료 인자 ${fx(Math.min(...mfs))}–${fx(Math.max(...mfs))}배`,
  ], { size: 11.5 });
  card(s, MX + cw + 0.25, 1.8, cw, ch, "good", "독자 체계: 척도마다 위험도, 같은 단위로 결합", [
    "L1 소형시료 → L2 중규모 → L3 시뮬레이션 → L4 건물 위험도 R = Σ F·C",
    "공통 단위: 층당 확산시간 t₆₀₀ · 자립 연소 확률 p_B (+ 급발 A)",
    "결합: 베이즈(TGA × 시험) · 검사: 정합성 z · 게이트: 등급 · 조성",
    "6개 위험차원(착화·확산·발열·지속·급발·유해)을 척도마다 측정",
  ], { size: 11.5 });
  card(s, MX, 1.8 + ch + 0.2, cw, ch, "good", "TGA–DSC는 재료 인자로 정량 반영", [
    "r_max,air → t₆₀₀ (확산), f₄₀₀,air → p_B (지속), 용융형 → 급발 표시",
    `TGA만으로(L1) PF 내부용을 가장 위험한 준불연 제품으로 맞춤(R_S ${supStr(RISK("L1", "PF_in").R_S_mean)})`,
    `시험과 결합 시 TGA 무게 ${pct(FUS("PIR_SC").w_tga, 1)} → 재료 t₆₀₀ 90 % 폭 ${f2(W["ratio90_L1+L2"])}배`,
    `지속 몫(대책 없음): PF 외부용 ${pct(DEC("PF_out").persist_part)}, PF 내부용 ${pct(DEC("PF_in").persist_part)}`,
  ], { size: 11.5 });
  card(s, MX + cw + 0.25, 1.8 + ch + 0.2, cw, ch, "poor", "다음: 콘 · FT-IR 게이트, FDS로 s 확정", [
    `게이트 없이 TGA만 쓰면 가연성 PIR(PUR) 위험을 ${f2(RISK("L2", "PIR_C").R_S_mean / RISK("L1u", "PIR_C").R_S_mean, 1)}배 낮게 봄(판정도 바뀜)`,
    "콘(준불연 기준)·FT-IR(난연제)로 가연/준불연을 시료 단계에서 가름",
    `FDS·실대형으로 s를 정하면: EPS ≥ ${f2(THR("EPS_SC").s_meet)} · PIR(PUR) ≥ ${f2(THR("PIR_SC").s_meet)} · PF 외부용 ≥ ${f2(THR("PF_out").s_meet)}이면 충족`,
    "건물 위험도 구간(10–30배)은 s와 빈도가 지배 → 다음 투자 우선",
  ], { size: 11.5 });
  box(s, MX, 6.42, CW, 0.48, { fill: HEX.ink, line: HEX.ink, r: 0.08, name: "caveat" });
  text(s, [{ text: "주의  ", options: { bold: true } }, { text: "빈도·분기확률은 예시값, 콘·FT-IR·FDS는 미수행(엑셀에 양식·계산식 제공) — 수치는 체계 시연용이며, 자료가 들어오면 같은 절차로 갱신됩니다.", options: {} }],
    { x: MX + 0.25, y: 6.42, w: CW - 0.5, h: 0.48, fontSize: 12, color: C.background1, valign: "middle" });
  addNotes(s, [
    "[스크립트]",
    "- 왼쪽 위: EFFECT는 재료를 가연/불연, 그리고 외장재·단열재 계열·공기층 조합의 범주로만 다룹니다. 이 시험 재료에 적용하면 다섯 가지가 모두 같은 판정이 나옵니다.",
    "- 오른쪽 위: 우리 체계는 척도마다 위험도를 내되, 그 결과를 모두 t600과 p_B라는 같은 단위로 바꿉니다.",
    "- 아래: TGA–DSC는 그 단위를 예측하는 '재료 인자'로 들어가고, 게이트가 없으면 가연성 제품을 크게 오판한다는 것이 이번 자료로 확인됩니다. 다음 투자는 콘·FT-IR(게이트)과 FDS(s)입니다.",
  ]);
}

// ================================================================================================ 3 EFFECT review 1
pres.addSection({ title: "EFFECT 검토" });
{
  const s = content("EFFECT 검토", "EFFECT 검토 ① — Lamont & Ingolfsson (2018), NFPA 연구재단 보고서",
    "EFFECT는 기존 고층 건물을 줄 세우는 2단계 '정성' 도구 — 재료는 Tier 1에서 가연 여부,\nTier 2에서 외장재 등급 × 단열재 계열 × 공기층 두께의 범주(Low/Medium/High)로만 들어감");
  const lw = 5.0;
  box(s, MX, 1.85, lw, 4.45, { fill: "F3F4F6", line: "F3F4F6", r: 0.08, name: "effect overview" });
  text(s, "구조", { x: MX + 0.2, y: 1.95, w: lw - 0.4, h: 0.34, fontSize: 14, bold: true });
  text(s, para([
    "Arup 개발 · Jensen Hughes 검토, AHP 설문 34명(Arup 20 · JH 10 · NFPA 패널 4)",
    "대상: 18 m 초과 주거 · 업무, 기존 건물 전용, 생명안전만",
    "Tier 1 데스크톱 설문 → 포트폴리오 우선순위 / Tier 2 도면 · 현장 · 시료 · 실험실 시험 / Tier 3 상세 평가(도구 범위 밖)",
    "Process A 외벽 위험(연료·발화원·수직 연결) / B 피난·경보·구획",
    "판정: PAS 79 행렬 확장 → A 사소 … E 용인 불가",
    "정성 방법을 택한 이유: 위험 순위 · 점수에 쓸 통계 · 시험자료 부족(초록)",
  ], { space: 5 }), { x: MX + 0.2, y: 2.35, w: lw - 0.4, h: 3.9, fontSize: 12 });
  const rx = MX + lw + 0.25, rw = CW - lw - 0.25;
  const rows = [
    ["Tier 1", "단열재 가연? · 외장재 가연? (모름 = High), 수직 연결 여부"],
    ["Tier 2 (그림 33·34)", "외장재(100 % LDPE / A2 / B / HPL·GRP / 목재 / 불연) × 단열재(PIR·페놀 / EPS·XPS·PUR·SPF / 광물면) × 공기층(25·75 mm) → Low·Medium·High"],
    ["AHP 가중치", "외벽 위험 35 % · 피난·경보 37 % · 구획·소화 29 %; 외벽 위험 안에서 구성 재료 29 %(최대), 수직 연결 21 %, 발화원 20 %"],
    ["실험실 시험(11.2.2)", "콘(pHRR·연소열·질량감소), SEM/EDS, FT-IR, TGA/XRD — 재료 '식별'용, 점수에는 미반영"],
    ["변수에서 제외(9.3.12)", "층간 방화 · 캐비티 배리어(AHP 가중치 낮음, 현장 확인 어려움), 열차단재(국가별 규정 차이)"],
    ["후속 과제(13장)", "외장재 · 단열재 · 공기층 조합 점수를 화재시험으로 보강, 소형 발화원 착화, 연기 유해성, Tier 3 체계 → 이 체계의 L2 · 콘(D1) · FT-IR(D6) · 다척도 정량평가가 각각 대응"],
  ];
  table(s, [[TH("항목"), TH("재료가 들어가는 방식")], ...rows.map(([a, b]) => [TD(a, { bold: true }), TD(b)])],
    { x: rx, y: 1.85, w: rw, colW: [1.85, rw - 1.85], fontSize: 11, border: BORDER, rowH: [0.34, 0.5, 0.86, 0.72, 0.62, 0.62, 0.78], margin: [0.04, 0.08, 0.04, 0.08] });
  box(s, MX, 6.42, CW, 0.48, { fill: STATUS.good.tint, line: STATUS.good.tint, r: 0.08, name: "effect gap" });
  text(s, [{ text: "EFFECT가 남긴 빈자리  ", options: { bold: true } }, { text: "= 재료 조합별 시험 근거와 정량 점수 → 이 체계의 소형시료·중규모 자료가 그 자리를 정량으로 채움", options: {} }],
    { x: MX + 0.2, y: 6.42, w: CW - 0.4, h: 0.48, fontSize: 12.5, valign: "middle" });
  addNotes(s, [
    "[스크립트]",
    "- 보고서 초록이 직접 '통계나 시험 자료가 부족해 정성 평가를 택했다'고 밝힙니다. 그래서 재료는 가연 여부와 범주 조합으로만 들어갑니다.",
    "- 흥미로운 점: 콘·FT-IR·TGA를 실험실 시험으로 소개하지만 재료 '식별'에만 쓰고 점수에는 넣지 않습니다. 후속 과제로 '시험으로 조합 점수를 보강하라'고 남겼습니다. 우리 자료가 정확히 그 역할을 합니다.",
  ]);
}

// ================================================================================================ 4 EFFECT review 2
{
  const s = content("EFFECT 검토", "EFFECT 검토 ② — 이 시험 재료에 Tier 2를 적용하면",
    "다섯 재료 모두 연료 High · 판정 D(상당)로 같아짐 — 그러나 같은 D 안에서\n시험 t₆₀₀ 30–1346 s(45배), 재료 인자 33–630배로 갈림 → 범주형은 제품 차이를 못 봄");
  const head = [TH("재료"), TH("EFFECT 단열재 계열"), TH("연료"), TH("가능성"), TH("EFFECT 판정"), TH("시험 t₆₀₀ (s)"), TH("R = P×S (최대)"),
    TH("중규모 기반 R_S (평균)"), TH("재료 인자 MF")];
  const rows = MATS.map((m) => {
    const p = PROF(m), r = RISK("L2", m), e = D.effect.find((x) => x.material === m);
    return [TD(LABEL[m], { bold: true }), TD(e.family, { align: "center" }), TD("High", { align: "center" }), TD("High", { align: "center" }),
      TD("D (상당)", { align: "center", bold: true, fill: { color: "FCE6DC" } }), TD(fx(p.t600_obs), { align: "center" }), TD(String(p.R_max), { align: "center" }),
      TD(supStr(r.R_S_mean), { align: "center" }), TD(fx(r.MF_p50), { align: "center", bold: true })];
  });
  table(s, [head, ...rows], { x: MX, y: 1.85, w: CW, colW: [2.1, 2.3, 0.8, 0.9, 1.35, 1.3, 1.25, 1.45, 0.883], fontSize: 11.5, border: BORDER,
    rowH: [0.5, ...MATS.map(() => 0.44)], margin: [0.04, 0.06, 0.04, 0.06] });
  MATS.forEach((m, i) => s.addShape(pres.ShapeType.rect, { x: MX - 0.14, y: 1.85 + 0.5 + i * 0.44 + 0.07, w: 0.08, h: 0.3, fill: { color: MAT[m] }, line: { color: MAT[m], width: 0.5 } }));
  text(s, para([
    "적용 가정: 불연 외장재, 공기층 100 mm, 15층(30–50 m) 주거·전체 경보(all-out), 스프링클러 설치, 외벽 하부 발화원 있음(발화원 Medium), 연료가 3 m 넘게 수직 연결",
    "EFFECT Tier 2: PIR/페놀은 공기층 75 mm 이하이면 Medium, 이 시험(100 mm)은 High; EPS·XPS·PUR은 공기층과 관계없이 High",
    "재료 인자 MF = 중규모 시험 기반 R_S ÷ 불연 단열재 기준안의 R_S (예시 건물·빈도, 중앙값)",
  ], { space: 4 }), { x: MX, y: 5.15, w: CW, h: 1.2, fontSize: 11, color: C.text2 });
  NOTE(s, "→ 범주형 평가는 '가연성 PIR(PUR) 30 s 전면 착화'와 'PF 외부용 22분 확산'을 같은 D로 묶음. 정량 재료 인자가 필요한 이유.", 6.45, { fontSize: 11.5, color: C.text1, bold: true });
  addNotes(s, [
    "[스크립트]",
    "- EFFECT Tier 2 그림 33·34를 이 시험 구성(공기층 100 mm, 외장재 불연 가정)에 적용하면 다섯 재료 모두 연료 High, 판정 D입니다.",
    "- 같은 D 안에서 실제 거동은 30초 전면 착화부터 20분 넘는 확산까지 갈립니다. EFFECT의 목적(포트폴리오 우선순위)에는 충분해도, 재료·제품 평가에는 부족합니다.",
  ]);
}

// ================================================================================================ 5 framework diagram
pres.addSection({ title: "체계" });
{
  const s = content("체계", "다척도 평가체계 — 척도마다 위험도를 내고, 중규모 등가값으로 결합",
    "소형시료는 중규모를 '예측', 중규모는 '측정', 시뮬레이션은 건물로 '외삽' —\n세 결과가 같은 두 숫자(t₆₀₀ · p_B)로 건물 위험도 모델에 들어가 결합 · 검사됨");
  const cols = [
    ["L1 소형시료", "콘 · TGA–DSC · FT-IR", ["재료 위험 프로파일(6차원)", "전이함수 → 예측 t̂₆₀₀ · p̂_B", "게이트: 등급(콘) · 조성(FT-IR)"], "L1 위험도 (예측)", "FEF2D3"],
    ["L2 중규모", "3 m 중공층 시험", ["t₆₀₀ · t_ig · t_sp, 플래그 A · B", "P · S · R = P × S", "L1 보정 · L3 검증 기준"], "L2 위험도 (측정)", "E2F3E2"],
    ["L3 시뮬레이션", "FDS / PyroSim", ["1단계: 중규모 재현(검증)", "2단계: 건물 외벽 3–6개 층", "층별 도달시각 → s, 창 열류속, FED"], "L3 위험도 (외삽)", "DCEBFA"],
    ["L4 건물 위험도", "NFPA 551 정량법", ["사건수목 × 층간 확산 모델", "2중 몬테카를로", "R = Σ F·C, 초과빈도, 판정"], "최종 판정 + 증거 수준", "F3F4F6"],
  ];
  const gap = 0.62, w = (CW - 3 * gap) / 4, y = 1.85, h = 2.85;
  cols.forEach(([t1, t2, items, out, fill], i) => {
    const x = MX + i * (w + gap);
    box(s, x, y, w, h, { fill, line: fill, r: 0.08, name: t1 });
    text(s, t1, { x: x + 0.18, y: y + 0.1, w: w - 0.36, h: 0.36, fontSize: 15, bold: true });
    text(s, t2, { x: x + 0.18, y: y + 0.46, w: w - 0.36, h: 0.3, fontSize: 11.5, color: C.text2 });
    text(s, para(items, { space: 3 }), { x: x + 0.18, y: y + 0.82, w: w - 0.36, h: 1.35, fontSize: 11.5 });
    box(s, x + 0.15, y + h - 0.55, w - 0.3, 0.42, { fill: HEX.ink, line: HEX.ink, r: 0.06 });
    text(s, out, { x: x + 0.15, y: y + h - 0.55, w: w - 0.3, h: 0.42, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle" });
    if (i < 3) arrow(s, x + w + 0.03, y + h / 2, x + w + gap - 0.03, y + h / 2, { width: 2 });
  });
  ["전이함수", "검증 · s", "τ_F 입력"].forEach((lab, i) => text(s, lab, { x: MX + (i + 1) * w + i * gap, y: y + h / 2 - 0.36, w: gap, h: 0.28,
    fontSize: 9, color: C.text2, align: "center" }));
  box(s, MX, 4.88, CW, 0.52, { fill: "FFFFFF", line: HEX.ink, lw: 1.5, r: 0.08, name: "common unit" });
  text(s, [{ text: "공통 단위 — 중규모 등가값:  ", options: { bold: true } }, { text: "층당 확산시간 t₆₀₀ (3000 mm 600 K 도달) · 자립 연소 확률 p_B (버너 소화 후)  + 급발 A 표시", options: {} }],
    { x: MX + 0.25, y: 4.88, w: CW - 0.5, h: 0.52, fontSize: 13, valign: "middle", align: "center" });
  const bw = (CW - 2 * 0.25) / 3;
  [["게이트", "등급(콘 준불연 기준) · 조성(FT-IR 고분자·난연제) → 전이함수 적용 여부"],
    ["결합", "베이즈: ln t₆₀₀ 정규–정규, p_B 베타–이항 → 통합 사후분포"],
    ["검사 · 증거 수준", "척도 간 정합성 z, 예측확률 / E0 등급 → E1 소형 → E2 중규모 → E3 시뮬레이션"]].forEach(([h1, b1], i) => {
    const x = MX + i * (bw + 0.25);
    box(s, x, 5.55, bw, 0.85, { fill: "F3F4F6", line: "F3F4F6", r: 0.08, name: h1 });
    text(s, h1, { x: x + 0.18, y: 5.6, w: bw - 0.36, h: 0.3, fontSize: 12.5, bold: true });
    text(s, b1, { x: x + 0.18, y: 5.9, w: bw - 0.36, h: 0.48, fontSize: 11 });
  });
  NOTE(s, "각 척도의 위험도는 같은 건물·빈도(facade_fra 예시) 위에서 재료 입력만 바꿔 계산 → 척도 사이의 차이는 순수하게 '재료에 대한 지식'의 차이입니다.", 6.5);
  addNotes(s, [
    "[스크립트]",
    "- 네 칸이 네 척도입니다. 각 칸 아래 검은 띠가 그 척도가 내는 위험도입니다.",
    "- 핵심은 가운데 띠: 모든 척도가 t600과 p_B라는 같은 두 숫자로 말합니다. 소형시료는 이를 예측하고, 중규모는 측정하고, 시뮬레이션은 건물 규모로 옮깁니다.",
    "- 그래서 결합(베이즈), 검사(정합성), 증거 수준 표시가 가능합니다. 게이트는 예측식을 써도 되는 재료인지를 가립니다.",
  ]);
}

// ================================================================================================ 6 dimension x scale
{
  const s = content("체계", "6개 위험차원 × 척도 — 무엇으로 재고, 지금 무엇이 있나",
    "확산(D2) · 지속(D4)은 소형 · 중규모 자료가 모두 있어 위험도에 정량 반영 —\n착화 · 발열 · 유해(D1 · D3 · D6)는 콘 · FT-IR · FDS가 들어와야 채워지는 칸");
  const R = [
    ["D1 착화성", "t_ig, TRP·CHF", "T_onset,air (예측력 약함: 시스템 지배)", "난연제 유무", "t_ig 188–288 s (준불연)", "외벽 착화시각"],
    ["D2 확산성", "pHRR/t_pHRR, b, M", "r_max,air → t̂₆₀₀ (T1)", "이소시아누레이트 지수", "t_sp, t₆₀₀, FRI, P", "층별 도달시각 → s"],
    ["D3 발열성", "THR_600, pHRR, EHC", "절대량 불가(초기 질량 없음)", "—", "H600, 최고온도", "외벽 HRR, 창 열류속"],
    ["D4 지속성", "연소 종료·작열", "f₄₀₀,air → p̂_B (T2)", "—", "플래그 B", "소화 후 연소"],
    ["D5 급발성", "용융·적하", "용융형(f₄₀₀ < 0.5) (T3)", "—", "플래그 A", "(모델 한계)"],
    ["D6 유해성", "CO 수율, TSP", "—", "HCN·CO·NCO·HCl (EGA)", "(미측정)", "위층 FED"],
    ["G 조성·등급", "준불연 판정", "탄화형/용융형", "고분자·난연제 식별", "가연성 30 s 전면 착화", "—"],
  ];
  const ST = [
    [0, 1, 0, 1, 0], [0, 1, 0, 1, 0], [0, 2, 3, 1, 0], [0, 1, 3, 1, 0], [0, 1, 3, 1, 2], [0, 3, 0, 0, 0], [0, 1, 0, 1, 3],
  ];   // 1 have, 0 missing, 2 not possible, 3 n/a
  const fillOf = (k) => ({ 1: "E2F3E2", 0: "FDE2E2", 2: "F3F4F6", 3: "F3F4F6" })[k];
  const head = [TH("위험 차원"), TH("L1 콘"), TH("L1 TGA–DSC"), TH("L1 FT-IR"), TH("L2 중규모"), TH("L3 시뮬레이션")];
  const body = R.map((row, i) => row.map((v, j) => (j === 0 ? TD(v, { bold: true }) : TD(v, { fontSize: 10.5, fill: { color: fillOf(ST[i][j - 1]) } }))));
  table(s, [head, ...body], { x: MX, y: 1.8, w: CW, colW: [1.55, 2.0, 2.55, 2.0, 2.15, 2.083], fontSize: 10.5, border: BORDER,
    rowH: [0.36, ...R.map(() => 0.52)], margin: [0.03, 0.07, 0.03, 0.07] });
  legend(s, [{ kind: "box", color: "E2F3E2", label: "자료 있음", w: 0.95 }, { kind: "box", color: "FDE2E2", label: "시험·해석 필요", w: 1.3 },
    { kind: "box", color: "F3F4F6", label: "해당 없음·불가", w: 1.3 }], MX, 5.92);
  NOTE(s, "건물 모델에 들어가는 자리: D1·D2 → t₆₀₀ = t_ig + t_sp (τ_F = s·t₆₀₀), D4 → p_B, D5 → 표시(급발), D3·D6 → 확장 항목(창 파손·유해 노출).", 6.32, { h: 0.5 });
  addNotes(s, [
    "[스크립트]",
    "- 행은 위험의 종류, 열은 척도입니다. 같은 위험을 척도마다 다른 시험으로 잽니다.",
    "- 연두 칸이 지금 있는 자료, 분홍 칸이 앞으로 시험·해석으로 채울 칸입니다. 확산·지속은 소형과 중규모 모두 있어 정량 결합이 됩니다.",
  ]);
}

// ================================================================================================ 7 L1 roles
pres.addSection({ title: "L1 소형시료" });
{
  const s = content("L1 소형시료", "L1 소형시료 평가 — 세 시험의 역할과 사각지대",
    "콘은 착화 · 발열, TGA–DSC는 분해 · 산화, FT-IR은 조성을 봄 —\n서로의 사각지대를 메워야 소형시료로 중규모 거동을 예측할 수 있음");
  const cw = (CW - 2 * 0.25) / 3, y = 1.85, h = 3.05;
  [["콘칼로리미터 (ISO 5660-1)", "fair", ["보는 것: 착화시간, 열방출률(pHRR), 총방출열(THR), 연기·CO", "국내 준불연 판정이 콘 기준 → 가연/준불연 게이트", "TRP·CHF(다열류속), b·M 지수로 상향 확산 예측", "못 보는 것: 분해 단계·탄화층 산화의 온도 구조"]],
    ["TGA–DSC (보유)", "good", ["보는 것: 분해속도 r_max,air, 고온 산화 몫 f₄₀₀,air, 용융/탄화형", "준불연 탄화형의 확산시간 예측(T1), 지속(T2), 급발(T3)", "FDS 반응 입력(PyroSim 자료)", "못 보는 것: 기상 난연 효과, 발열 절대량(초기 질량 없음)"]],
    ["FT-IR (ATR · TGA–FTIR)", "conditional", ["보는 것: 고분자 종류, 난연제(P=O·C–Cl), 무기 충전재", "제품 식별 → 전이함수 적용 여부(게이트 검증)", "발생가스(HCN·CO·NCO·HCl) → 유해성, FDS 수율", "못 보는 것: 연소 속도·열량(정성·반정량)"]],
  ].forEach(([h1, st, items], i) => card(s, MX + i * (cw + 0.25), y, cw, h, st, h1, items, { size: 12, headSize: 14 }));
  const pc = PROF("PIR_C"), ps = PROF("PIR_SC"), cc = CONS("PIR_C"), cs = CONS("PIR_SC");
  box(s, MX, 5.1, CW, 1.75, { fill: "FFFFFF", line: HEX.ink, lw: 1.25, r: 0.08, name: "pir example" });
  text(s, para([
    [["사각지대가 실제로 드러난 예 — 가연성 PIR(PUR)", { bold: true }]],
    [[`TGA로는 준불연과 큰 차이 없음: 분해 시작은 오히려 높고(공기 ${fx(pc.T_onset_air)} vs ${fx(ps.T_onset_air)} °C), r_max는 ${f2(pc.r_max_air)} vs ${f2(ps.r_max_air)} %/°C → 예측 t₆₀₀ ${fx(cc.t600_tga_s)} s vs ${fx(cs.t600_tga_s)} s`, {}]],
    [[`중공층 시험은 ${fx(cc.t600_obs_geo_s)} s vs ${fx(cs.t600_obs_geo_s)} s(${fx(cs.t600_obs_geo_s / cc.t600_obs_geo_s)}배 차이), TGA 예측의 1/${fx(cc.t600_tga_s / cc.t600_obs_geo_s)} — 차이는 질량감소가 아닌 기상 난연(화염 억제)에서 옴`, {}]],
    [["→ 콘(열방출 · 준불연 판정)과 FT-IR(난연제)이 봐야 하는 영역 = 게이트", { bold: true }]],
  ], { bullet: false, space: 5 }), { x: MX + 0.25, y: 5.2, w: CW - 0.5, h: 1.6, fontSize: 12.5 });
  addNotes(s, [
    "[스크립트]",
    "- 세 시험은 서로 다른 것을 봅니다. 콘은 착화와 열방출, TGA는 분해와 산화의 온도 구조, FT-IR은 화학 조성입니다.",
    `- 아래 예가 핵심입니다. 가연성 PIR(PUR)은 TGA상 준불연과 큰 차이가 없는데(분해 시작은 오히려 높음), 중공층에서는 준불연보다 ${fx(cs.t600_obs_geo_s / cc.t600_obs_geo_s)}배 빨리 번졌습니다. 난연 처리의 효과는 질량감소가 아니라 불꽃 억제에서 나타나기 때문입니다. 그래서 콘과 FT-IR이 게이트로 필요합니다.`,
  ]);
}

// ================================================================================================ 8 cone
{
  const s = content("L1 소형시료", "L1 콘칼로리미터 — 지표, 계산식, 게이트 (엑셀 2_L1_콘 양식)",
    "50 kW/m² 10분 결과로 가연/준불연을 가르고(게이트), 다열류속 착화시간과 열방출로\n착화(TRP·CHF) · 확산(b, M) · 발열(THR) 지표를 계산 — 지금은 자료 없음, 양식과 수식만 준비");
  const lw = 7.0;
  const rows = [
    ["착화 (D1)", "t_ig(50), TRP = 1/기울기(1/√t_ig vs q″), CHF = x절편", "다열류속 25·35·50·65 kW/m²"],
    ["성장 (D2)", "FIGRA_콘 = pHRR / t_pHRR, t_ig / pHRR (Petrella)", "Schartel & Hull 2007"],
    ["상향 확산 (D2)", "b = k_f·Q̇″ − 1 − t_b/t_ig, k_f ≈ 0.01 m²/kW, t_b = THR/pHRR", "b > 0 → 가속 (Quintiere)"],
    ["중규모 등가 (D2)", "M = Q̇″^(1/3)/TRP → t̂₆₀₀ = 1079 s × (M_ref / M)², ref = PIR(PUR) 준불연", "RISK_MODEL 상대식(잠정)"],
    ["발열 (D3)", "THR_600, pHRR, EHC = THR·A/Δm", "A = 0.0088 m²"],
    ["유해 (D6)", "CO 수율, TSP (ISO 5660-2)", "FDS 수율"],
  ];
  table(s, [[TH("차원"), TH("지표 · 식"), TH("비고")], ...rows.map(([a, b, c]) => [TD(a, { bold: true }), TD(b), TD(c, { fontSize: 10.5, color: HEX.ink2 })])],
    { x: MX, y: 1.85, w: lw, colW: [1.45, 3.85, 1.7], fontSize: 11, border: BORDER, rowH: [0.34, ...rows.map(() => 0.6)], margin: [0.04, 0.08, 0.04, 0.08] });
  const rx = MX + lw + 0.25, rw = CW - lw - 0.25;
  card(s, rx, 1.85, rw, 2.35, "fair", "게이트: 국내 준불연 기준 (50 kW/m²)", [
    "10분 총방출열량 THR_600 ≤ 8 MJ/m²",
    "200 kW/m² 초과가 10초 이상 연속되지 않을 것",
    "관통 균열·구멍·용융 없음 (+ 가스유해성 별도)",
    "미달 → 가연성: TGA 전이함수 미적용, 중규모 필수",
  ], { size: 11.5, headSize: 13.5 });
  card(s, rx, 4.35, rw, 2.0, "conditional", "시험 조건 (권장)", [
    "ISO 5660-1, 100 × 100 mm, 면재 포함 · 심재 각각",
    "50 kW/m² 10분 3회 + 25·35·65 kW/m² 착화",
    "같은 제품의 중공층 시험과 짝지어 전이함수 보정",
  ], { size: 11.5, headSize: 13.5 });
  NOTE(s, "준불연 기준: 국토교통부 고시 「건축자재등 품질인정 및 관리기준」. 콘 자료가 들어오면 엑셀이 지표 · 게이트 · 콘 기반 t̂₆₀₀을 바로 계산합니다(회색 예시 행 = 가상값).", 6.5, { h: 0.45 });
  addNotes(s, [
    "[스크립트]",
    "- 콘은 국내 준불연 판정의 근거 시험이라 게이트로 바로 씁니다. 등급 인증서 대신 같은 시료로 직접 측정하면 c/sc 판단이 데이터가 됩니다.",
    "- 다열류속 착화시간으로 TRP와 CHF를, 열방출로 상향 확산 인자 b와 재료지수 M을 계산합니다. M은 앞서 제안한 상대식으로 중규모 t600을 예측합니다(잠정, 보정 필요).",
  ]);
}

// ================================================================================================ 9 FT-IR
{
  const s = content("L1 소형시료", "L1 FT-IR — 식별 · 난연제 · 발생가스 (엑셀 4_L1_FT-IR 양식)",
    "ATR로 고분자 · 난연제를 확인해 게이트를 검증하고(선언 제품 = 실제 시료?),\nTGA–FTIR로 분해 가스(HCN · CO · NCO · HCl)를 측정해 유해성 · FDS 수율에 씀 — 양식 준비");
  const lw = 6.6;
  const bands = [
    ["EPS", "756 · 698, 1601 · 1493 · 1452", "일치환 벤젠, 방향족 고리"],
    ["PIR/PUR", "1410 / 1595", "이소시아누레이트 C–N / 방향족(기준)"],
    ["PIR/PUR", "1710–1730, 1510–1530, 2270", "C=O, 아마이드 II, 잔류 NCO"],
    ["PF", "3300(넓음), 1475 · 1440, 1230", "O–H, 메틸렌 가교, 페놀 C–O"],
    ["인계 난연제", "1270–1290, 1030–1050", "P=O, P–O–C (TCPP 등)"],
    ["ATH", "3620 · 3523 · 3450", "O–H (무기 충전재)"],
  ];
  table(s, [[TH("대상"), TH("대역 (cm⁻¹)"), TH("귀속")], ...bands.map(([a, b, c]) => [TD(a, { bold: true }), TD(b), TD(c, { fontSize: 10.5 })])],
    { x: MX, y: 1.85, w: lw, colW: [1.25, 2.45, 2.9], fontSize: 11, border: BORDER, rowH: [0.34, ...bands.map(() => 0.44)], margin: [0.03, 0.08, 0.03, 0.08] });
  text(s, "대표 흡수대(문헌값) — 실험실 기준 스펙트럼으로 확인 후 사용", { x: MX, y: 4.95, w: lw, h: 0.26, fontSize: 10, color: C.text2 });
  const rx = MX + lw + 0.25, rw = CW - lw - 0.25;
  card(s, rx, 1.85, rw, 2.25, "conditional", "정량 지수", [
    "이소시아누레이트 지수 = A₁₄₁₀ / A₁₅₉₅ (PIR 정도 ↔ T_onset · f₄₀₀)",
    "난연(P) 지수 = A_P=O / A_기준, 무기충전 지수 = A_OH / A_기준",
    "EGA: 가스별 적분 면적 / 시료 mg (같은 조건 상대 비교)",
  ], { size: 11.5, headSize: 13.5 });
  card(s, rx, 4.25, rw, 2.1, "fair", "게이트 규칙 (제안)", [
    "선언 고분자 ≠ 판정 → 제품 확인 전 전이함수 보류",
    "난연제 미검출 → 인증서와 무관하게 가연성으로 다룸(보수적)",
    "브롬계는 적외선 감도 낮음 → XRF · SEM/EDS 병행",
  ], { size: 11.5, headSize: 13.5 });
  text(s, para([
    "EGA 대역: CO 2180 · 2110, HCN 713, NH₃ 966 · 931, –NCO 2270, HCl 2700–3100, CO₂ 2360",
  ], { bullet: false }), { x: MX, y: 5.3, w: lw, h: 0.6, fontSize: 11 });
  NOTE(s, "EFFECT 보고서(11.2.2)도 FT-IR을 '재료 식별'에 쓰라고 권하지만 점수에는 넣지 않음 → 이 체계는 식별 결과를 게이트로, 가스를 유해성 입력으로 씀.", 6.45, { h: 0.45 });
  addNotes(s, [
    "[스크립트]",
    "- FT-IR의 첫 역할은 게이트 검증입니다. 시험한 시료가 선언한 제품과 같은지, 난연제가 들어 있는지를 봅니다.",
    "- 이소시아누레이트 지수처럼 정량 지수도 만들 수 있어, TGA의 분해 시작 온도·고온 산화 몫과 연결해 볼 수 있습니다. 발생가스 분석은 시뮬레이션의 유해가스 입력이 됩니다.",
  ]);
}

// ================================================================================================ 10 TGA results
{
  const s = content("L1 소형시료", "L1 TGA–DSC (보유 자료) → 재료 인자와 중규모 등가 예측",
    "준불연 탄화형 3종은 TGA 전이함수로 t₆₀₀ · p_B를 예측 — 가연성은 게이트로 제외,\nEPS 준불연은 용융형이라 확산 예측 대신 급발(A)만 표시");
  const head = [TH("재료"), TH("등급"), TH("T_onset,air (°C)"), TH("r_max,air (%/°C)"), TH("f₄₀₀,air"), TH("자료 판정"), TH("재료형"), TH("게이트"),
    TH("t̂₆₀₀ (s)"), TH("P̂"), TH("p̂_B (사전)"), TH("급발")];
  const ord = ["EPS_C", "EPS_SC", "PIR_C", "PIR_SC", "PF_out", "PF_in"];
  const vko = { reliable: "신뢰", conditional: "조건부", endpoint: "종점 의존", remeasure: "재측정" };
  const rows = ord.map((m) => {
    const p = PROF(m), grade = D.grade[m] || "c", melt = p.melt;
    const gate = grade === "c" ? "가연성 → 미적용" : melt ? "용융형 → 확산 미적용" : "적용";
    const use = gate === "적용";
    const t = use ? (D.fusion.find((f) => f.material === m) || {}).t600_tga_loo || p.t600_tga : p.t600_tga;
    const fri = 900 / t, P = fri >= 2 ? 4 : fri >= 1 ? 3 : fri >= 0.5 ? 2 : 1;
    return [TD(LABEL[m], { bold: true }), TD(grade, { align: "center" }), TD(fx(p.T_onset_air), { align: "center" }), TD(f2(p.r_max_air, 3), { align: "center" }),
      TD(f2(p.f400_air, 3), { align: "center" }), TD(vko[p.tga_verdict] || p.tga_verdict, { align: "center", fontSize: 10.5 }),
      TD(melt ? "용융형" : "탄화형", { align: "center" }), TD(gate, { fontSize: 10.5, fill: { color: use ? "E2F3E2" : "F3F4F6" } }),
      TD(use ? fx(t) : `(${fx(t)})`, { align: "center", bold: use, color: use ? HEX.ink : HEX.ink2 }), TD(use ? String(P) : "—", { align: "center" }),
      TD(f2(p.pB_tga, 2), { align: "center" }), TD(melt ? "A (용융)" : grade === "c" ? "A (가연)" : "—", { fontSize: 10.5, align: "center" })];
  });
  const rh10 = 0.5;
  table(s, [head, ...rows], { x: MX, y: 1.85, w: CW, colW: [1.75, 0.6, 1.1, 1.15, 0.85, 0.95, 0.85, 1.75, 0.85, 0.55, 0.95, 1.033], fontSize: 11, border: BORDER,
    rowH: [0.5, ...ord.map(() => rh10)], margin: [0.03, 0.05, 0.03, 0.05] });
  ord.forEach((m, i) => s.addShape(pres.ShapeType.rect, { x: MX - 0.14, y: 1.85 + 0.5 + i * rh10 + 0.1, w: 0.08, h: 0.3, fill: { color: MAT[m] }, line: { color: MAT[m], width: 0.5 } }));
  text(s, para([
    `t̂₆₀₀ = t_ig,sys + C_sys / r_max,air (t_ig,sys ${fx(D.constants.t_ig_sys)} s, C_sys ${fx(D.constants.c_sys)} s·%/°C) — 적용 3종은 해당 제품을 빼고 맞춘 상수로 예측(괄호 = 게이트로 미적용)`,
    "p̂_B: f₄₀₀,air 0.60 → 0.70에서 0.05 → 0.95 램프, 무게 2회 시험분(잠정) · 급발: 용융형(f₄₀₀ < 0.5) 또는 가연성 전면 착화",
    "EPS 가연성: 공기 자료 '재측정' + 가연성 + 용융형 → 소형시료로 판정 불가, 중규모 재시험(30분) 필요",
  ], { space: 4 }), { x: MX, y: 5.5, w: CW, h: 1.35, fontSize: 11.5, color: C.text2 });
  addNotes(s, [
    "[스크립트]",
    "- 지금 있는 TGA 자료로 재료 인자를 계산한 표입니다. 게이트를 통과한 준불연 탄화형 3종만 예측 t600을 냅니다.",
    "- 예측 P 등급은 PF 내부용 3, 나머지 2로 중규모 시험 결과와 같습니다(교차검증 6/6).",
  ]);
}

// ================================================================================================ 11 transfer functions
{
  const s = content("L1 소형시료", "전이함수 — 소형시료 지표를 중규모 등가값으로 바꾸는 식",
    "TGA 3식(T1–T3)은 이번 자료로 보정 · 확인됨, 콘 · FT-IR 5식(T4–T8)은 제안 단계 —\n중규모 → 건물(s)은 가정이며 L3(시뮬레이션)·실대형 시험으로 대체");
  const rows = [
    ["T1", "TGA → 확산", "t̂₆₀₀ = t_ig,sys + C_sys / r_max,air", "준불연 · 탄화형", "보정", "D2"],
    ["T2", "TGA → 지속", "p̂_B ~ Beta(n₀m, n₀(1−m)), m = 램프(f₄₀₀,air)", "준불연 · 탄화형", "잠정", "D4"],
    ["T3", "TGA → 급발", "f₄₀₀,air < 0.5 → 용융형: 확산식 미적용, A 표시", "전체", "보정", "D5"],
    ["T4", "콘 → 게이트", "준불연 기준 미달 → 가연성: TGA 미적용", "전체", "규정", "G"],
    ["T5", "콘 → 확산", "M = Q̇″^(1/3)/TRP → t̂₆₀₀ ∝ M⁻², b > 0 → 가속", "준불연", "제안", "D2"],
    ["T6", "콘 → 발열", "THR_600 ↔ H600 (3000 mm 초과 열량)", "전체", "제안", "D3"],
    ["T7", "FT-IR → 게이트", "식별 일치 · 난연제 검출 → 준불연 처리 확인", "전체", "제안", "G"],
    ["T8", "FT-IR·콘 → 유해", "HCN · CO 발생량 → FDS 수율 → 위층 FED", "전체", "제안", "D6"],
    ["S", "중규모 → 건물", "τ_F = s · t₆₀₀ · h / 3 m", "전체", "가정", "D2"],
  ];
  const y0 = 1.8, rh = 0.42;
  table(s, [[TH("번호"), TH("연결"), TH("식"), TH("적용 범위"), TH("상태"), TH("차원")], ...rows.map(([a, b, c, d, e, f]) => [TD(a, { bold: true, align: "center" }),
    TD(b, { bold: true }), TD(c), TD(d, { align: "center" }), TD(""), TD(f, { align: "center" })])],
  { x: MX, y: y0, w: CW, colW: [0.7, 1.9, 5.6, 1.6, 1.65, 0.883], fontSize: 11.5, border: BORDER, rowH: [0.36, ...rows.map(() => rh)], margin: [0.02, 0.08, 0.02, 0.08] });
  rows.forEach((r, i) => statusChip(s, MX + 0.7 + 1.9 + 5.6 + 1.6 + 0.12, y0 + 0.36 + i * rh + (rh - 0.3) / 2, 1.4, r[4]));
  NOTE(s, `T1 교차검증(제품 제외): 오차 −14 ~ +29 %, P 등급 6/6 일치 · TGA 예측 편차 τ = ${f2(ST.tau, 3)} (ln), 시험 반복 편차 σ_rep = ${f2(ST.s_rep, 3)} (ln)`, 6.2);
  NOTE(s, "엑셀 6_전이함수에 상수·통계가 수식으로 들어 있어, 시험이 늘면 σ_rep · τ가 자동으로 다시 계산됩니다.", 6.52);
  addNotes(s, [
    "[스크립트]",
    "- 전이함수가 이 체계의 '번역기'입니다. 소형시료 숫자를 중규모 등가값으로 바꿉니다.",
    "- TGA 식은 이번 자료로 보정과 교차검증까지 했고, 콘·FT-IR 식은 자료가 들어오면 같은 방법으로 보정합니다. 건물로 가는 s는 아직 가정입니다.",
  ]);
}

// ================================================================================================ 12 consistency
pres.addSection({ title: "결합 · 정합성" });
{
  const s = content("결합 · 정합성", "척도 간 정합성 — TGA가 예측한 t₆₀₀과 중공층 시험의 비교",
    `준불연 탄화형 3종은 예측 구간 안(|z| ≤ ${f2(Math.max(...CHAR.map((m) => Math.abs(CONS(m).z_spread))), 1)}), 가연성 PIR(PUR)은 z ${mn(CONS("PIR_C").z_spread, 1)}로 완전히 벗어남 —\n게이트 없이 TGA를 쓰면 가장 위험한 제품을 가장 크게 오판함`);
  const ch = { x: MX, y: 1.8, w: 6.3, h: 4.75, L: { x: 0.15, y: 0.05, w: 0.8, h: 0.8 } };
  const xMin = 200, xMax = 2000, yMin = 20, yMax = 3000;   // log axes tick at min·10ⁿ: 200 · 2000 and 20 · 200 · 2000
  const lx = (v) => ch.x + ch.w * ch.L.x + ch.w * ch.L.w * (Math.log10(v) - Math.log10(xMin)) / (Math.log10(xMax) - Math.log10(xMin));
  const ly = (v) => ch.y + ch.h * ch.L.y + ch.h * ch.L.h * (1 - (Math.log10(v) - Math.log10(yMin)) / (Math.log10(yMax) - Math.log10(yMin)));
  const sd1 = Math.sqrt(ST.tau ** 2 + ST.s_rep ** 2);
  const xs = [], series = [];
  const N = 2 + 2 + MATS.length * 3;
  const slot = () => { const a = new Array(N).fill(null); return a; };
  // 1:1 and band
  xs.push(xMin, xMax);
  const one = slot(); one[0] = xMin; one[1] = xMax;
  series.push({ name: "1:1", ys: one, color: HEX.ink2, style: { noMarker: true } });
  xs.push(xMin, xMax);
  const lo = slot(); lo[2] = xMin * Math.exp(-1.96 * sd1); lo[3] = xMax * Math.exp(-1.96 * sd1);
  series.push({ name: "band lo", ys: lo, color: "9CA3AF", style: { noMarker: true, dash: "dash" } });
  const hi = slot(); hi[2] = xMin * Math.exp(1.96 * sd1); hi[3] = xMax * Math.exp(1.96 * sd1);
  series.push({ name: "band hi", ys: hi, color: "9CA3AF", style: { noMarker: true, dash: "dash" } });
  MATS.forEach((m) => {
    const c = CONS(m), p = PROF(m);
    const base = xs.length;
    const x = c.t600_tga_s;
    const used = c.spread_used;
    const obs = D.specs.L2[m].base;
    [0, 1, 2].forEach((k) => xs.push(x));
    const ys = slot();
    obs.forEach((v, k) => { if (v !== null && k < 3) ys[base + k] = v; });
    // gated-out products: light fill + coloured ring (an unfilled marker loses its ring in renderers that take it from the series line)
    series.push({ name: `pt ${m}`, ys, color: MAT[m], style: { noLine: true, marker: { symbol: "circle", size: 9, fill: used ? MAT[m] : tint(MAT[m], 0.35), line: MAT[m], lineW: 19050 } } });
  });
  curves(s, ch, xs, series, { lineSize: 1.5, valAxisLogScaleBase: 10, valAxisLabelFormatCode: "0", catAxisMinVal: xMin, catAxisMaxVal: xMax, valAxisMinVal: yMin, valAxisMaxVal: yMax,
    valAxisMajorUnit: undefined, catAxisMajorUnit: undefined });
  text(s, "TGA 예측 t₆₀₀ (s, 로그)", { x: ch.x + ch.w * ch.L.x, y: ch.y + ch.h * (ch.L.y + ch.L.h) + 0.3, w: ch.w * ch.L.w, h: 0.28, fontSize: 11, color: C.text2, align: "center" });
  vlabel(s, "중공층 시험 t₆₀₀ (s, 로그)", ch.x + 0.17, ch.y + ch.h * (ch.L.y + ch.L.h / 2), 2.6);
  const labPos = { PIR_C: [12, 0], EPS_SC: [-120, -18], PIR_SC: [10, 10], PF_out: [-120, -22], PF_in: [10, 6] };
  MATS.forEach((m) => {
    const c = CONS(m);
    const yv = c.t600_obs_geo_s;
    const [dx, dy] = labPos[m];
    text(s, LABEL[m], { x: lx(c.t600_tga_s) + dx / 72, y: ly(yv) - 0.12 + dy / 72, w: 1.65, h: 0.26, fontSize: 10, color: HEX.ink, align: dx < 0 ? "right" : "left" });
  });
  // non-reach EPS test marker
  const xe = CONS("EPS_SC").t600_tga_s;
  s.addShape(pres.ShapeType.triangle, { x: lx(xe) - 0.07, y: ly(2400) - 0.07, w: 0.14, h: 0.13, fill: { color: "FFFFFF" }, line: { color: MAT.EPS_SC, width: 1.5 } });
  text(s, "미도달(> 30분)", { x: lx(xe) + 0.1, y: ly(2400) - 0.13, w: 1.3, h: 0.26, fontSize: 9.5, color: C.text2 });
  legend(s, [{ kind: "dot", color: HEX.ink2, label: "전이함수 적용", w: 1.15 }], ch.x + 0.9, 6.55);
  s.addShape(pres.ShapeType.ellipse, { x: ch.x + 3.25, y: 6.61, w: 0.16, h: 0.16, fill: { color: tint(HEX.ink2, 0.35) }, line: { color: HEX.ink2, width: 1.5 } });
  text(s, "게이트로 제외", { x: ch.x + 3.48, y: 6.55, w: 1.2, h: 0.28, fontSize: 11, color: C.text2, valign: "middle" });
  s.addShape(pres.ShapeType.line, { x: ch.x + 4.75, y: 6.69, w: 0.36, h: 0, line: { color: "9CA3AF", width: 1.5, dashType: "dash" } });
  text(s, "95 % 구간", { x: ch.x + 5.18, y: 6.55, w: 0.9, h: 0.28, fontSize: 11, color: C.text2, valign: "middle" });
  // right: table
  const rx = MX + 6.55, rw = CW - 6.55;
  const head = [TH("재료", { fontSize: 10.5 }), TH("게이트", { fontSize: 10.5 }), TH("z 확산", { fontSize: 10.5 }), TH("P(k) 지속", { fontSize: 10.5 }), TH("급발", { fontSize: 10.5 })];
  const rows = MATS.map((m) => {
    const c = CONS(m);
    // |z| ≤ 2 consistent; z < −2 TGA under-predicts the hazard (non-conservative); z > 2 over-predicts it (conservative)
    const zst = c.z_spread < -2 ? ["B42318", "F8DEDE"] : c.z_spread > 2 ? ["8A5A00", "FDF0D5"] : [HEX.ink, "E2F3E2"];
    return [TD(LABEL[m], { bold: true, fontSize: 10.5 }), TD(c.gate, { fontSize: 10 }),
      TD(mn(c.z_spread), { align: "center", fontSize: 10.5, bold: Math.abs(c.z_spread) > 2, color: zst[0], fill: { color: zst[1] } }),
      TD(c.p_pred_B === null ? "n/a" : f2(c.p_pred_B, 2), { align: "center", fontSize: 10.5, fill: { color: c.p_pred_B === null ? "F3F4F6" : c.p_pred_B >= 0.05 ? "E2F3E2" : "F8DEDE" } }),
      TD(`${c.melt_type ? "용융" : "—"} / A ${c.A_count}/${c.n_tests}`, { align: "center", fontSize: 10 })];
  });
  table(s, [head, ...rows], { x: rx, y: 1.85, w: rw, colW: [1.55, 1.45, 0.75, 0.85, 1.183], fontSize: 10.5, border: BORDER, rowH: [0.34, ...MATS.map(() => 0.44)], margin: [0.03, 0.05, 0.03, 0.05] });
  text(s, para([
    "z = (ln 시험 − ln 예측) / √(τ² + σ_rep²/n): |z| ≤ 2 정합, z < −2 위험 과소평가(비보수적), z > 2 과대평가(보수적)",
    "P(k): TGA 사전분포에서 시험의 B 횟수가 나올 확률(≥ 0.05 정합) — 4종 모두 정합",
    "급발 A는 두 기작: 용융형(EPS, TGA로 예측) · 가연성 전면 착화(PIR(PUR), 콘·FT-IR로 예측해야 함)",
  ], { space: 5 }), { x: rx, y: 4.65, w: rw, h: 1.8, fontSize: 11 });
  addNotes(s, [
    "[스크립트]",
    "- 가로축이 TGA 예측, 세로축이 중공층 시험입니다. 점선 사이가 시험 한 번을 예측할 때의 95 % 구간입니다.",
    `- 게이트를 통과한 세 제품은 구간 안에 있습니다. 가연성 PIR(PUR)은 예측 ${fx(CONS("PIR_C").t600_tga_s)}초, 실측 ${fx(CONS("PIR_C").t600_obs_geo_s)}초로 z가 ${mn(CONS("PIR_C").z_spread, 1)}입니다(위험 과소평가). EPS 준불연은 용융형이라 TGA가 위험을 높게 봅니다(z ${mn(CONS("EPS_SC").z_spread)}, 보수적).`,
    "- 지속성(B)은 네 제품 모두 TGA 사전분포와 시험이 정합합니다.",
  ]);
}

// ================================================================================================ 13 Bayes fusion
{
  const s = content("결합 · 정합성", "베이즈 결합 — TGA 사전분포 × 중공층 시험",
    `TGA는 시험 2회 대비 무게 ${pct(FUS("PIR_SC").w_tga, 1)} — 통합 t₆₀₀은 시험값에서 TGA 쪽으로 조금 이동하고,\n재료 평균 t₆₀₀의 90 % 폭은 등급만 ${f2(W.ratio90_L0, 1)}배 → TGA ${f2(W.ratio90_L1, 2)}배 → 중규모 ${f2(W.ratio90_L2, 2)}배 → 통합 ${f2(W["ratio90_L1+L2"], 2)}배`);
  const lw = 7.4;
  text(s, para([
    [["확산  ", { bold: true }], ["ln t₆₀₀ ~ N(μ, σ_rep²) 시험별, 사전 μ ~ N(μ₀ = ln t̂_TGA, τ²) → μ_post = w·μ₀ + (1 − w)·ȳ, w = (1/τ²) / (1/τ² + n/σ_rep²)", {}]],
    [["지속  ", { bold: true }], ["p_B ~ Beta(a₀, b₀) (TGA 램프) → Beta(a₀ + k, b₀ + n − k) (시험 n회 중 B k회)", {}]],
  ], { space: 5 }), { x: MX, y: 1.8, w: lw, h: 0.95, fontSize: 12 });
  const head = [TH("제품", { fontSize: 10.5 }), TH("t₆₀₀ TGA", { fontSize: 10.5 }), TH("t₆₀₀ 시험", { fontSize: 10.5 }), TH("t₆₀₀ 통합", { fontSize: 10.5 }),
    TH("p_B TGA", { fontSize: 10.5 }), TH("B k/n", { fontSize: 10.5 }), TH("p_B 시험만", { fontSize: 10.5 }), TH("p_B 통합", { fontSize: 10.5 })];
  const rows = CHAR.map((m) => {
    const f = FUS(m);
    return [TD(LABEL[m], { bold: true, fontSize: 10.5 }), TD(fx(f.t600_tga_loo), { align: "center" }), TD(fx(f.t600_obs_geo), { align: "center" }),
      TD(fx(f.t600_post), { align: "center", bold: true, fill: { color: "E2F3E2" } }), TD(f2(f.pB_prior, 2), { align: "center" }), TD(`${f.B_k}/${f.B_n}`, { align: "center" }),
      TD(f2(f.pB_jeffreys, 2), { align: "center" }), TD(f2(f.pB_post, 2), { align: "center", bold: true, fill: { color: "E2F3E2" } })];
  });
  const eps = D.specs["L1+L2"].EPS_SC;
  rows.push([TD(LABEL.EPS_SC, { bold: true, fontSize: 10.5 }), TD("(용융형)", { align: "center", color: HEX.ink2 }), TD(fx(CONS("EPS_SC").t600_obs_geo_s), { align: "center" }),
    TD("시험값 사용", { align: "center", color: HEX.ink2 }), TD(f2(PROF("EPS_SC").pB_tga, 2), { align: "center" }), TD("0/3", { align: "center" }),
    TD(f2(0.5 / 4, 2), { align: "center" }), TD(f2(eps.a / (eps.a + eps.b), 2), { align: "center", bold: true, fill: { color: "E2F3E2" } })]);
  table(s, [head, ...rows], { x: MX, y: 2.85, w: lw, colW: [1.75, 0.8, 0.8, 0.9, 0.75, 0.65, 0.85, 0.9], fontSize: 11, border: BORDER, rowH: [0.34, ...rows.map(() => 0.42)], margin: [0.03, 0.05, 0.03, 0.05] });
  text(s, para([
    `σ_rep = ${f2(ST.s_rep, 3)} (시험 반복, 합동) · τ = ${f2(ST.tau, 3)} (TGA 예측, 제품 제외) · 준불연 제품 간 SD = ${f2(ST.sd_between, 3)} (등급만 아는 경우)`,
    "t₆₀₀ 단위 s · TGA 예측은 해당 제품을 빼고 맞춘 상수 · EPS는 확산 예측 대신 지속(p_B)만 TGA가 기여",
  ], { space: 3 }), { x: MX, y: 5.1, w: lw, h: 0.75, fontSize: 10.5, color: C.text2 });
  // right: uncertainty bars
  const rx = MX + lw + 0.3, rw = CW - lw - 0.3;
  s.addChart(pres.ChartType.bar, [{ name: "90 % 폭", labels: ["등급만 (L0)", "TGA (L1)", "중규모 (L2)", "통합 (L1+L2)"], values: [W.ratio90_L0, W.ratio90_L1, W.ratio90_L2, W["ratio90_L1+L2"]] }], {
    x: rx, y: 1.85, w: rw, h: 3.9, barDir: "col", chartColors: ["4B5563"], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.00",
    dataLabelFontSize: 11, dataLabelColor: HEX.ink, catAxisLabelColor: HEX.ink2, valAxisLabelColor: HEX.ink2, catAxisLabelFontSize: 10.5, valAxisLabelFontSize: 10,
    valGridLine: { color: HEX.grid, size: 0.75 }, catGridLine: { style: "none" }, valAxisMinVal: 1, valAxisMaxVal: 6, valAxisMajorUnit: 1,
    showTitle: true, title: "재료 평균 t₆₀₀의 90 % 구간 (상한/하한 비)", titleFontSize: 12, titleColor: HEX.ink, showLegend: false,
    catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", titleFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", barGapWidthPct: 60,
  });
  text(s, "같은 값이 세 제품에 공통(σ_rep · τ 공유). 건물 위험도의 90 % 구간은 s · 빈도 때문에 여전히 10–30배 → 17–18쪽", { x: rx, y: 5.85, w: rw, h: 0.6, fontSize: 10.5, color: C.text2 });
  NOTE(s, "엑셀 7_통합(베이즈)에 같은 계산이 수식으로 들어 있고, 13_검증에서 파이썬 값과 일치(최대 상대차 10⁻¹⁰ 수준)함을 확인합니다.", 6.5);
  addNotes(s, [
    "[스크립트]",
    "- TGA는 사전분포, 시험은 자료입니다. 시험 두 번의 정밀도가 높아 TGA 무게는 13.5 %지만, 시험이 없는 제품에서는 TGA가 유일한 근거입니다.",
    "- 오른쪽 막대: 재료에 대한 지식이 늘수록 t600 불확실성이 5.3배에서 1.2배 폭까지 줄어듭니다.",
  ]);
}

// ================================================================================================ 14 risk by evidence
pres.addSection({ title: "척도별 위험도" });
{
  const s = content("척도별 위험도", "척도(증거 수준)별 건물 위험도 — 같은 건물, 재료 지식만 다름",
    "등급만으로는 준불연 4종이 구분되지 않음 → TGA를 더하면 PF 내부용이 가장 위험,\n중규모 시험이 확정(불충족) — 통합(L1+L2)은 중규모와 비슷하고 구간이 약간 좁음");
  const modes = ["L0", "L1", "L2", "L1+L2"];
  // L0/L1 (no test) drawn light, L2/L1+L2 solid
  const mk = { L0: { symbol: "circle", size: 8, light: true }, L1: { symbol: "square", size: 8, light: true }, L2: { symbol: "circle", size: 9 }, "L1+L2": { symbol: "diamond", size: 10 } };
  const ch = { x: MX, y: 1.8, w: 8.6, h: 4.55, L: { x: 0.1, y: 0.04, w: 0.88, h: 0.8 }, xMin: -0.5, xMax: ORDER.length - 0.5, xMajor: 1, yMin: 1e-6, yMax: 1e-1, yLog: true };
  const xs = [], series = [];
  const npts = ORDER.length * modes.length * 3;
  ORDER.forEach((m, i) => {
    modes.forEach((mode, k) => {
      const r = RISK(mode, m);
      const x = i + (k - 1.5) * 0.17, base = xs.length;
      xs.push(x, x, x);
      const ys = (pairs) => { const a = new Array(npts).fill(null); pairs.forEach(([j, v]) => (a[base + j] = v)); return a; };
      series.push({ name: `${m} ${mode} 90%`, ys: ys([[0, r.R_S_p05], [1, r.R_S_p95]]), color: MAT[m], style: { noMarker: true } });
      const M = mk[mode];
      series.push({ name: `${m} ${mode} 평균`, ys: ys([[2, r.R_S_mean]]), color: MAT[m],
        style: { noLine: true, marker: { symbol: M.symbol, size: M.size, fill: M.light ? tint(MAT[m], 0.35) : MAT[m], line: MAT[m], lineW: M.light ? 19050 : 12700 } } });
    });
  });
  curves(s, ch, xs, series, { lineSize: 1.75, valAxisLogScaleBase: 10, valAxisLabelFormatCode: "0E+0", catAxisHidden: true, valAxisMajorUnit: undefined });
  const mp = plotMap(ch);
  ORDER.forEach((m, i) => text(s, SHORT[m], { x: mp.X(i) - 0.65, y: mp.py + mp.ph + 0.06, w: 1.3, h: 0.5, fontSize: 10.5, align: "center", color: HEX.ink }));
  vlabel(s, "R_S (층/동·년, 로그; 1E-3 = 10⁻³)", ch.x + 0.15, mp.py + mp.ph / 2, 3.2);
  const ly = 6.52;
  const legendItems = [["circle", true, "등급만 (L0)"], ["square", true, "소형시료 TGA (L1)"], ["circle", false, "중규모 (L2)"], ["diamond", false, "통합 (L1+L2)"]];
  legendItems.forEach(([sym, light, lab], i) => {
    const x = ch.x + 0.9 + i * 1.95;
    const shp = sym === "square" ? pres.ShapeType.rect : sym === "diamond" ? pres.ShapeType.diamond : pres.ShapeType.ellipse;
    s.addShape(shp, { x, y: ly + 0.06, w: 0.16, h: 0.16, fill: { color: light ? tint(HEX.ink2, 0.35) : HEX.ink2 }, line: { color: HEX.ink2, width: 1.5 } });
    text(s, lab, { x: x + 0.22, y: ly, w: 1.7, h: 0.28, fontSize: 11, color: C.text2, valign: "middle" });
  });
  const rx = MX + 8.85, rw = CW - 8.85;
  const head = [TH("재료", { fontSize: 10 }), TH("L1u TGA 단독", { fontSize: 10 }), TH("L2", { fontSize: 10 })];
  const r1 = RISK("L1u", "PIR_C"), r2 = RISK("L2", "PIR_C");
  table(s, [head, [TD(LABEL.PIR_C, { bold: true, fontSize: 10 }), TD(supStr(r1.R_S_mean), { align: "center", fontSize: 10, fill: { color: "F8DEDE" } }), TD(supStr(r2.R_S_mean), { align: "center", fontSize: 10 })],
    [TD("판정", { bold: true, fontSize: 10 }), TD(r1.verdict, { align: "center", fontSize: 10 }), TD(r2.verdict, { align: "center", fontSize: 10 })]],
  { x: rx, y: 1.85, w: rw, colW: [1.25, 1.1, rw - 2.35], fontSize: 10, border: BORDER, rowH: [0.42, 0.42, 0.36], margin: [0.03, 0.04, 0.03, 0.04] });
  text(s, para([
    `게이트 없는 TGA(L1u)는 가연성 PIR(PUR)의 위험을 ${f2(r2.R_S_mean / r1.R_S_mean, 1)}배 낮게 보고 판정도 '판단 유보'로 틀림`,
    `L1 순위: PF 내부용 ${supStr(RISK("L1", "PF_in").R_S_mean)} > 다른 준불연 ≈ ${supStr(RISK("L1", "PIR_SC").R_S_mean)}`,
    "구간 폭은 모든 수준에서 10–30배 → 재료보다 s·빈도가 지배",
  ], { space: 6 }), { x: rx, y: 3.25, w: rw, h: 3.1, fontSize: 11 });
  addNotes(s, [
    "[스크립트]",
    "- 재료마다 네 개의 점이 증거 수준입니다. 연한 원은 등급만, 연한 네모는 TGA를 더한 것, 진한 원은 중규모 시험, 마름모는 둘을 결합한 것입니다. 세로선은 90 % 구간입니다.",
    "- 등급만으로는 준불연 네 제품이 구분되지 않습니다. TGA만 더해도 PF 내부용이 가장 위험하다는 순위가 나오고, 중규모 시험이 확인합니다.",
    "- 오른쪽 표: 게이트 없이 TGA만 쓰면 가연성 PIR(PUR)을 6배 낮게 평가해 판정까지 바뀝니다.",
  ]);
}

// ================================================================================================ 15 verdict grid
{
  const s = content("척도별 위험도", "판정이 증거 수준에 따라 어떻게 바뀌나 (예시 기준 F(n ≥ 2) ≤ 1×10⁻⁴/동·년, 신뢰 90 %)",
    "대책 없음(M0): 증거가 쌓이면 PF 내부용은 '불충족' 확정, 다른 준불연은 '판단 유보' —\n차단재 + 배리어(M1): 모든 증거 수준에서 준불연 '충족', 가연성 '불충족'");
  const modes = ["L0", "L1", "L1u", "L2", "L1+L2"];
  const mlab = { L0: "등급만", L1: "소형(TGA)", L1u: "TGA 단독", L2: "중규모", "L1+L2": "통합" };
  const pw = (CW - 0.35) / 2;
  [["M0", "M0 대책 없음"], ["M1", "M1 차단재 + 층간 배리어"]].forEach(([g, gt], gi) => {
    const x0 = MX + gi * (pw + 0.35), y0 = 1.85, lw = 1.65, cw = (pw - lw) / modes.length, rh = 0.62;
    text(s, gt, { x: x0, y: y0, w: pw, h: 0.34, fontSize: 13.5, bold: true });
    modes.forEach((mode, j) => text(s, mlab[mode], { x: x0 + lw + j * cw, y: y0 + 0.4, w: cw, h: 0.3, fontSize: 11, bold: true, align: "center", color: mode === "L1u" ? C.text2 : C.text1 }));
    MATS.forEach((m, i) => {
      const y = y0 + 0.78 + i * rh;
      s.addShape(pres.ShapeType.rect, { x: x0, y: y + 0.12, w: 0.08, h: 0.32, fill: { color: MAT[m] }, line: { color: MAT[m], width: 0.5 } });
      text(s, LABEL[m], { x: x0 + 0.15, y, w: lw - 0.15, h: 0.56, fontSize: 11, bold: true, valign: "middle" });
      modes.forEach((mode, j) => {
        // a text-only cell: the dotted chip is too wide for "판단 유보" at this column width
        const v = RISK(mode, m, g).verdict, st = STATUS[VERDICT[v] || "poor"];
        box(s, x0 + lw + j * cw + 0.04, y + 0.11, cw - 0.08, 0.36, { fill: st.tint, line: st.color, lw: 1, r: 0.06, name: `verdict ${g} ${m} ${mode}` });
        text(s, v, { x: x0 + lw + j * cw + 0.04, y: y + 0.11, w: cw - 0.08, h: 0.36, fontSize: 10.5, bold: true, align: "center", valign: "middle" });
      });
    });
  });
  text(s, para([
    "충족 비율(기준을 만족하는 인식론적 표본 비율) ≥ 90 % → 충족, ≤ 10 % → 불충족, 그 사이 → 판단 유보 · 판단 유보 = 이 증거로는 결론 불가 → 다음 시험",
    "TGA 단독(회색 열)은 게이트 없이 쓴 경우 — 가연성 PIR(PUR)을 '판단 유보'로 잘못 봄",
  ], { space: 4 }), { x: MX, y: 5.95, w: CW, h: 0.9, fontSize: 11, color: C.text2 });
  addNotes(s, [
    "[스크립트]",
    "- 칸마다 같은 기준으로 판정했습니다. 증거 수준이 올라갈수록 판정이 확정되는 방향을 볼 수 있습니다.",
    "- 대책이 없는 경우 PF 내부용은 중규모 시험에서 불충족으로 확정되고, 다른 준불연 세 제품은 끝까지 판단 유보입니다. 판단을 막는 것은 재료가 아니라 s입니다(18쪽).",
  ]);
}

// ================================================================================================ 16 material factor
{
  const s = content("척도별 위험도", "재료 인자 MF와 그 구성 — 확산 몫과 자립 연소(지속) 몫",
    `재료 인자 MF: 가연성 PIR(PUR) ${fx(RISK("L1+L2", "PIR_C").MF_p50)}배, PF 내부용 ${fx(RISK("L1+L2", "PF_in").MF_p50)}배, 나머지 ${fx(RISK("L1+L2", "EPS_SC").MF_p50)}–${fx(RISK("L1+L2", "PF_out").MF_p50)}배 —\nPF 외부용은 위험 증가분의 ${pct(DEC("PF_out").persist_part)}가 자립 연소(f₄₀₀,air 높음 · B 2/2)에서 옴`);
  const lw = 6.2;
  const modes = ["L0", "L1", "L2", "L1+L2"];
  const head = [TH("재료"), ...modes.map((m) => TH({ L0: "등급만", L1: "TGA", L2: "중규모", "L1+L2": "통합" }[m]))];
  const rows = MATS.map((m) => [TD(LABEL[m], { bold: true }), ...modes.map((mode) => TD(fx(RISK(mode, m).MF_p50), { align: "center", bold: mode === "L1+L2" }))]);
  table(s, [head, ...rows], { x: MX, y: 1.85, w: lw, colW: [2.1, 1.025, 1.025, 1.025, 1.025], fontSize: 11.5, border: BORDER, rowH: [0.36, ...MATS.map(() => 0.44)], margin: [0.03, 0.06, 0.03, 0.06] });
  text(s, "MF 중앙값 (대책 없음, 스프링클러 설치) — 불연 단열재 = 1", { x: MX, y: 4.5, w: lw, h: 0.28, fontSize: 10.5, color: C.text2 });
  text(s, para([
    "확산 몫 = (R_S(p_B = 0) − R_S,불연) / (R_S − R_S,불연), 지속 몫 = 나머지",
    "점추정에서는 진압(15분)이 분출 종료(20분)보다 먼저라 지속이 드러나지 않음 → 몬테카를로 평균으로 계산(진압 지연 사건)",
    "TGA의 f₄₀₀,air가 이 지속 몫을 예측하는 재료 인자(T2)",
  ], { space: 4 }), { x: MX, y: 4.85, w: lw, h: 1.5, fontSize: 11 });
  const rx = MX + lw + 0.3, rw = CW - lw - 0.3;
  const mats = MATS;
  s.addChart(pres.ChartType.bar, [
    { name: "확산 몫", labels: mats.map((m) => LABEL[m]), values: mats.map((m) => DEC(m).spread_part * 100) },
    { name: "지속 몫", labels: mats.map((m) => LABEL[m]), values: mats.map((m) => DEC(m).persist_part * 100) },
  ], {
    // shares already sum to 100: plain stacked bars on a 0–100 axis; labels under 5 % are blanked by the number format
    x: rx, y: 1.85, w: rw, h: 4.45, barDir: "bar", barGrouping: "stacked", chartColors: ["9CA3AF", "EB6834"],
    showValue: true, dataLabelPosition: "ctr", dataLabelFormatCode: '[<5]"";0', dataLabelFontSize: 10.5, dataLabelColor: "FFFFFF",
    valAxisMinVal: 0, valAxisMaxVal: 100, valAxisMajorUnit: 20, valAxisLabelFormatCode: "0",
    catAxisLabelColor: HEX.ink, valAxisLabelColor: HEX.ink2, catAxisLabelFontSize: 10.5, valAxisLabelFontSize: 10, valGridLine: { color: HEX.grid, size: 0.75 },
    catGridLine: { style: "none" }, showLegend: true, legendPos: "b", legendFontSize: 10.5, legendColor: HEX.ink2, showTitle: true,
    title: "위험 증가분의 구성 (%, 통합 L1+L2, 대책 없음)", titleFontSize: 12, titleColor: HEX.ink, catAxisOrientation: "maxMin",
    catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", titleFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", barGapWidthPct: 50,
  });
  addNotes(s, [
    "[스크립트]",
    "- 재료 인자는 '이 재료 때문에 외벽 위험이 몇 배가 되는가'입니다. 증거 수준이 달라도 순위는 같습니다.",
    "- 오른쪽은 위험 증가분을 확산과 지속으로 나눈 것입니다. PF 외부용은 확산은 느리지만 소화 후에도 계속 타서, 위험의 절반 이상이 지속에서 옵니다. 이것이 TGA의 고온 산화 몫(f400)이 재료 인자로 의미 있는 이유입니다.",
  ]);
}

// ================================================================================================ 17 L3 protocol
pres.addSection({ title: "L3 시뮬레이션" });
{
  const s = content("L3 시뮬레이션", "L3 시뮬레이션 평가 — 중규모 재현으로 검증한 뒤 건물 외벽으로 외삽 (엑셀 8_L3_시뮬레이션 양식)",
    "FDS는 ① 중공층 시험 재현으로 검증한 뒤 ② 건물 외벽 3–6개 층을 계산해 s를 구함 —\n건물 위험도의 가장 큰 불확실성(s)을 직접 줄이는 단계");
  const steps = [
    ["1", "입력 준비", ["반응: TGA → 반응 온도 · 속도 (PyroSim 자료)", "열물성: 실측 ρ · c_p · k (콘 TRP로 확인)", "연소열 · 수율: 콘 EHC, FT-IR · 콘 CO · HCN"]],
    ["2", "중규모 재현 · 검증", ["3 m 중공층 · 30 kW 버너 · 열전대 500–3000 mm", "격자 민감도 (중공층 안 1–2 cm)", "기준(예시): t₆₀₀ ±30 %, 최고온도 ±100 K"]],
    ["3", "건물 외벽 시나리오", ["IE1 창 분출화염(최성기 실내 화재)", "IE2 외벽 하부 외부 화원", "재료 5종 × 2, 배리어 유무"]],
    ["4", "위험도로 연결", ["층별 도달시각 T_j → s_j = ΔT_j / t₆₀₀", "s 추정 → 판정 경계와 비교(18쪽)", "창 열류속 · 위층 FED → 확장 항목"]],
  ];
  const gap = 0.3, w = (CW - 3 * gap) / 4, y = 1.85, h = 2.85;
  steps.forEach(([n, t, items], i) => {
    const x = MX + i * (w + gap);
    box(s, x, y, w, h, { fill: i === 1 ? "E2F3E2" : "F3F4F6", line: i === 1 ? STATUS.good.color : "F3F4F6", lw: i === 1 ? 1.5 : 0.5, r: 0.08, name: `step ${n}` });
    badge(s, n, x + 0.15, y + 0.15);
    text(s, t, { x: x + 0.55, y: y + 0.1, w: w - 0.65, h: 0.4, fontSize: 14, bold: true, valign: "middle" });
    text(s, para(items, { space: 5 }), { x: x + 0.18, y: y + 0.65, w: w - 0.32, h: h - 0.75, fontSize: 11.5 });
    if (i < 3) arrow(s, x + w + 0.02, y + h / 2, x + w + gap - 0.02, y + h / 2, { width: 1.75 });
  });
  const cw = (CW - 0.25) / 2;
  card(s, MX, 4.95, cw, 1.6, "fair", "시뮬레이션이 줄이는 것", [
    "s: 현재 가정 LogU(0.25, 1) — 건물 위험도 민감도 1위",
    "층마다 t₆₀₀을 반복하는가, 연속 전선인가(모델 구조)",
  ], { size: 12, headSize: 13.5 });
  card(s, MX + cw + 0.25, 4.95, cw, 1.6, "poor", "시뮬레이션의 한계(문서화 필요)", [
    "용융·적하(EPS), 면재 박리·낙하는 FDS로 재현 어려움",
    "기상 난연 효과는 반응 입력으로 직접 표현 안 됨 → 콘 HRR로 보정",
  ], { size: 12, headSize: 13.5 });
  addNotes(s, [
    "[스크립트]",
    "- 시뮬레이션은 두 단계입니다. 먼저 중공층 시험을 재현해 모델을 검증하고, 검증된 모델로 실제 건물 외벽을 계산합니다.",
    "- 건물 계산에서 층마다 화염이 닿는 시각을 얻으면, 지금 가정으로 두고 있는 s를 실제 값으로 바꿀 수 있습니다.",
    "- 용융·적하처럼 FDS가 못 그리는 현상은 한계로 문서에 남기고, 그 부분은 중규모 시험 결과를 우선합니다.",
  ]);
}

// ================================================================================================ 18 VOI
{
  const s = content("L3 시뮬레이션", "L3의 정보가치 — s를 알면 판정이 어떻게 갈리는가 (통합 L1+L2, 대책 없음)",
    `판단 유보 3종은 s로 갈림 — 충족: s ≥ ${f2(THR("EPS_SC").s_meet)}(EPS) · ${f2(THR("PIR_SC").s_meet)}(PIR(PUR)) · ${f2(THR("PF_out").s_meet)}(PF 외부용)\nPF 내부용은 s ≤ ${f2(THR("PF_in").s_fail)}이면 불충족(현재 사전분포에서 ${pct(THR("PF_in").P_fail_after_L3)}) · 대책(M1)은 s와 무관하게 충족`);
  const sg = D.s_grid;
  const ch = { x: MX, y: 1.8, w: 7.4, h: 4.6, L: { x: 0.11, y: 0.05, w: 0.86, h: 0.8 }, xMin: 0.25, xMax: 1.0, xMajor: 0.25, yMin: 0, yMax: 1, yMajor: 0.2 };
  const series = MATS.map((m) => ({ name: `voi ${m}`, ys: sg.map((sv) => D.voi.find((v) => v.material === m && v.mitigation === "M0" && Math.abs(v.s - sv) < 1e-9).conf_meets_F2), color: MAT[m],
    style: { marker: { symbol: "circle", size: 6, fill: MAT[m], line: MAT[m] } } }));
  curves(s, ch, sg, series, { lineSize: 2.25, valAxisLabelFormatCode: "0%", catAxisLabelFormatCode: "0.00" });
  const mp = plotMap(ch);
  [[D.constants.crit_conf, "충족 경계"], [1 - D.constants.crit_conf, "불충족 경계"]].forEach(([v, lab]) => {
    s.addShape(pres.ShapeType.line, { x: mp.px, y: mp.Y(v), w: mp.pw, h: 0, line: { color: HEX.ink2, width: 1, dashType: "dash" } });
    text(s, lab, { x: mp.px + 0.05, y: mp.Y(v) - (v > 0.5 ? 0.3 : -0.02), w: 1.3, h: 0.26, fontSize: 9.5, color: C.text2 });
  });
  text(s, "실물/시험 층당 확산시간 비 s (L3 · 실대형으로 확정할 값)", { x: mp.px, y: mp.py + mp.ph + 0.32, w: mp.pw, h: 0.28, fontSize: 11, color: C.text2, align: "center" });
  vlabel(s, "기준 충족 비율", ch.x + 0.15, mp.py + mp.ph / 2, 2.2);
  const SHORT1 = { EPS_SC: "EPS 준불연", PIR_C: "PIR 가연성", PIR_SC: "PIR 준불연", PF_out: "PF 외부용", PF_in: "PF 내부용" };
  legend(s, MATS.map((m) => ({ label: SHORT1[m], color: MAT[m], w: 0.9 })), ch.x + 0.1, 6.55);
  const rx = MX + 7.65, rw = CW - 7.65;
  const head = [TH("재료", { fontSize: 10.5 }), TH("s_meet", { fontSize: 10.5 }), TH("s_fail", { fontSize: 10.5 }), TH("L3로 확정될 확률", { fontSize: 10.5 })];
  const rows = MATS.map((m) => {
    const t = THR(m);
    return [TD(LABEL[m], { bold: true, fontSize: 10.5 }), TD(t.s_meet === null ? "—" : f2(t.s_meet), { align: "center" }), TD(t.s_fail === null ? "—" : f2(t.s_fail), { align: "center" }),
      TD(pct(t.P_decided_after_L3), { align: "center", bold: true, fill: { color: "E2F3E2" } })];
  });
  table(s, [head, ...rows], { x: rx, y: 1.85, w: rw, colW: [1.6, 0.85, 0.85, rw - 3.3], fontSize: 11, border: BORDER, rowH: [0.36, ...MATS.map(() => 0.42)], margin: [0.03, 0.05, 0.03, 0.05] });
  text(s, para([
    "s ≥ s_meet → 충족, s ≤ s_fail → 불충족, 사이 → 유보",
    "확정될 확률 = 현재 s 사전분포 LogU(0.25, 1)에서 s가 두 경계 밖에 있을 확률",
    "판단 유보 3종의 결론은 재료 시험보다 s(시뮬레이션·실대형)에 달려 있음",
  ], { space: 5 }), { x: rx, y: 4.55, w: rw, h: 1.85, fontSize: 11 });
  addNotes(s, [
    "[스크립트]",
    "- 가로축 s를 하나로 정했을 때 기준을 만족하는 비율입니다. 점선 위면 충족, 아래면 불충족입니다.",
    `- 판단 유보인 세 준불연 제품은 s가 ${f2(Math.min(...["EPS_SC", "PIR_SC", "PF_out"].map((m) => THR(m).s_meet)))}–${f2(Math.max(...["EPS_SC", "PIR_SC", "PF_out"].map((m) => THR(m).s_meet)))} 이상이면 충족, ${f2(Math.min(...["EPS_SC", "PIR_SC", "PF_out"].map((m) => THR(m).s_fail)))}–${f2(Math.max(...["EPS_SC", "PIR_SC", "PF_out"].map((m) => THR(m).s_fail)))} 이하이면 불충족입니다. 즉 이 제품들의 결론은 재료 시험을 더 해서가 아니라 s를 정해서 나옵니다.`,
    "- 시뮬레이션의 가치를 숫자로 보인 것입니다. 대책을 넣으면 s와 무관하게 충족이라 시뮬레이션이 필요 없습니다.",
  ]);
}

// ================================================================================================ 19 evidence levels + new product flow
pres.addSection({ title: "적용" });
{
  const s = content("적용", "증거 수준과 판정 규칙 — 새 제품(예: EPS 가연성)은 어떻게 평가하나",
    "판정은 가장 높은 증거 수준으로 내리고 수준을 함께 표기(예: '충족 · E2') —\n새 제품은 게이트부터 확인 → 소형시료로 끝낼지, 중규모가 필요한지 결정");
  const lw = 6.3;
  const rows = [
    ["E0", "등급(인증서)만", "준불연 모집단 분포 · 가연성 DB", "우선순위 선별"],
    ["E1", "+ 소형시료(콘·TGA·FT-IR)", "게이트 통과 시 전이함수 예측", "준불연 탄화형의 상대 비교"],
    ["E2", "+ 중규모 시험", "시험 t₆₀₀ · B, TGA와 결합", "재료 판정(현재 최고 수준)"],
    ["E3", "+ 시뮬레이션 · 실대형", "s 확정, 층별 시간", "건물 판정"],
  ];
  table(s, [[TH("수준"), TH("가진 자료"), TH("재료 입력"), TH("쓸 수 있는 판단")], ...rows.map(([a, b, c, d]) => [TD(a, { bold: true, align: "center" }), TD(b, { bold: true }), TD(c), TD(d)])],
    { x: MX, y: 1.85, w: lw, colW: [0.6, 1.95, 2.05, 1.7], fontSize: 11, border: BORDER, rowH: [0.34, 0.55, 0.55, 0.55, 0.55], margin: [0.04, 0.07, 0.04, 0.07] });
  text(s, para([
    "가연성 등급은 E1으로 끝낼 수 없음(게이트) → E2 필수",
    "정합성 불일치(|z| > 2) → 낮은 수준 결과를 버리고 높은 수준을 따름 + 원인 기록",
    "판단 유보 → 민감도·정보가치가 가리키는 다음 시험",
  ], { space: 5 }), { x: MX, y: 4.5, w: lw, h: 1.6, fontSize: 11.5 });
  const rx = MX + lw + 0.3, rw = CW - lw - 0.3;
  text(s, "예: EPS 가연성 (TGA만 있고 중공층 기록 4분)", { x: rx, y: 1.85, w: rw, h: 0.34, fontSize: 13.5, bold: true });
  const flow = [
    ["poor", "게이트(등급)", "가연성 → TGA 전이함수 미적용"],
    ["poor", "TGA 자료 판정", "공기 자료 '재측정' — 쓰기 전 재측정"],
    ["fair", "재료형", `용융형(f₄₀₀,air ${f2(PROF("EPS_C").f400_air, 3)}) → 급발 A 예상`],
    ["conditional", "콘 · FT-IR", "준불연 기준·난연제 확인(등급 재확인)"],
    ["good", "중규모 재시험", "30분 평가 완주 → t₆₀₀ · B → E2 판정"],
  ];
  flow.forEach(([st, a, b], i) => {
    const y = 2.3 + i * 0.78;
    box(s, rx, y, rw, 0.66, { fill: STATUS[st].tint, line: STATUS[st].tint, r: 0.06, name: a });
    badge(s, String(i + 1), rx + 0.12, y + 0.18);
    text(s, a, { x: rx + 0.55, y: y + 0.04, w: 1.75, h: 0.58, fontSize: 12, bold: true, valign: "middle" });
    text(s, b, { x: rx + 2.3, y: y + 0.04, w: rw - 2.4, h: 0.58, fontSize: 11.5, valign: "middle" });
    if (i < flow.length - 1) arrow(s, rx + 0.27, y + 0.66, rx + 0.27, y + 0.78, { width: 1 });
  });
  NOTE(s, "증거 수준과 판정은 NFPA 551 7.3.3.4(정량 결과 문서)의 '모델·자료 평가'와 '한계' 항목에 그대로 기록합니다(facade_fra 서술양식).", 6.5, { h: 0.4 });
  addNotes(s, [
    "[스크립트]",
    "- 판정에는 항상 증거 수준을 붙입니다. 같은 '충족'이라도 E1과 E2는 무게가 다릅니다.",
    "- 오른쪽은 실제로 남아 있는 사례입니다. EPS 가연성은 TGA만 있는데 가연성이라 게이트에서 막히고, TGA 공기 자료도 재측정 대상이며, 중공층 기록도 4분뿐입니다. 그래서 체계는 '중규모 재시험이 필요하다'고 답합니다.",
  ]);
}

// ================================================================================================ 20 test plan
{
  const s = content("적용", "시험 · 해석 계획 — 판정을 바꿀 수 있는 것부터",
    "1순위 = 게이트(콘 · FT-IR) · 지속 표본(중규모 반복) · 대책 근거(배리어) · s(FDS) —\nTGA 보완과 발생가스 분석은 2순위");
  const rows = [
    ["L1", "콘칼로리미터", "50 kW/m² 10분 3회 + 25·35·65 kW/m² 착화", "D1·D2·D3·G", "가연/준불연 게이트, 착화·발열", "1"],
    ["L1", "FT-IR (ATR)", "심재·면재 3 위치, 4 cm⁻¹, 32회", "G", "제품 식별·난연제 → 전이함수 적용", "1"],
    ["L1", "TGA–FTIR", "N₂·공기 5 K/min, 전송관 200–250 °C", "D6", "유해가스 → FDS 수율·FED", "2"],
    ["L1", "TGA–DSC 보완", "초기 질량, 블랭크·사파이어, EPS 가연성 재측정", "D3·D4", "잔류량·c_p·반응열, 신뢰도 개선", "2"],
    ["L2", "중공층 시험 보완", "소화 후 60분 관찰, CO/CO₂·열류속, 반복 3회", "D4·D5·D6", "p_B 표본 확대(현재 2–3회)", "1"],
    ["L2", "EPS 가연성 재시험", "30분 평가 완주", "D2·D4", "정량평가 제외 해소", "1"],
    ["L2", "배리어·차단재 포함 시험", "층간 배리어 위·아래 열전대", "—", "M1 효과 p_cb · p_e1", "1"],
    ["L3", "FDS 중규모 재현 → 건물", "격자 민감도, 3–6개 층, IE1·IE2", "D2·D3·D6", "s (건물 위험도 1위 불확실성)", "1"],
    ["L3+", "실대형 KS F 8414", "2–3종", "D2", "s 직접 보정", "2"],
  ];
  table(s, [[TH("척도"), TH("시험 · 해석"), TH("조건"), TH("차원"), TH("줄이는 불확실성"), TH("순위")],
    ...rows.map(([a, b, c, d, e, f]) => [TD(a, { align: "center", bold: true }), TD(b, { bold: true }), TD(c), TD(d, { align: "center" }), TD(e),
      TD(f, { align: "center", bold: true, fill: { color: f === "1" ? "E2F3E2" : "F3F4F6" } })])],
  { x: MX, y: 1.8, w: CW, colW: [0.7, 2.4, 3.9, 1.2, 3.33, 0.803], fontSize: 11, border: BORDER, rowH: [0.34, ...rows.map(() => 0.47)], margin: [0.03, 0.07, 0.03, 0.07] });
  NOTE(s, "규격·시편·반복 세부는 엑셀 12_시험계획. 각 시험 결과는 해당 시트(2·3·4·5·8)에 넣으면 전이함수·결합·판정이 다시 계산됩니다(MC 위험도는 multiscale.py 재실행).", 6.45, { h: 0.45 });
  addNotes(s, [
    "[스크립트]",
    "- 순위는 판정을 바꿀 수 있는지로 정했습니다. 게이트, 지속 표본, 대책 근거, s가 1순위입니다.",
  ]);
}

// ================================================================================================ 21 limitations
{
  const s = content("적용", "한계와 해석 주의",
    "전이함수는 준불연 탄화형 3종(시험 6회)으로 보정한 잠정식이고, 건물 쪽 입력은 예시값 —\n체계와 절차는 확정, 숫자는 자료가 늘수록 갱신");
  const cw = (CW - 0.25) / 2;
  card(s, MX, 1.85, cw, 2.9, "poor", "한계", [
    "T1·T2는 제품 3종·시험 6회로 보정 → τ·램프 경계는 잠정(시험 추가 시 재계산)",
    "콘·FT-IR·FDS 자료 없음 → T4–T8은 제안식, L3 위험도 미산출",
    "TGA–DSC: 초기 질량·DSC 보정 없음 → 발열 절대량·반응열 미사용",
    "건물 모델: 빈도·분기확률·시간은 예시값, s 가정, 분기 독립",
    "가연성 c 등급 DB는 1종(PIR(PUR)) → 새 가연성 제품은 중규모 필수",
  ], { size: 12 });
  card(s, MX + cw + 0.25, 1.85, cw, 2.9, "good", "해석 주의 · 갱신 방법", [
    "판정에는 증거 수준(E0–E3)과 기준·신뢰수준(예시)을 함께 표기",
    "|z| > 2 불일치는 오류가 아니라 '전이함수 적용 범위 밖' 신호 → 게이트 재검토",
    "시험이 늘면: 엑셀에 입력 → σ_rep·τ·결합 자동 갱신, multiscale.py 재실행으로 위험도 갱신",
  ], { size: 12 });
  // where this framework sits next to the two references the user supplied
  text(s, "세 틀의 역할 — 겹치지 않고 이어짐", { x: MX, y: 4.95, w: CW, h: 0.32, fontSize: 13.5, bold: true });
  const tw = (CW - 2 * 0.25) / 3;
  [["NFPA 551 (2022)", "절차 · 문서", "방법을 정하지 않는 가이드: 목표 → 방법 선택 → 문서(7.3.3.4) → 검토(8장). 이 체계의 결과를 담는 그릇(facade_fra 서술양식)"],
    ["EFFECT (2018)", "포트폴리오 선별 · 정성", "기존 고층 건물을 A–E로 줄 세움, 재료는 범주. 어느 건물을 먼저 볼지 정하는 1차 선별"],
    ["이 체계", "재료 · 제품 판정 · 정량", "소형시료 · 중규모 · 시뮬레이션 → t₆₀₀ · p_B → R = Σ F·C, 증거 수준 표기. EFFECT가 남긴 '재료 조합 점수'를 시험으로 채움"],
  ].forEach(([h1, h2, b], i) => {
    const x = MX + i * (tw + 0.25);
    box(s, x, 5.35, tw, 1.5, { fill: i === 2 ? "E2F3E2" : "F3F4F6", line: i === 2 ? STATUS.good.color : "F3F4F6", lw: i === 2 ? 1.25 : 0.5, r: 0.08, name: h1 });
    text(s, [{ text: h1, options: { bold: true } }, { text: `  ${h2}`, options: { color: C.text2 } }], { x: x + 0.18, y: 5.42, w: tw - 0.36, h: 0.32, fontSize: 12.5 });
    text(s, b, { x: x + 0.18, y: 5.78, w: tw - 0.36, h: 1.0, fontSize: 11 });
  });
  addNotes(s, [
    "[스크립트]",
    "- 체계 자체(척도, 공통 단위, 결합·검사·게이트)는 자료와 무관하게 쓸 수 있고, 숫자는 시험이 늘수록 바뀝니다.",
    "- 아래 세 칸: NFPA 551은 절차와 문서의 틀, EFFECT는 건물 포트폴리오의 1차 선별, 이 체계는 재료·제품을 시험 근거로 정량 판정하는 부분입니다. 서로 대체가 아니라 이어지는 관계입니다.",
  ]);
}

// ================================================================================================ 22 appendix
pres.addSection({ title: "부록" });
{
  const s = content("부록", "부록 — 식, 상수, 재현", "모든 수치는 multiscale.py로 재현되며, 결합·정합성 계산은 엑셀 수식과 일치함");
  const cw = (CW - 0.25) / 2;
  text(s, para([
    [["T1 ", { bold: true }], ["t̂₆₀₀ = t_ig,sys + C_sys / r_max,air", {}]],
    [["T2 ", { bold: true }], ["m = min(0.95, max(0.05, (f₄₀₀,air − 0.60)/0.10)), p_B ~ Beta(2m, 2(1 − m))", {}]],
    [["결합 ", { bold: true }], ["w = (1/τ²)/(1/τ² + n/σ_rep²), μ_post = wμ₀ + (1 − w)ȳ, σ_post = (1/τ² + n/σ_rep²)^(−1/2)", {}]],
    [["정합 ", { bold: true }], ["z = (ȳ − μ₀)/√(τ² + σ_rep²/n); P(k) = C(n,k)·B(k + a₀, n − k + b₀)/B(a₀, b₀)", {}]],
    [["건물 ", { bold: true }], ["τ_F = s·t₆₀₀·h/3 m, R_S = F₁[p_e C₁ + (1 − p_e)C_L] + F₂ p_e C₂ (facade_fra)", {}]],
    [["재료 인자 ", { bold: true }], ["MF = R_S/R_S,불연; 지속 몫 = (R_S − R_S(p_B = 0))/(R_S − R_S,불연)", {}]],
    [["상수 ", { bold: true }], [`t_ig,sys ${fx(D.constants.t_ig_sys, 1)} s, C_sys ${f2(D.constants.c_sys, 1)}, σ_rep ${f2(ST.s_rep, 4)}, τ ${f2(ST.tau, 4)}, 제품 간 SD ${f2(ST.sd_between, 4)}`, {}]],
  ], { space: 6 }), { x: MX, y: 1.85, w: cw, h: 4.6, fontSize: 11.5 });
  box(s, MX + cw + 0.25, 1.85, cw, 4.6, { fill: "F3F4F6", line: "F3F4F6", r: 0.08, name: "reproduce" });
  text(s, para([
    [["재현 (multiscale_fra/)", { bold: true }]],
    [["python multiscale.py", { fontFace: "Consolas" }], [` → results/, figures/, report/ms_data.json (MC: 인식론적 ${D.n_outer} × 우연적 ${D.n_inner})`, {}]],
    [["python report/build_ms_xlsx.py", { fontFace: "Consolas" }], [" → 다척도_위험성평가_계산서.xlsx (recalc 후 13_검증)", {}]],
    [["node report/build_ms_deck.js", { fontFace: "Consolas" }], [" → 이 덱", {}]],
    [["자료 출처", { bold: true }]],
    [["tga_dsc (TGA 지표·신뢰도), cavity_fire_index (t₆₀₀·P·S·R·A), coupled_analysis (t_ig·t_sp), integrated_model (C_sys·LOO·B), facade_fra (건물 모델)", {}]],
    [["Lamont & Ingolfsson (2018), High Rise Buildings with Combustible Exterior Wall Assemblies: Fire Risk Assessment Tool, NFPA/FPRF", {}]],
    [["L2 모드 = facade_fra/qfra.py와 같은 난수, 결과 일치 확인(상대차 ", {}], [supStr(D.l2_check_rel_err), {}], [")", {}]],
  ], { bullet: false, space: 6 }), { x: MX + cw + 0.45, y: 1.95, w: cw - 0.4, h: 4.4, fontSize: 11 });
  addNotes(s, ["[스크립트]", "- 식과 상수, 재현 방법입니다. 엑셀과 파이썬이 같은 값을 내는지 13_검증 시트에서 확인할 수 있습니다."]);
}

// pptxgenjs writes the X axis of a scatter chart (its first <c:valAx>) with the Y axis number format and has no log
// option for it: fix the X axis of the chart that holds the named series
const JSZip = require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));
const X_AXES = [{ series: "1:1", fmt: "0", log: true }, { series: "voi PF_in", fmt: "0.00", log: false }];
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
      if (ax.log) blk = blk.replace("<c:scaling>", '<c:scaling><c:logBase val="10"/>');
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
