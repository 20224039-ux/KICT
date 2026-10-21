#!/usr/bin/env node
// TGA·DSC 재료 특성 그래프와 PyroSim 입력값 덱 (pptxgenjs, native charts).
// Reads deck_data.json (curves, metrics, verdicts) and pyrosim_data.json (TGA stages -> FDS reaction inputs).
//
//   python ../pyrosim_inputs.py && python deck_data.py
//   NODE_PATH=<folder holding pptxgenjs> node build_pyrosim_deck.js [out.pptx]
"use strict";
const fs = require("fs");
const path = require("path");
const { makeDeck } = require("./deck_lib");

const D = JSON.parse(fs.readFileSync(path.join(__dirname, "deck_data.json"), "utf8"));
const P = JSON.parse(fs.readFileSync(path.join(__dirname, "pyrosim_data.json"), "utf8"));
const OUT = process.argv[2] || path.join(__dirname, "TGA-DSC_재료특성그래프_PyroSim입력.pptx");
const {
  pres, C, HEX, STATUS, MX, CW, LANG, fmt, tab, richText,
  content, text, vlabel, table, paras, badge, legend, chip, axisStyle, plotMap, curves, finish,
} = makeDeck({
  title: "TGA·DSC 재료 특성 그래프와 PyroSim 입력값",
  subject: "단열재 6종 TGA–DSC: 재료 특성 그래프, PyroSim(FDS) 재료 입력값과 밀도·비열을 구하는 방법",
  footer: "KICT 중공층 외단열 · TGA·DSC 재료 특성과 PyroSim 입력값",
});

const ORDER = D.order;
const LABEL = D.label;
const ATM_N2 = "N₂", ATM_AIR = "공기(O)";
// material colours: categorical slots in an order that passes the palette check (adjacent CVD and normal vision)
const MAT = { EPS_C: "EDA100", EPS_SC: "E87BA4", PIR_C: "6250D6", PIR_SC: "008300", PF_out: "2A78D6", PF_in: "E34948" };
const N2_LINE = "4B5563";
const G = D.grid, GD = D.grid_dsc;
const mean = (m, atm, key) => {
  const a = D.curves[m][`${atm}#1`][key], b = D.curves[m][`${atm}#2`][key];
  return a.map((v, i) => (v + b[i]) / 2);
};
const T = (m, atm) => tab(D.table, { material: m, atm });
const V = (m, atm) => tab(D.verdict, { material: m, atm });
const ST = (m, atm) => P.stages.filter((s) => s.material === m && s.atmosphere === atm);
const R = (m, atm, k) => P.stages.find((s) => s.material === m && s.atmosphere === atm && s.stage === k);
const VERDICT = { reliable: "신뢰", conditional: "조건부", endpoint: "종점 의존", remeasure: "재측정" };
const sci = (v, n = 2) => Number(v).toExponential(n).toUpperCase().replace("E+", "E");
const REACTIONS = ORDER.flatMap((m) => [R(m, ATM_N2, 1), ...(m.startsWith("EPS") ? [] : [R(m, ATM_AIR, 2)])]);

// dot on a peak and its label (optionally away from the peak with a leader line); m = plotMap of the chart
function peak(s, m, x, y, color, label, o = {}) {
  const px = m.X(x), py = m.Y(y);
  s.addShape(pres.ShapeType.ellipse, { x: px - 0.05, y: py - 0.05, w: 0.1, h: 0.1, fill: { color }, line: { color: "FFFFFF", width: 0.75 } });
  if (o.at) {
    const tx = m.X(o.at[0]), ty = m.Y(o.at[1]);
    s.addShape(pres.ShapeType.line, { x: Math.min(px, tx), y: Math.min(py, ty), w: Math.abs(tx - px), h: Math.abs(ty - py),
      flipH: (tx < px) !== (ty < py), line: { color: HEX.ink2, width: 0.75 } });
    text(s, label, { x: tx + 0.04, y: ty - 0.12, w: o.w || 1.3, h: 0.24, fontSize: 10, valign: "middle" });
  } else {
    const w = o.w || 1.2, dx = o.dx || 0;
    const x0 = o.align === "right" ? px - w - 0.06 + dx : o.align === "left" ? px + 0.06 + dx : px - w / 2 + dx;
    text(s, label, { x: x0, y: py - 0.3 + (o.dy || 0), w, h: 0.24, fontSize: 10, align: o.align || "center", valign: "bottom" });
  }
}
function niceStep(range) {
  const raw = range / 7, p = 10 ** Math.floor(Math.log10(raw));
  return [1, 2, 2.5, 5, 10].map((k) => k * p).find((k) => k >= raw);
}
function axisRange(lo, hi) {
  const step = niceStep(hi - lo || 1);
  return { yMin: lo >= 0 ? 0 : Math.floor(lo / step) * step, yMax: Math.ceil(hi / step) * step, yMajor: step };
}
const matLegend = (s, y, x = MX) =>
  legend(s, ORDER.map((m) => ({ label: LABEL[m], color: MAT[m], w: LABEL[m].length > 8 ? 1.3 : 0.98 })), x, y);

// ---------------------------------------------------------------- 1 title
{
  pres.addSection({ title: "결론" });
  const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "결론" });
  s.addText("PyroSim 입력 자료 · 2026-10-21", { placeholder: "kicker", lang: LANG });
  s.addText("TGA·DSC 재료 특성 그래프와\nPyroSim 입력값 정리", { placeholder: "title", lang: LANG });
  s.addText("단열재 6종 — 밀도·비열은 어디서 오고, TGA–DSC는 PyroSim의 어디에 쓰는가", { placeholder: "subtitle", lang: LANG });
  const facts = [
    ["실측 필요", "밀도 · 비열 · 열전도율", "TGA–DSC로는 구할 수 없음"],
    [`${REACTIONS.length}개`, "TGA 반응 입력값", "열분해 6 + 탄화층 산화 4"],
    ["TGA_ANALYSIS", "FDS로 반응 입력 검증", "측정 곡선을 엑셀에 수록"],
  ];
  facts.forEach(([big, lab, sub], i) => {
    const x = 0.8 + i * 3.95;
    s.addShape(pres.ShapeType.roundRect, { x, y: 4.55, w: 3.7, h: 1.65, rectRadius: 0.08, fill: { color: C.background1, transparency: 90 }, line: { color: C.background2, width: 0.75 } });
    text(s, big, { x: x + 0.25, y: 4.7, w: 3.3, h: 0.6, fontSize: 26, bold: true, color: C.background1 });
    text(s, lab, { x: x + 0.25, y: 5.32, w: 3.3, h: 0.32, fontSize: 13, bold: true, color: C.background1 });
    text(s, sub, { x: x + 0.25, y: 5.66, w: 3.3, h: 0.32, fontSize: 11, color: C.background2 });
  });
  text(s, "함께 드리는 파일: PyroSim_입력값_정리.xlsx (입력값 · 반응 파라미터 · 밀도/비열 측정 기록표 · 측정 곡선), figures/material_*.png",
    { x: 0.8, y: 6.65, w: 11.7, h: 0.3, fontSize: 11, color: C.background2 });
  s.addNotes([
    "[스크립트]",
    "- 질문은 두 가지였습니다. TGA–DSC 결과를 PyroSim의 밀도·비열로 쓸 수 있는가, 그리고 재료 특성이 보이는 TGA·DSC 그래프입니다.",
    "- 답부터 말씀드리면, 밀도와 비열은 이 TGA–DSC로 구할 수 없습니다. 대신 TGA는 PyroSim 재료의 '열분해 반응' 입력(기준온도, 기준속도, 단계 수)에 바로 씁니다.",
    "- 순서: 결론과 입력 항목별 출처(2–3쪽) → 6종 비교 그래프(4–6쪽) → 재료별 특성(7–12쪽) → 반응 입력값과 입력 예시(13–14쪽) → 밀도·비열 측정 방법(15쪽) → 적용 시 주의(16쪽).",
    "",
    "[질문 대비]",
    "- 같은 숫자가 엑셀(PyroSim_입력값_정리.xlsx)에 있고, 기준온도·속도를 바꾸면 FDS가 계산할 A·E가 수식으로 다시 계산됩니다.",
  ].join("\n"));
}

// ---------------------------------------------------------------- 2 conclusion
{
  const s = content("결론", "결론 — PyroSim에 무엇을 어떻게 쓰나",
    "밀도·비열은 이 TGA–DSC로 구할 수 없음 — TGA는 PyroSim ‘열분해 반응’ 입력에 쓰고,\n밀도·비열·열전도율은 보드 실측이나 제품 시험성적서 값을 씀");
  const cards = [
    ["good", "TGA → 그대로 사용", [
      ["REFERENCE_TEMPERATURE", "DTG 피크 온도"],
      ["REFERENCE_RATE", "5 K/min × 피크 속도 ÷ 단계 몫"],
      ["N_REACTIONS", "EPS 1 · PIR(PUR) · PF 2"],
      ["HEATING_RATE", "5 K/min"],
    ], "전환율 기준이라 초기 질량 없이 계산됨"],
    ["fair", "부호 · 조건만 확인", [
      ["HEAT_OF_REACTION", "부호만: EPS 398 °C 흡열, PF 내부용 474 °C 발열"],
      ["NU_MATL (잔류물)", "초기 질량 기록이 있으면 계산"],
      ["산화 A · E", "O 가스가 공기인지 확인 후 A ÷ 0.209"],
    ], "기록 확인 · 교정 DSC가 있으면 값이 됨"],
    ["poor", "이 시험으로 불가 → 실측", [
      ["DENSITY 밀도", "보드 실측 (KS M ISO 845)"],
      ["SPECIFIC_HEAT 비열", "사파이어법 DSC (ASTM E1269)"],
      ["CONDUCTIVITY 열전도율", "시험성적서 (KS L 9016)"],
      ["HEAT_OF_COMBUSTION", "MCC · 콘칼로리미터"],
    ], "임시값: ρ 15–60 kg/m³, c 1.3–1.5 kJ/(kg·K), k 0.018–0.040 W/(m·K)"],
  ];
  const cw = (CW - 2 * 0.25) / 3;
  cards.forEach(([st, head, items, foot], i) => {
    const x = MX + i * (cw + 0.25), y = 1.8;
    s.addShape(pres.ShapeType.roundRect, { x, y, w: cw, h: 4.2, rectRadius: 0.08, fill: { color: STATUS[st].tint }, line: { color: STATUS[st].tint, width: 0.5 }, objectName: `card ${head}` });
    s.addShape(pres.ShapeType.ellipse, { x: x + 0.22, y: y + 0.27, w: 0.22, h: 0.22, fill: { color: STATUS[st].color }, line: { color: STATUS[st].color, width: 0.5 } });
    text(s, head, { x: x + 0.54, y: y + 0.17, w: cw - 0.7, h: 0.42, fontSize: 17, bold: true, valign: "middle" });
    items.forEach(([k, v], j) => {
      const yy = y + 0.8 + j * 0.7;
      text(s, k, { x: x + 0.24, y: yy, w: cw - 0.44, h: 0.28, fontSize: 12.5, bold: true });
      text(s, v, { x: x + 0.24, y: yy + 0.29, w: cw - 0.44, h: 0.36, fontSize: 11.5, color: C.text2 });
    });
    text(s, foot, { x: x + 0.24, y: y + 3.6, w: cw - 0.44, h: 0.5, fontSize: 10.5, color: C.text2, valign: "bottom" });
  });
  s.addShape(pres.ShapeType.roundRect, { x: MX, y: 6.2, w: CW, h: 0.66, rectRadius: 0.08, fill: { color: C.text1 }, line: { color: C.text1, width: 0.5 }, objectName: "guide" });
  text(s, [
    { text: "주의  ", options: { bold: true, fontSize: 13.5 } },
    { text: "카탈로그 물성만 넣은 맨 단열재는 시험보다 훨씬 빨리 착화함(모델 16–51 s, 준불연 시험 188–288 s) → 면재 · 마감층을 넣고 중공층 시험으로 검증 (16쪽)", options: { fontSize: 12 } },
  ], { x: MX + 0.25, y: 6.2, w: CW - 0.5, h: 0.66, color: C.background1, valign: "middle" });
  s.addNotes([
    "[스크립트]",
    "- 왼쪽(초록): TGA로 바로 쓰는 값입니다. PyroSim(FDS)은 TGA 피크 온도(REFERENCE_TEMPERATURE)와 그 온도의 반응속도(REFERENCE_RATE), 승온속도(HEATING_RATE)를 넣으면 반응식의 A와 E를 스스로 계산합니다.",
    "- 이 자료의 전환율 α는 '790 °C까지 감소량' 기준이라, 단계의 몫 f로 나누면 그 단계(반응 성분)의 속도가 됩니다. 그래서 초기 질량이 없어도 REFERENCE_RATE는 계산됩니다.",
    "- 가운데(노랑): 반응열은 부호만, 잔류물 수율은 초기 질량 기록이 있어야 하고, 산화 반응의 A는 O 가스가 공기인지 확인해야 합니다.",
    "- 오른쪽(빨강): 밀도·비열·열전도율·연소열은 이 시험으로 구할 수 없습니다. 밀도는 보드를 재고, 비열은 같은 장비로 사파이어법 측정을 하면 됩니다(15쪽).",
    "- 아래 띠: 가장 큰 함정입니다. 기존 역산에서 카탈로그 물성의 맨 단열재는 16–51 s면 분해온도에 도달했지만, 시험의 준불연 재료는 188–288 s에 착화했습니다.",
    "",
    "[질문 대비]",
    "- 임시값의 출처: 기존 연계 분석(coupled_analysis)에서 쓴 단열 보드 카탈로그 범위입니다. 제품 시험성적서 값이 있으면 그것을 우선합니다.",
    "- DSC 열류로 비열을 못 구하는 이유: 시료의 현열 신호가 0.1–0.35 mW인데 기준선이 111–180 mW 움직이고, 교정·블랭크·초기 질량이 없기 때문입니다(15쪽).",
  ].join("\n"));
}

// ---------------------------------------------------------------- 3 input map
{
  const s = content("결론", "PyroSim 재료(MATL) 입력 항목별 출처",
    "PyroSim 재료 입력 12개 항목 중 TGA로 바로 채우는 것은 반응 4개 항목 —\n나머지는 실측 · 시험성적서 · 기록 확인으로 채움");
  const hh = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 }, valign: "middle" } });
  const can = { ok: ["good", "가능"], part: ["fair", "부호만"], rec: ["fair", "기록 필요"], asm: ["fair", "가정 필요"], no: ["poor", "불가"] };
  const rows = [
    ["밀도", "DENSITY", "kg/m³", "no", "보드 실측 (KS M ISO 845) 또는 시험성적서 — 임시 15–60"],
    ["비열", "SPECIFIC_HEAT (_RAMP)", "kJ/(kg·K)", "no", "사파이어법 DSC (ASTM E1269 / ISO 11357-4) — 임시 1.3–1.5"],
    ["열전도율", "CONDUCTIVITY (_RAMP)", "W/(m·K)", "no", "시험성적서 (KS L 9016 / ISO 8301) — 임시 0.018–0.040"],
    ["방사율", "EMISSIVITY", "–", "no", "FDS 기본값 0.9"],
    ["반응 수", "N_REACTIONS", "–", "ok", "DTG 단계: EPS 1, PIR(PUR) · PF 2 (열분해 + 탄화층 산화)"],
    ["기준온도", "REFERENCE_TEMPERATURE", "°C", "ok", "DTG 피크 온도 (반응 1 = N₂, 반응 2 = 공기)"],
    ["기준속도", "REFERENCE_RATE", "1/s", "ok", "5 K/min × 피크 속도 ÷ 단계 몫 f"],
    ["승온속도", "HEATING_RATE", "K/min", "ok", "5 (TGA 조건)"],
    ["잔류물 수율", "NU_MATL", "kg/kg", "rec", "790 °C 잔류 질량 ÷ 초기 질량 — 초기 질량 24개 기록 필요"],
    ["반응열", "HEAT_OF_REACTION", "kJ/kg", "part", "교정 DSC (블랭크 + 열류 교정 + 초기 질량) 또는 문헌"],
    ["산소 의존 산화", "N_O2, A, E", "–", "asm", "공기 TGA 산화 피크 → A ÷ 0.209 (O 가스가 공기일 때)"],
    ["연소열", "HEAT_OF_COMBUSTION", "kJ/kg", "no", "MCC (ASTM D7309) 또는 콘칼로리미터 (ISO 5660-1)"],
  ];
  const tb = [[hh("PyroSim 항목"), hh("FDS 키워드"), hh("단위"), hh("이 자료로"), hh("값을 얻는 방법")]];
  rows.forEach(([a, b, u, k, how]) => {
    const [st, lab] = can[k];
    tb.push([{ text: a, options: { bold: true } }, { text: b }, { text: u, options: { align: "center" } },
      { text: lab, options: { align: "center", bold: true, fill: { color: STATUS[st].tint } } }, { text: how, options: { color: C.text2 } }]);
  });
  table(s, tb, { x: MX, y: 1.8, w: CW, colW: [1.5, 2.45, 1.05, 1.1, 6.233], fontSize: 11, color: C.text1, valign: "middle",
    border: { type: "solid", pt: 0.75, color: "D1D5DB" }, margin: [0.03, 0.08, 0.03, 0.08], rowH: 0.35 });
  text(s, "임시값 = 기존 분석에서 쓴 단열 보드 카탈로그 범위 (실측 전 시험 계산용). FDS 키워드는 PyroSim 재료 편집창의 같은 이름 항목에 대응",
    { x: MX, y: 6.72, w: CW, h: 0.26, fontSize: 10, color: C.text2 });
  s.addNotes([
    "[스크립트]",
    "- PyroSim 재료 편집창의 열물성(Thermal Properties) 항목과 열분해(Pyrolysis) 반응 항목을 FDS 키워드와 함께 정리했습니다.",
    "- TGA로 바로 채우는 것은 반응 수, 기준온도, 기준속도, 승온속도 네 가지입니다.",
    "- 잔류물 수율은 초기 질량 기록만 있으면 계산되고, 반응열은 부호만 확인된 상태입니다. 산화 반응의 A는 O 가스가 공기라는 가정이 필요합니다.",
    "- 밀도·비열·열전도율·연소열은 다른 시험이나 시험성적서로 채웁니다.",
    "",
    "[질문 대비]",
    "- 'REFERENCE_RATE와 A·E를 같이 넣어도 되나': FDS 설명서는 REFERENCE_TEMPERATURE를 쓰면 A·E를 넣지 말라고 합니다. 단, 산소 의존(N_O2 = 1) 산화 반응은 A·E를 직접 넣습니다(14쪽).",
  ].join("\n"));
}

// ---------------------------------------------------------------- 4, 5 overview TGA
function overviewPair(s, key, yAx, ylab, peaks) {
  const L = { x: 0.1, y: 0.04, w: 0.86, h: 0.82 };
  const maps = [];
  [ATM_N2, ATM_AIR].forEach((atm, c) => {
    const ch = { x: MX + c * 6.25, y: 2.45, w: 6.08, h: 4.45, L, xMin: 0, xMax: 800, xMajor: 100, ...yAx };
    text(s, `${c ? "(b) 공기(O) — 산화" : "(a) N₂ — 열분해"}`, { x: ch.x + 0.05, y: 2.15, w: 4, h: 0.3, fontSize: 12.5, bold: true, valign: "middle" });
    curves(s, ch, G, ORDER.map((m) => ({ name: `${LABEL[m]} ${atm}`, ys: mean(m, atm === ATM_N2 ? "N2" : "O", key), color: MAT[m] })),
      { lineSize: 1.75, objectName: `${key} overview ${atm}` });
    const m = plotMap(ch);
    maps.push(m);
    text(s, "온도 (°C)", { x: m.px, y: m.py + m.ph + 0.27, w: m.pw, h: 0.22, fontSize: 10.5, color: C.text2, align: "center" });
    if (peaks) peaks(s, m, atm);
  });
  vlabel(s, ylab, 0.2, maps[0].py + maps[0].ph / 2, 3.4);
  matLegend(s, 1.8);
}
{
  pres.addSection({ title: "재료 특성 그래프" });
  const s = content("재료 특성 그래프", "TGA 그래프 ① — 6종 질량 잔류 (두 반복 평균)",
    "EPS는 한 단계로 거의 다 사라지고, PIR(PUR) · PF는 탄화층이 남아 N₂에서 천천히 줄어듦 —\n공기에서는 탄화층이 산화돼 PIR(PUR) · PF 외부용은 600 °C 전후, PF 내부용은 480 °C에 끝남");
  overviewPair(s, "rem", { yMin: -10, yMax: 110, yMajor: 20 }, "잔류 질량 100·(1−α) (%)");
  s.addNotes([
    "[스크립트]",
    "- 질량 잔류는 승온 시작부터 790 °C까지 줄어든 양을 100 %로 놓은 값입니다. 초기 질량 대비가 아니므로 0 %가 '재가 없다'는 뜻은 아닙니다.",
    "- 왼쪽 N₂(열분해): EPS 2종은 400 °C 전후 한 번에 떨어집니다. PIR(PUR)은 300 °C대에서 절반쯤 줄고 나머지(탄화층)는 790 °C까지 천천히 줄어듭니다. PF는 처음부터 완만하게 줄어 790 °C에서도 진행 중입니다.",
    "- 오른쪽 공기(산화): EPS는 80–100 °C 일찍 분해하고, PIR(PUR)과 PF는 탄화층이 산화되면서 600 °C 전후에 끝납니다. PF 내부용만 480 °C 부근에서 급격히 끝납니다.",
    "",
    "[질문 대비]",
    "- EPS 가연성 공기 곡선이 100 %를 넘는 구간: 부력에 의한 겉보기 증가입니다(질량감소 0.16 · 0.30 mg, 재측정 대상).",
  ].join("\n"));
}
{
  const s = content("재료 특성 그래프", "TGA 그래프 ② — 6종 질량감소속도 DTG (두 반복 평균)",
    "N₂ 열분해 피크: EPS 397 · 426 °C, PIR(PUR) 315 · 332 °C (PF는 뚜렷한 피크 없음)\n공기 산화 피크: PF 내부용 472 °C(1.17 %/°C)가 가장 빠르고 PF 외부용 488 °C는 넓고 낮음");
  const at = { [`PIR_C ${ATM_N2}`]: [150, 1.05], [`PIR_SC ${ATM_N2}`]: [150, 0.8], [`PIR_C ${ATM_AIR}`]: [600, 1.25],
    [`PIR_SC ${ATM_AIR}`]: [600, 1.0], [`PF_out ${ATM_AIR}`]: [600, 0.75] };
  overviewPair(s, "dtg", { yMin: -0.5, yMax: 3.0, yMajor: 0.5 }, "dα/dT (%/°C)", (sl, m, atm) => {
    ORDER.forEach((mat) => {
      const st = ST(mat, atm).filter((x) => x.finished).sort((a, b) => b.peak_rate_pct_per_C - a.peak_rate_pct_per_C)[0];
      if (!st || (mat.startsWith("PF") && atm === ATM_N2)) return;
      const lab = `${fmt(st.REFERENCE_TEMPERATURE_C, 0)} °C · ${fmt(st.peak_rate_pct_per_C)}`;
      peak(sl, m, st.REFERENCE_TEMPERATURE_C, st.peak_rate_pct_per_C, MAT[mat], lab, { at: at[`${mat} ${atm}`], w: 1.35 });
    });
  });
  text(s, "점 = 주 피크 (온도 · 최대속도 %/°C). PF 2종은 N₂에서 뚜렷한 주 피크가 없어 표시하지 않음", { x: MX, y: 6.84, w: CW, h: 0.2, fontSize: 9.5, color: C.text2 });
  s.addNotes([
    "[스크립트]",
    "- DTG는 1 °C마다 줄어드는 몫입니다. 피크 온도가 PyroSim의 REFERENCE_TEMPERATURE가 됩니다.",
    "- N₂: EPS 가연성 397 °C(2.65 %/°C)가 가장 날카롭고, 준불연은 426 °C로 27 °C 늦고 속도는 절반 이하입니다. PIR(PUR)은 315 · 332 °C의 1단계 피크입니다. PF는 피크 없이 800 °C를 향해 계속 올라가 표시하지 않았습니다.",
    "- 공기: PF 내부용 472 °C(1.17 %/°C)가 탄화층 산화 중 가장 빠르고, PF 외부용은 488 °C(0.40)로 넓고 낮습니다. PIR(PUR) 탄화층 산화는 524 · 534 °C입니다.",
    "",
    "[질문 대비]",
    "- EPS 가연성 공기 피크(294 °C)는 재측정 대상 자료라 참고만 합니다.",
  ].join("\n"));
}

// ---------------------------------------------------------------- 6 overview DSC (offset traces)
{
  const s = content("재료 특성 그래프", "DSC 그래프 — 6종 국소 열류 (정성, 두 반복)",
    "DSC에서 두 반복이 같게 나온 특성은 EPS 가연성 398 °C 흡열과 PF 내부용 474 °C 발열\n나머지 재료의 신호는 반복마다 크기 · 온도가 달라 판정에 쓰지 않음 (보정 없는 정성 자료)");
  const step = 4.5, scale = { "PF_in O": 0.2 };
  const L = { x: 0.03, y: 0.03, w: 0.94, h: 0.86 };
  [["N2", ATM_N2], ["O", ATM_AIR]].forEach(([atm, atmKo], c) => {
    const ch = { x: MX + c * 6.25, y: 2.2, w: 6.08, h: 4.7, L, xMin: 150, xMax: 720, xMajor: 100, yMin: -ORDER.length * step + 0.6, yMax: 2.4, yMajor: 100 };
    text(s, `${c ? "(b) 공기(O) — 산화" : "(a) N₂ — 열분해"} · 흡열 ↑ 발열 ↓`, { x: ch.x + 0.05, y: 1.85, w: 5.5, h: 0.3, fontSize: 12.5, bold: true, valign: "middle" });
    const series = [];
    ORDER.forEach((m, i) => {
      const k = scale[`${m} ${atm}`] || 1;
      [1, 2].forEach((rep) => series.push({ name: `${LABEL[m]} ${atmKo} #${rep}`, ys: D.dsc[m][`${atm}#${rep}`].map((v) => -i * step + k * v), color: MAT[m], style: rep === 2 ? { dash: "dash" } : undefined }));
    });
    curves(s, ch, GD, series, { valAxisHidden: true, valGridLine: { style: "none" }, lineSize: 1.25, objectName: `dsc overview ${atm}` });
    const m = plotMap(ch);
    ORDER.forEach((mat, i) => {
      const k = scale[`${mat} ${atm}`] || 1;
      text(s, LABEL[mat] + (k !== 1 ? " (×1/5)" : ""), { x: m.px + 0.05, y: m.Y(-i * step) - 0.42, w: 2.4, h: 0.22, fontSize: 10, bold: true, valign: "bottom" });
    });
    text(s, "온도 (°C)", { x: m.px, y: m.py + m.ph + 0.27, w: m.pw, h: 0.22, fontSize: 10.5, color: C.text2, align: "center" });
    // 2 mW scale bar
    const bx = m.X(705), y0 = m.Y(-ORDER.length * step + 1.0);
    s.addShape(pres.ShapeType.line, { x: bx, y: y0 - (m.Y(0) - m.Y(2)), w: 0, h: m.Y(0) - m.Y(2), line: { color: HEX.ink, width: 1.75 } });
    text(s, "2 mW", { x: bx - 0.75, y: y0 - 0.25, w: 0.7, h: 0.22, fontSize: 10, align: "right" });
    if (atm === "N2") {
      const y = m.Y(0.95);
      text(s, "398 °C 흡열 (해중합) — 두 반복 재현", { x: m.X(410), y: y - 0.2, w: 3.0, h: 0.24, fontSize: 10, bold: true });
    } else {
      const i = ORDER.indexOf("PF_in");
      text(s, "474 °C 발열 — 두 반복 재현", { x: m.X(455) - 2.0, y: m.Y(-i * step - 2.4), w: 2.0, h: 0.24, fontSize: 10, bold: true, align: "right" });
    }
  });
  text(s, "실선 #1 · 점선 #2. 국소 열류 = 미보정 열류에서 100 °C 창 추세를 뺀 값(정성) — 큰 발열 양옆의 + 곁봉우리는 필터 효과",
    { x: MX, y: 6.82, w: CW, h: 0.2, fontSize: 9.5, color: C.text2 });
  s.addNotes([
    "[스크립트]",
    "- DSC는 보정 없이 기록돼 정성 자료입니다. 재료마다 세로로 띄워 그렸고, 오른쪽 아래 막대가 2 mW입니다.",
    "- 두 반복에서 같은 온도 · 부호로 나온 재료 특성은 2개입니다. EPS 가연성 N₂ 398 °C 흡열(질량감소 피크 397 °C와 같은 온도, 해중합)과 PF 내부용 공기 474 °C 발열(탄화층 산화, 1/5로 줄여 그림)입니다.",
    "- 나머지는 반복마다 크기나 온도가 달라 재료 특성으로 쓰지 않습니다. PIR(PUR) 600 °C 이후의 큰 흔들림도 한 반복에만 있습니다.",
    "",
    "[질문 대비]",
    "- PyroSim 반응열에 쓸 수 있나: 부호(흡열 · 발열)만 씁니다. 크기(kJ/kg)는 교정 DSC가 필요합니다.",
  ].join("\n"));
}

// ---------------------------------------------------------------- 7-12 one slide per material
const SHEET_TITLE = {
  EPS_C: "EPS 가연성 — N₂ 397 °C에서 한 단계로 급격히 분해(2.65 %/°C)하고\n공기에서는 296 °C로 약 100 °C 앞당겨짐 · DSC 398 °C 흡열(해중합)이 두 반복에서 재현",
  EPS_SC: "EPS 준불연 — N₂ 426 °C로 가연성보다 27 °C 늦고 최대속도는 절반 이하\n공기에서는 342 °C에서 주분해한 뒤 잔류물 산화가 790 °C까지 이어짐",
  PIR_C: "PIR(PUR) 가연성 — N₂ 315 °C 1단계 뒤 탄화층이 천천히 분해\n공기에서는 302 °C 분해와 524 °C 탄화층 산화로 600 °C 전에 다 탐",
  PIR_SC: "PIR(PUR) 준불연 — 가연성과 같은 2단계(N₂ 332 °C, 공기 310 · 534 °C)\n속도는 가연성의 0.7–0.8배이고, 공기 중에는 탄화층 산화가 주 단계",
  PF_out: "PF 준불연 외부용 — N₂에서는 완만한 탄화가 790 °C에서도 진행\n공기에서는 488 °C의 넓고 낮은 산화 피크(0.40 %/°C)로 천천히 소실",
  PF_in: "PF 준불연 내부용 — N₂ 열분해는 외부용과 비슷하지만\n공기에서는 472 °C 좁은 피크(1.17 %/°C)로 급속 전소 · 같은 온도에 DSC 큰 발열",
};
ORDER.forEach((mat, idx) => {
  const s = content("재료 특성 그래프", `재료별 특성 ${idx + 1}/6 — ${LABEL[mat]}`, SHEET_TITLE[mat]);
  const col = { N2: N2_LINE, O: MAT[mat] };
  legend(s, [{ label: "N₂ (열분해)", color: N2_LINE, w: 1.05 }, { label: "공기(O) (산화)", color: MAT[mat], w: 1.25 }, { label: "DSC 2회차 (흡열 +)", color: HEX.ink2, dash: true, w: 1.8 }], MX, 1.78);
  const L = { x: 0.1, y: 0.06, w: 0.88, h: 0.76 };
  const W8 = 8.15, H = 1.55, Y0 = 2.12, GAP = 0.04;
  // (a) mass
  const remAll = [...mean(mat, "N2", "rem"), ...mean(mat, "O", "rem")];
  const ca = { x: MX, y: Y0, w: W8, h: H, L, xMin: 0, xMax: 800, xMajor: 100, yMin: Math.min(...remAll) < -1 ? -20 : 0,
    yMax: Math.max(...remAll) > 105 ? 120 : 110, yMajor: 20 };
  curves(s, ca, G, [{ name: `N₂ 평균 ${mat}`, ys: mean(mat, "N2", "rem"), color: col.N2 }, { name: `공기 평균 ${mat}`, ys: mean(mat, "O", "rem"), color: col.O }], { lineSize: 1.75, objectName: `mass ${mat}` });
  // (b) DTG
  const dAll = [...mean(mat, "N2", "dtg"), ...mean(mat, "O", "dtg")];
  const cb = { x: MX, y: Y0 + H + GAP, w: W8, h: H, L, xMin: 0, xMax: 800, xMajor: 100, ...axisRange(Math.min(...dAll) < -0.02 ? Math.min(...dAll) : 0, Math.max(...dAll) * 1.12) };
  curves(s, cb, G, [{ name: `N₂ DTG ${mat}`, ys: mean(mat, "N2", "dtg"), color: col.N2 }, { name: `공기 DTG ${mat}`, ys: mean(mat, "O", "dtg"), color: col.O }], { lineSize: 1.75, objectName: `dtg ${mat}` });
  const mb = plotMap(cb);
  const pmax = Math.max(...P.stages.filter((x) => x.material === mat).map((x) => x.peak_rate_pct_per_C));
  // label the PyroSim reaction stages and any other stage at least a quarter of the material's highest peak
  [[ATM_N2, "N2"], [ATM_AIR, "O"]].forEach(([atmKo, atm]) => ST(mat, atmKo).filter((x) => x.finished && (x.role.startsWith("반응") || x.peak_rate_pct_per_C >= 0.25 * pmax)).forEach((st) =>
    peak(s, mb, st.REFERENCE_TEMPERATURE_C, st.peak_rate_pct_per_C, col[atm], `${fmt(st.REFERENCE_TEMPERATURE_C, 0)} °C`, { align: atm === "N2" ? "right" : "left", w: 0.85 })));
  // (c) DSC, both runs
  const hAll = ["N2#1", "N2#2", "O#1", "O#2"].flatMap((k) => D.dsc[mat][k]);
  const cc = { x: MX, y: Y0 + 2 * (H + GAP), w: W8, h: H, L, xMin: 0, xMax: 800, xMajor: 100, ...axisRange(Math.min(...hAll) * 1.1, Math.max(...hAll) * 1.1) };
  curves(s, cc, GD, [
    { name: `N₂ DSC ${mat} #1`, ys: D.dsc[mat]["N2#1"], color: col.N2 }, { name: `N₂ DSC ${mat} #2`, ys: D.dsc[mat]["N2#2"], color: col.N2, style: { dash: "dash" } },
    { name: `공기 DSC ${mat} #1`, ys: D.dsc[mat]["O#1"], color: col.O }, { name: `공기 DSC ${mat} #2`, ys: D.dsc[mat]["O#2"], color: col.O, style: { dash: "dash" } },
  ], { lineSize: 1.25, objectName: `dsc ${mat}` });
  const mc = plotMap(cc);
  if (mat === "EPS_C") text(s, "398 °C 흡열 — 재현", { x: mc.X(420), y: mc.Y(1.0) - 0.05, w: 1.8, h: 0.22, fontSize: 10, bold: true });
  if (mat === "PF_in") text(s, "474 °C 발열 — 재현", { x: mc.X(490), y: mc.Y(-15) - 0.12, w: 1.8, h: 0.22, fontSize: 10, bold: true });
  [["질량 잔류 (%)", ca], ["DTG (%/°C)", cb], ["DSC (mW)", cc]].forEach(([t, ch]) => {
    const m = plotMap(ch);
    vlabel(s, t, 0.6, m.py + m.ph / 2, m.ph);
  });
  text(s, "온도 (°C)", { x: mc.px, y: mc.py + mc.ph + 0.25, w: mc.pw, h: 0.2, fontSize: 10, color: C.text2, align: "center" });
  // numbers panel
  const x0 = 8.95, w0 = 3.88;
  s.addShape(pres.ShapeType.roundRect, { x: x0, y: 1.8, w: w0, h: 5.12, rectRadius: 0.08, fill: { color: C.background2 }, line: { color: C.background2, width: 0.5 }, objectName: `numbers ${mat}` });
  const tn = T(mat, ATM_N2), ta = T(mat, ATM_AIR);
  const r1 = R(mat, ATM_N2, 1), r2 = mat.startsWith("EPS") ? null : R(mat, ATM_AIR, 2);
  const dscTxt = { EPS_C: "N₂ 398 °C 흡열 (해중합) — 재현", PF_in: "공기 474 °C 발열 (탄화층 산화) — 재현" }[mat] || "두 반복에서 재현된 사건 없음";
  text(s, paras([
    [["열분해 (N₂) · ", { bold: true, color: C.text1 }], [VERDICT[V(mat, ATM_N2).overall], { bold: true, color: C.text1 }]],
    `T_onset ${fmt(tn.T_onset_C, 0)} · T_peak ${fmt(tn.T_peak_C, 0)} °C`,
    `r_max ${fmt(tn.r_max_pct_C)} %/°C · f₄₀₀ ${fmt(tn.f_above400)}`,
    [["산화 (공기) · ", { bold: true, color: C.text1 }], [VERDICT[V(mat, ATM_AIR).overall], { bold: true, color: C.text1 }]],
    `T_onset ${fmt(ta.T_onset_C, 0)} · T_peak ${fmt(ta.T_peak_C, 0)} °C`,
    `r_max ${fmt(ta.r_max_pct_C)} %/°C · f₄₀₀ ${fmt(ta.f_above400)}`,
    [["DSC", { bold: true, color: C.text1 }]],
    dscTxt,
    [["PyroSim 반응 시작값", { bold: true, color: C.text1 }]],
    `반응 1 (N₂): ${fmt(r1.REFERENCE_TEMPERATURE_C, 0)} °C, ${sci(r1.REFERENCE_RATE_per_s)} 1/s`,
    ...(r2 ? [`반응 2 (공기, N_O2 = 1): ${fmt(r2.REFERENCE_TEMPERATURE_C, 0)} °C, ${sci(r2.REFERENCE_RATE_per_s)} 1/s`] : []),
  ], { bullet: false, space: 3 }), { x: x0 + 0.2, y: 1.95, w: w0 - 0.35, h: 4.0, fontSize: 11, color: C.text2 });
  text(s, "밀도 · 비열 · 열전도율은 이 시험에서 나오지 않음 → 15쪽", { x: x0 + 0.2, y: 6.38, w: w0 - 0.35, h: 0.45, fontSize: 10, color: C.text2, valign: "bottom" });
  s.addNotes([
    "[스크립트]",
    `- ${LABEL[mat]}의 TGA(질량 잔류 · DTG)와 DSC를 N₂(회색)와 공기(색)로 겹쳤습니다. 오른쪽에 지표와 PyroSim 반응 시작값을 적었습니다.`,
    `- N₂: T_peak ${fmt(tn.T_peak_C, 0)} °C, r_max ${fmt(tn.r_max_pct_C)} %/°C, 자료 판정 ${VERDICT[V(mat, ATM_N2).overall]}. 공기: T_peak ${fmt(ta.T_peak_C, 0)} °C, r_max ${fmt(ta.r_max_pct_C)} %/°C, 자료 판정 ${VERDICT[V(mat, ATM_AIR).overall]}.`,
    `- PyroSim 반응 1(N₂ 열분해): REFERENCE_TEMPERATURE ${fmt(r1.REFERENCE_TEMPERATURE_C, 0)} °C, REFERENCE_RATE ${r1.REFERENCE_RATE_per_s} 1/s (두 반복 ${r1.REFERENCE_RATE_runs}), HEATING_RATE 5 K/min.`,
    ...(r2 ? [`- 반응 2(공기 탄화층 산화): ${fmt(r2.REFERENCE_TEMPERATURE_C, 0)} °C, ${r2.REFERENCE_RATE_per_s} 1/s. N_O2 = 1로 넣으면 A = ${r2.A_per_s_if_N_O2_1_air} 1/s, E = ${r2.E_kJ_mol} kJ/mol을 직접 입력(공기 가정).`] : []),
    "",
    "[질문 대비]",
    "- 질량 잔류 0 %는 '790 °C까지 감소량 기준'의 끝이라는 뜻이고 실제 잔류율이 아닙니다. 잔류물 수율(NU_MATL)은 초기 질량 기록이 있어야 합니다.",
    ...(mat === "EPS_C" ? ["- 공기 곡선은 재측정 대상(질량감소 0.16 · 0.30 mg)이라 형태만 참고합니다."] : []),
    ...(mat.startsWith("PF") ? ["- N₂ 2단계는 790 °C에서 끝나지 않아 반응 입력에서 뺐습니다. 800 °C 등온 유지 측정 뒤 보완합니다."] : []),
  ].join("\n"));
});

// ---------------------------------------------------------------- 13 reaction table
{
  pres.addSection({ title: "PyroSim 입력" });
  const s = content("PyroSim 입력", "PyroSim 반응 입력 시작값",
    "반응 1(N₂ 열분해) 6종 · 반응 2(공기 탄화층 산화) 4종 — 기준온도는 두 반복이 거의 같음\n기준속도는 근사값이므로 FDS TGA_ANALYSIS로 측정 곡선과 맞춰 확인");
  const hh = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 }, align: "center", valign: "middle" } });
  const tb = [[hh("재료"), hh("반응"), hh("REFERENCE_TEMPERATURE\n(°C, 두 반복)"), hh("REFERENCE_RATE\n(1/s, 두 반복)"), hh("FDS가 계산하는\nE (kJ/mol) · A (1/s)"), hh("N_O2 = 1일 때 A\n(공기 가정)"), hh("자료 판정 · 비고")]];
  const note = { [`EPS_C ${ATM_N2}`]: "잔류물 거의 없을 것(확인)", [`PIR_C ${ATM_N2}`]: "탄화층 분해 미완료", [`PIR_SC ${ATM_N2}`]: "탄화층 분해 미완료",
    [`PF_out ${ATM_N2}`]: "피크 넓음 · 2단계 미완료", [`PF_in ${ATM_N2}`]: "피크 넓음 · 2단계 미완료", [`PF_out ${ATM_AIR}`]: "열분해와 겹친 값" };
  REACTIONS.forEach((r) => {
    const air = r.atmosphere === ATM_AIR;
    tb.push([
      { text: LABEL[r.material], options: { bold: true } },
      { text: air ? "2: 탄화층 산화 (공기)" : "1: 열분해 (N₂)" },
      { text: `${fmt(r.REFERENCE_TEMPERATURE_C, 0)}  (${r.T_peak_runs_C.split(" / ").join(" · ")})`, options: { align: "center" } },
      { text: `${sci(r.REFERENCE_RATE_per_s)}  (${r.REFERENCE_RATE_runs.split(" / ").map((v) => sci(v)).join(" · ")})`, options: { align: "center" } },
      { text: `${fmt(r.E_kJ_mol, 0)} · ${sci(r.A_per_s)}`, options: { align: "center" } },
      { text: air ? sci(r.A_per_s_if_N_O2_1_air) : "—", options: { align: "center" } },
      { text: VERDICT[r.tga_verdict] + (note[`${r.material} ${r.atmosphere}`] ? ` · ${note[`${r.material} ${r.atmosphere}`]}` : ""),
        options: { fill: { color: STATUS[r.tga_verdict].tint } } },
    ]);
  });
  table(s, tb, { x: MX, y: 1.8, w: CW, colW: [1.5, 1.75, 1.45, 2.55, 1.45, 1.1, 2.533], fontSize: 10.5, color: C.text1, valign: "middle",
    border: { type: "solid", pt: 0.75, color: "D1D5DB" }, margin: [0.03, 0.06, 0.03, 0.06], rowH: [0.55, ...new Array(REACTIONS.length).fill(0.34)] });
  text(s, paras([
    "REFERENCE_RATE = 5 K/min × DTG 피크 속도 ÷ 단계 몫 f — 790 °C까지 감소량 기준 전환율로 구해 초기 질량이 필요 없음",
    "E · A는 FDS가 기준온도 · 기준속도 · 승온속도로 스스로 계산하는 값(1차 반응): REFERENCE_TEMPERATURE를 쓰면 A · E는 비워 둠",
    "탄화층 산화를 N_O2 = 1로 넣을 때만 A(공기 가정) · E를 직접 입력 — O 가스가 순수 산소였다면 A를 나누지 않음",
  ], { space: 2 }), { x: MX, y: 6.14, w: CW, h: 0.8, fontSize: 10, color: C.text2 });
  s.addNotes([
    "[스크립트]",
    "- 재료마다 반응 1은 N₂ TGA의 열분해 피크, 반응 2는 공기 TGA의 탄화층 산화 피크에서 구했습니다. EPS는 한 단계라 반응 1만 있습니다.",
    "- 괄호 안이 두 반복 값입니다. 기준온도는 반복 간 차이가 대부분 몇 °C로 작습니다. 기준속도는 단계를 나누는 위치에 따라 달라지는 근사값입니다.",
    "- E와 A는 FDS가 내부에서 계산할 값을 미리 보여 드린 것입니다. EPS 가연성 274 kJ/mol은 폴리스티렌 열분해로 알려진 범위(대략 200–300 kJ/mol)에 들어갑니다(참고).",
    "",
    "[질문 대비]",
    "- PF 반응 1(N₂ 300 °C)의 E가 38–45 kJ/mol로 작은 이유: 피크가 넓고 낮아(0.13–0.19 %/°C) 1차 반응 한 개로 맞추면 넓은 반응이 됩니다. PF의 N₂ 분해는 790 °C에서도 끝나지 않아 반응식 자체를 보완해야 합니다.",
    "- 엑셀 '반응 파라미터' 시트에서 기준온도나 속도를 바꾸면 E · A가 같은 식으로 다시 계산됩니다.",
  ].join("\n"));
}

// ---------------------------------------------------------------- 14 how to enter (FDS example)
{
  const s = content("PyroSim 입력", "입력 방법 — PyroSim 재료 편집창과 FDS 입력 예시",
    "PIR(PUR) · PF는 ‘원재료 → 탄화층(반응 1)’과 ‘탄화층 산화(반응 2)’를 두 재료로 나눠 입력 —\nEPS는 원재료 한 개, 반응 한 개");
  badge(s, 1, MX, 1.85, "PyroSim에서", 4.5);
  text(s, paras([
    [["재료 편집 → 열물성: ", { bold: true }], ["밀도 · 비열(온도별 표 가능) · 열전도율 · 방사율 — 실측값", {}]],
    [["재료 편집 → 열분해 → 반응 추가: ", { bold: true }], ["반응 속도를 '기준온도 · 기준속도 · 승온속도'로 지정", {}]],
    [["생성물: ", { bold: true }], ["잔류물(탄화층 재료)과 수율, 기체(연료) 수율 = 1 − 잔류물 수율", {}]],
    [["반응열: ", { bold: true }], ["흡열 +, 발열 − (현재는 부호만 확인)", {}]],
    [["탄화층 산화: ", { bold: true }], ["탄화층 재료에 반응 추가, 산소 차수 1, A · E 직접 입력", {}]],
    [["표면(SURF): ", { bold: true }], ["면재 · 마감층 + 단열재 100 mm 층 구성 (16쪽)", {}]],
  ], { space: 7 }), { x: MX, y: 2.3, w: 5.35, h: 4.4, fontSize: 12 });
  badge(s, 2, 6.1, 1.85, "FDS 입력 예시 — PIR(PUR) 준불연", 6);
  const r1 = R("PIR_SC", ATM_N2, 1), r2 = R("PIR_SC", ATM_AIR, 2);
  const code = [
    "&MATL ID='PIR_SC'",
    "      DENSITY=___, SPECIFIC_HEAT=___, CONDUCTIVITY=___",
    "      EMISSIVITY=0.9, N_REACTIONS=1",
    `      REFERENCE_TEMPERATURE(1)=${fmt(r1.REFERENCE_TEMPERATURE_C, 0)}.`,
    `      REFERENCE_RATE(1)=${Number(r1.REFERENCE_RATE_per_s).toExponential(2).toUpperCase()}, HEATING_RATE(1)=5.`,
    "      MATL_ID(1,1)='PIR_SC_CHAR', NU_MATL(1,1)=___",
    "      SPEC_ID(1,1)='___', NU_SPEC(1,1)=___",
    "      HEAT_OF_REACTION(1)=___ /",
    "&MATL ID='PIR_SC_CHAR'",
    "      DENSITY=___, SPECIFIC_HEAT=___, CONDUCTIVITY=___",
    "      N_REACTIONS=1, N_O2(1)=1.",
    `      A(1)=${Number(r2.A_per_s_if_N_O2_1_air).toExponential(2).toUpperCase()}, E(1)=${Math.round(r2.E_kJ_mol * 1000)}.`,
    "      SPEC_ID(1,1)='___', NU_SPEC(1,1)=1.",
    "      HEAT_OF_REACTION(1)=___ /     ! 발열 = 음수",
  ];
  s.addText(code.map((l, i) => ({ text: l, options: { breakLine: i < code.length - 1 } })), {
    x: 6.1, y: 2.3, w: 6.73, h: 3.9, fontSize: 10.5, fontFace: "Consolas", color: C.text1, fill: { color: C.background2 }, margin: [0.1, 0.14, 0.1, 0.14], valign: "top", isTextBox: true, lang: "en-US" });
  text(s, "___ = 실측 · 기록 확인 · 문헌으로 채울 값. A · E는 탄화층 산화(N_O2 = 1)에만 직접 입력하고, 반응 1은 기준온도 · 기준속도로 둠",
    { x: 6.1, y: 6.3, w: 6.73, h: 0.55, fontSize: 10.5, color: C.text2 });
  s.addNotes([
    "[스크립트]",
    "- PyroSim의 재료 편집창은 FDS의 MATL 줄을 만듭니다. 열물성 탭에 밀도 · 비열 · 열전도율 · 방사율을, 열분해 탭에서 반응을 추가해 기준온도 · 기준속도 · 승온속도를 넣습니다.",
    "- PIR(PUR)과 PF는 재료를 둘로 나눕니다. 원재료는 반응 1로 탄화층과 연료 기체를 만들고, 탄화층 재료가 반응 2(산화)를 합니다. FDS는 '한 재료 성분당 한 반응'으로 기준온도 방식을 쓰기 때문입니다.",
    "- 탄화층 산화는 산소가 있어야 진행되므로 산소 차수(N_O2)를 1로 둡니다. 이때 FDS 반응속도에 산소 몰분율이 곱해지므로, 공기 중 TGA에서 구한 A를 0.209로 나눠 A · E를 직접 넣습니다.",
    "",
    "[질문 대비]",
    "- 연료 기체(SPEC_ID): 중공층 모델에서 쓰는 연소 반응(REAC)의 연료와 맞춥니다. 재료별 연소열은 MCC나 콘칼로리미터 값이 있어야 합니다.",
    "- 반응 1의 잔류물 수율이 없으면: 초기 질량 기록 확인이 가장 빠릅니다. 기록이 없으면 같은 조건 TGA 한 번(초기 질량 기록)으로 구할 수 있습니다.",
  ].join("\n"));
}

// ---------------------------------------------------------------- 15 density and specific heat
{
  const s = content("PyroSim 입력", "밀도 · 비열 · 열전도율을 구하는 방법",
    "밀도는 보드를 재고(KS M ISO 845), 비열은 같은 STA로 사파이어법 측정(ASTM E1269) —\n지금 DSC 기록은 시료 신호 0.1–0.35 mW가 기준선 이동 111–180 mW에 묻혀 계산 불가");
  const colW = (CW - 2 * 0.25) / 3;
  const blocks = [
    ["밀도 (DENSITY)", [
      "KS M ISO 845 (겉보기 밀도)",
      "시편 3개 이상, 부피 100 cm³ 이상 (예: 100 × 100 × 50 mm)",
      "23 °C · 50 %RH 상태 조절 → 치수 · 질량 측정",
      "ρ = m / V (kg/m³)",
      "면재가 붙은 제품은 심재와 전체를 따로",
      "엑셀 '밀도 측정' 시트에 입력 → 평균이 입력값으로",
    ]],
    ["비열 (SPECIFIC_HEAT)", [
      "ASTM E1269 / ISO 11357-4 (사파이어법)",
      "같은 장비 · 도가니 · 가스 · 승온(10–20 °C/min)으로 빈 도가니 → 사파이어 → 시료",
      "범위: 분해 시작 전(약 30–150 °C), 시료 질량 기록",
      "c_p = c_p,사파이어 × (m_사파이어 / m_시료) × (HF_시료 − HF_빈) / (HF_사파이어 − HF_빈)",
      "2–3회 반복, 온도별 값은 SPECIFIC_HEAT_RAMP로",
      "엑셀 '비열 측정' 시트에 입력",
    ]],
    ["열전도율 (CONDUCTIVITY)", [
      "제품 시험성적서의 열전도율 (KS L 9016 / ISO 8301)",
      "고온 값은 CONDUCTIVITY_RAMP — 문헌 또는 추가 시험",
      "TGA–DSC와 중공층 온도로는 계산 불가 (기존 연계 분석)",
      "카탈로그 범위 0.018–0.040 W/(m·K)",
    ]],
  ];
  blocks.forEach(([head, items], i) => {
    const x = MX + i * (colW + 0.25);
    badge(s, i + 1, x, 1.85, head, colW - 0.5);
    text(s, paras(items, { space: 5 }), { x, y: 2.3, w: colW, h: 2.95, fontSize: 11.5 });
  });
  s.addShape(pres.ShapeType.roundRect, { x: MX, y: 5.35, w: CW, h: 1.5, rectRadius: 0.08, fill: { color: C.background2 }, line: { color: C.background2, width: 0.5 }, objectName: "why not" });
  text(s, "지금 DSC 기록으로 비열을 못 구하는 이유", { x: MX + 0.25, y: 5.45, w: 6, h: 0.3, fontSize: 12.5, bold: true, valign: "middle" });
  const facts = [["0.1–0.35 mW", "시료 현열 신호 = c_p × 승온속도 × 질량\n≈ 1.4 J/(g·K) × 0.083 K/s × 1–3 mg (추정)"], ["111–180 mW", "열류 기준선 이동 (100 → 700 °C)\n블랭크 없이 보정 불가"], ["계수 1 · 기준선 0", "열류 교정 · 기준선 열이 비어 있음\n초기 질량도 기록 없음"]];
  facts.forEach(([big, sub], i) => {
    const x = MX + 0.25 + i * 4.0;
    text(s, big, { x, y: 5.78, w: 3.8, h: 0.4, fontSize: 17, bold: true, valign: "middle" });
    text(s, sub, { x, y: 6.2, w: 3.8, h: 0.6, fontSize: 10, color: C.text2, valign: "top" });
  });
  s.addNotes([
    "[스크립트]",
    "- 밀도는 TGA와 관계없이 보드에서 잽니다. KS M ISO 845 방식으로 시편 3개 이상의 치수와 질량을 재서 평균을 냅니다. 면재가 붙은 제품이면 심재와 전체를 따로 기록해 두면 PyroSim에서 층을 나눌 때 씁니다.",
    "- 비열은 같은 STA 장비로 잴 수 있습니다. 빈 도가니, 사파이어 기준시료, 시료를 같은 조건으로 차례로 측정해 식으로 계산합니다(ASTM E1269). 시료 질량을 반드시 기록하고, 분해가 시작되기 전 온도 범위에서만 구합니다.",
    "- 지금 기록으로 안 되는 이유: 시료가 1–3 mg이면 현열 신호가 0.1–0.35 mW인데, 기준선이 100 → 700 °C 사이 111–180 mW 움직입니다. 블랭크와 교정이 없어 이 차이를 걷어낼 수 없습니다.",
    "",
    "[질문 대비]",
    "- 시료 질량 1–3 mg의 근거: EPS처럼 잔류물이 거의 없는 재료의 총 질량감소가 0.16–2.5 mg이었으므로, 시료 질량은 대략 이 정도입니다(추정).",
    "- 탄화층의 비열: N₂ TGA를 800 °C까지 돌린 잔류물로 같은 방식으로 잴 수 있습니다(선택).",
    "- 열전도율은 TGA–DSC로도, 중공층 중앙 온도로도 구할 수 없다는 것을 기존 연계 분석에서 확인했습니다.",
  ].join("\n"));
}

// ---------------------------------------------------------------- 16 cautions
{
  const s = content("PyroSim 입력", "PyroSim 적용 시 주의와 검증",
    "맨 단열재 카탈로그 물성으로는 시험보다 훨씬 빨리 착화함 (16–51 s vs 188–288 s) —\n면재 · 마감층을 넣고 중공층 시험(착화 시점 · 3000 mm t₆₀₀)으로 맞춰야 함");
  const items = [
    ["착화가 너무 빠름", "기존 역산: 카탈로그 k · ρ · c로는 단열재 표면이 16–51 s에 분해온도 도달, 실제 준불연 재료 착화는 188–288 s → 면재 · 난연 처리 · 마감층을 모델에 넣고 시험으로 검증", "poor"],
    ["반응 입력 검증", "SURF에 TGA_ANALYSIS=T, TGA_HEATING_RATE=5 → CHID_tga.csv의 질량을 α로 바꿔 엑셀 '측정 곡선'과 비교, 피크가 맞을 때까지 REFERENCE_RATE 조정", "fair"],
    ["EPS 용융", "EPS는 분해 전에 녹아 수축 · 흘러내림 — FDS 고체 모델은 용융 흐름을 다루지 않음 → EPS 결과는 시험과 비교해 해석", "fair"],
    ["PF 반응식", "N₂ 2단계가 790 °C에서 끝나지 않음, 공기 산화 속도는 열분해와 겹친 값 → 800 °C 등온 유지 측정 후 보완", "fair"],
    ["채울 값", "잔류물 수율 · 반응열 · 산화 A(가스 종류) · 밀도 · 비열 — 앞 자료의 필수 추가 실험 · 기록 확인과 같은 항목", "fair"],
  ];
  items.forEach(([head, body, st], i) => {
    const y = 1.85 + i * 0.98;
    chip(s, MX, y + 0.05, 1.0, st, st === "poor" ? "중요" : "확인");
    text(s, head, { x: MX + 1.2, y, w: 2.4, h: 0.42, fontSize: 13.5, bold: true, valign: "middle" });
    text(s, body, { x: MX + 3.65, y, w: CW - 3.65, h: 0.85, fontSize: 12, color: C.text2 });
  });
  s.addNotes([
    "[스크립트]",
    "- 가장 큰 함정은 착화 시점입니다. 기존 연계 분석에서 카탈로그 물성의 맨 단열재는 버너 점화 후 16–51 s면 분해온도에 도달했는데, 준불연 재료가 실제로 열을 보태기 시작한 것은 188–288 s였습니다. 시스템(면재, 난연 처리, 마감)의 저항이 큽니다.",
    "- 그래서 PyroSim 모델은 단열재만이 아니라 면재 · 마감층까지 층으로 넣고, 중공층 시험의 착화 시점과 3000 mm t₆₀₀으로 맞춰야 합니다.",
    "- 반응 입력은 FDS의 TGA_ANALYSIS 기능으로 확인합니다. 시뮬레이션 없이 몇 초 만에 가상 TGA 곡선을 만들어 주므로, 엑셀의 측정 곡선과 비교해 피크 온도 · 높이를 맞춥니다.",
    "",
    "[질문 대비]",
    "- TGA_ANALYSIS 결과의 질량은 초기 질량 기준(Y)이라, α = (1 − Y)/(1 − Y(790 °C))로 바꿔 비교합니다(엑셀 '측정 곡선' 시트 설명).",
    "- EPS 용융: FDS 고체 모델에는 녹아 흘러내리는 현상이 없어서, EPS 계열은 시험보다 오래 버티거나 다르게 타는 결과가 나올 수 있습니다.",
  ].join("\n"));
}

// ---------------------------------------------------------------- 17 appendix
{
  pres.addSection({ title: "부록" });
  const s = content("부록", "부록 — 재현 방법과 파일",
    "모든 숫자 · 그래프는 원자료 24개에서 스크립트로 다시 만들어지며,\nFDS 설명서의 기준온도 방식 식을 그대로 씀");
  badge(s, 1, MX, 1.85, "재현 (tga_dsc 폴더에서)", 6);
  s.addText([
    "python analysis.py              # 요약표",
    "python qa_reliability.py        # 자료 판정",
    "python pyrosim_inputs.py        # 반응 파라미터 · 측정 곡선",
    "python material_figures.py      # 그림 8장 (figures/material_*.png)",
    "python report/deck_data.py      # 덱 자료",
    "python report/build_pyrosim_xlsx.py  # 엑셀 입력표",
    "node report/build_pyrosim_deck.js    # 이 PPT",
  ].map((l, i, a) => ({ text: l, options: { breakLine: i < a.length - 1 } })), {
    x: MX, y: 2.3, w: 6.3, h: 2.3, fontSize: 10.5, fontFace: "Consolas", color: C.text1, fill: { color: C.background2 }, margin: [0.1, 0.14, 0.1, 0.14], valign: "top", isTextBox: true, lang: "en-US" });
  badge(s, 2, 7.15, 1.85, "근거", 5);
  text(s, paras([
    "FDS User's Guide — Estimating Kinetic Parameters: REFERENCE_TEMPERATURE · REFERENCE_RATE · HEATING_RATE, PYROLYSIS_RANGE",
    "FDS Source read.f90 — E = e · r_p · R · T_p² / β, A = e · r_p · exp(E / R T_p) (잔류물 보정 없음 → 반응 성분 기준 속도)",
    "FDS User's Guide — Thermal Properties (DENSITY, SPECIFIC_HEAT(_RAMP), CONDUCTIVITY(_RAMP)), TGA_ANALYSIS",
    "이 저장소 — coupled_analysis(착화 역산 16–51 s vs 188–288 s, 카탈로그 범위), qa_reliability(자료 판정)",
  ], { space: 6 }), { x: 7.15, y: 2.3, w: 5.68, h: 2.6, fontSize: 11 });
  badge(s, 3, MX, 4.95, "파일", 6);
  text(s, paras([
    "report/PyroSim_입력값_정리.xlsx — MATL 입력값 · 반응 파라미터(수식) · 밀도/비열 측정 기록표 · 측정 곡선",
    "results/pyrosim_reactions.csv · pyrosim_tga_curves.csv — 단계별 값, 비교용 곡선",
    "figures/material_tga.png · material_dsc.png · material_<재료>.png — 재료 특성 그림",
  ], { space: 5 }), { x: MX, y: 5.4, w: CW, h: 1.4, fontSize: 11.5 });
  s.addNotes([
    "[스크립트]",
    "- 원자료(data/N, data/O)를 넣고 위 순서대로 실행하면 이 PPT, 엑셀, 그림이 모두 다시 만들어집니다.",
    "- 반응 입력 식은 FDS 설명서와 소스 코드(read.f90)에서 확인한 것을 그대로 썼습니다. 소스에는 잔류물 보정이 없어서, 기준속도는 '반응하는 성분'의 전환 속도입니다.",
    "",
    "[질문 대비]",
    "- 원자료는 공개 저장소에 올리지 않습니다. 결과 CSV와 덱 자료(JSON)만 저장소에 있습니다.",
  ].join("\n"));
}

pres.writeFile({ fileName: OUT }).then(async (f) => {
  await finish(f);
  console.log("wrote", f);
});
