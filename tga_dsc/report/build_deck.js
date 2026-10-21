#!/usr/bin/env node
// TGA-DSC 종합분석 · 데이터 신뢰성 평가 덱 (pptxgenjs, native charts). Reads deck_data.json next to this file.
//
//   python ../qa_reliability.py && python deck_data.py
//   NODE_PATH=<folder holding pptxgenjs> node build_deck.js [out.pptx]
//
// Theme, layouts, helpers and the post-write repairs (finish) live in deck_lib.js.
"use strict";
const fs = require("fs");
const path = require("path");
const { makeDeck } = require("./deck_lib");

const D = JSON.parse(fs.readFileSync(path.join(__dirname, "deck_data.json"), "utf8"));
const OUT = process.argv[2] || path.join(__dirname, "TGA-DSC_종합분석_신뢰성평가.pptx");
const {
  pres, C, HEX, STATUS, W, MX, CW, CAT, LANG, fmt, tab, statusOf, richText,
  content, text, vlabel, table, paras, badge, legend, chip, axisStyle, plotMap, curves, refLine, finish,
} = makeDeck({
  title: "TGA–DSC 종합분석과 데이터 신뢰성 평가",
  subject: "단열재 6종 TGA–DSC 원자료 24개 — 분석 결과, 신뢰성 점검, 추가 실험 판단",
  footer: "KICT 중공층 외단열 · TGA–DSC 종합분석과 데이터 신뢰성 평가",
});

const ORDER = D.order;
const LABEL = D.label;
const ATM_N2 = "N₂", ATM_AIR = "공기(O)";

// ---------------------------------------------------------------- data shortcuts
const V = (m, atm) => tab(D.verdict, { material: m, atm });
const T = (m, atm) => tab(D.table, { material: m, atm });
const runs = (m, atm) => D.runs.filter((r) => r.material === m && r.atm === atm).sort((a, b) => a.rep - b.rep);
const EP = (m, atm, e) => tab(D.endpoint, { material: m, atm, endpoint: e });
const counts = { reliable: 0, conditional: 0, endpoint: 0, remeasure: 0 };
D.verdict.forEach((v) => (counts[v.overall] += 1));
const rmaxAir = Object.fromEntries(ORDER.map((m) => [m, T(m, ATM_AIR).r_max_pct_C]));
const rmaxN2 = Object.fromEntries(ORDER.map((m) => [m, T(m, ATM_N2).r_max_pct_C]));

// ---------------------------------------------------------------- 1 title
{
  pres.addSection({ title: "결론" });
  const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "결론" });
  s.addText("TGA–DSC 분석 자료 · 2026-10-19", { placeholder: "kicker", lang: LANG });
  s.addText("단열재 6종 TGA–DSC 종합분석과\n데이터 신뢰성 평가", { placeholder: "title", lang: LANG });
  s.addText(richText("EPS · PIR(PUR) 가연성/준불연, PF 준불연 외부용/내부용 — N₂ · 공기(O) 각 2회, 원자료 24개"), { placeholder: "subtitle", lang: LANG });
  const facts = [
    ["24", "시험 · 12 조합", "재료 6종 × 분위기 2 × 반복 2"],
    [`${counts.reliable} · ${counts.conditional} · ${counts.endpoint} · ${counts.remeasure}`, "신뢰 · 조건부 · 종점 의존 · 재측정", "조합별 TGA 판정 (6개 점검)"],
    ["10회 + 4건", "추가 실험 (필수) + 기록 확인", "표적 보완 — 전면 재시험 불필요"],
  ];
  facts.forEach(([big, lab, sub], i) => {
    const x = 0.8 + i * 3.95;
    s.addShape(pres.ShapeType.roundRect, { x, y: 4.55, w: 3.7, h: 1.65, rectRadius: 0.08, fill: { color: C.background1, transparency: 90 }, line: { color: C.background2, width: 0.75 } });
    text(s, big, { x: x + 0.25, y: 4.7, w: 3.3, h: 0.6, fontSize: 28, bold: true, color: C.background1 });
    text(s, lab, { x: x + 0.25, y: 5.32, w: 3.3, h: 0.32, fontSize: 13, bold: true, color: C.background1 });
    text(s, sub, { x: x + 0.25, y: 5.66, w: 3.3, h: 0.32, fontSize: 11, color: C.background2 });
  });
  text(s, "재현: tga_dsc/qa_reliability.py → report/deck_data.py → report/build_deck.js (원자료는 저장소에 없음)", { x: 0.8, y: 6.65, w: 11.7, h: 0.3, fontSize: 11, color: C.background2 });
  s.addNotes([
    "[스크립트]",
    "- TGA–DSC 원자료 24개(재료 6종 × N₂·공기(O) × 2회)를 다시 계산하고, 각 결과를 어디까지 믿을 수 있는지 점검한 분석 자료입니다.",
    "- 순서: 결론과 추가 실험 판단(2–3쪽) → 자료와 산정 식(4쪽) → 분석 결과(5–9쪽: 곡선, 지표, DSC, 중공층 연계) → 신뢰성 점검 4가지(10–13쪽) → 판정 매트릭스(14쪽) → 추가 실험 계획(15쪽) → 부록(16쪽).",
    "- 표지의 세 숫자가 결론입니다: 12개 조합 중 신뢰 3 · 조건부 3 · 종점 의존 5 · 재측정 1, 추가 실험은 필수 10회와 기록 확인 4건.",
    "",
    "[질문 대비]",
    "- 원자료 위치: tga_dsc/data/N, data/O (CSV 24개, 저장소에는 올리지 않음). 수치는 qa_reliability.py → deck_data.py → build_deck.js 순서로 다시 만들어집니다.",
    "- '공기(O)' 표기: 원자료의 분위기 표기는 O(산화 분위기)뿐이고, 공기인지 순수 산소인지는 기록이 없습니다. 기록 확인 항목에 넣었습니다.",
  ].join("\n"));
}

// ---------------------------------------------------------------- 2 conclusion cards
{
  const s = content("결론", "결론 — 데이터 신뢰성과 추가 실험 판단",
    "공기 중 PF 2종 · PIR(PUR) 가연성 자료는 그대로 쓸 수 있고,\nEPS 가연성 공기 자료는 재측정이 필요함 — 추가 실험은 표적 보완으로 충분");
  const cards = [
    ["reliable", "공기: PF 내부용 · PF 외부용\n공기: PIR(PUR) 가연성", "6개 점검 모두 양호",
      "질량감소 ≥ 1.26 mg, 반복 차이 온도 ≤ 3 °C · r_max ≤ 1.3 %, 순위 분리도 9.7–47"],
    ["conditional", "N₂: EPS 가연성 · EPS 준불연\n공기: PIR(PUR) 준불연", "‘보통’ 항목 있음 — 값에 반복 범위 병기",
      "EPS N₂: 질량감소 0.66–0.87 mg. PIR(PUR) 준불연 공기: 온도 반복 23 °C · r_max 19 % · 분리도 1.5"],
    ["endpoint", "N₂: PIR(PUR) 2종 · PF 2종\n공기: EPS 준불연", "790 °C에서도 분해 진행 (피크의 23–102 %)",
      "r_max · T_α50 · T_α90 · f₄₀₀이 정규화 종점에 따라 변함 → T_onset · T_peak 위주로 사용"],
    ["remeasure", "공기: EPS 가연성", "신호 부족 · 겉보기 증가",
      "질량감소 0.16 · 0.30 mg, 200 °C 이전 겉보기 증가 17 · 10 %, r_max 반복 차이 33 %"],
  ];
  const cw = (CW - 3 * 0.25) / 4;
  cards.forEach(([st, who, head, body], i) => {
    const S = STATUS[st], x = MX + i * (cw + 0.25), y = 1.85;
    s.addShape(pres.ShapeType.roundRect, { x, y, w: cw, h: 3.3, rectRadius: 0.08, fill: { color: S.tint }, line: { color: S.tint, width: 0.5 }, objectName: `card ${S.label}` });
    s.addShape(pres.ShapeType.ellipse, { x: x + 0.22, y: y + 0.29, w: 0.24, h: 0.24, fill: { color: S.color }, line: { color: S.color, width: 0.5 } });
    text(s, S.label, { x: x + 0.56, y: y + 0.2, w: 1.7, h: 0.42, fontSize: 19, bold: true, valign: "middle" });
    text(s, String(counts[st]), { x: x + cw - 1.0, y: y + 0.12, w: 0.8, h: 0.6, fontSize: 30, bold: true, align: "right", valign: "middle" });
    text(s, who, { x: x + 0.22, y: y + 0.8, w: cw - 0.4, h: 0.66, fontSize: 12.5, bold: true });
    text(s, head, { x: x + 0.22, y: y + 1.5, w: cw - 0.44, h: 0.62, fontSize: 13, color: C.text1 });
    text(s, body, { x: x + 0.22, y: y + 2.13, w: cw - 0.44, h: 1.05, fontSize: 11.5, color: C.text2 });
  });
  text(s, paras([
    [["DSC  ", { bold: true }], ["교정·기준선 보정이 없어 정성 자료 — 두 반복에서 재현된 사건은 2건 (PF 내부용 공기 474 °C 발열, EPS 가연성 N₂ 398 °C 흡열)", {}]],
    [["산정 불가  ", { bold: true }], ["잔류율(%) · J/g · 비열 — 질량 열에 절대 영점이 없고 블랭크(빈 도가니) 측정이 없음", {}]],
  ], { bullet: false, space: 4 }), { x: MX, y: 5.3, w: CW, h: 0.7, fontSize: 13 });
  s.addShape(pres.ShapeType.roundRect, { x: MX, y: 6.08, w: CW, h: 0.82, rectRadius: 0.08, fill: { color: C.text1 }, line: { color: C.text1, width: 0.5 }, objectName: "decision" });
  text(s, [
    { text: "추가 실험 — 필요 (표적 보완, 전면 재시험은 불필요)", options: { bold: true, fontSize: 15, breakLine: true } },
    { text: "필수 10회: 블랭크 2 · EPS 가연성 공기 재측정 3 · 공기 3번째 반복 5   |   기록 확인 4건: 초기 질량 · O 가스 종류 · 교정 · 시료 채취", options: { fontSize: 12.5 } },
  ], { x: MX + 0.25, y: 6.08, w: CW - 0.5, h: 0.82, color: C.background1, valign: "middle" });
  s.addNotes([
    "[스크립트]",
    "- 결론부터 말씀드리면, 12개 조합(6종 × 2분위기) 가운데 공기의 PF 2종과 PIR(PUR) 가연성 3개는 6개 점검(신호량, 겉보기 증가, 종점 잔여, 온도 반복, 속도 반복, 순위 분리)이 모두 양호해서 그대로 씁니다.",
    "- 조건부 3개는 신호량이나 반복 차이가 '보통'인 조합입니다. 값을 쓸 때 반복 범위를 함께 적으면 됩니다.",
    "- 종점 의존 5개는 790 °C에서도 분해가 끝나지 않은 조합입니다. 전환율을 총 질량감소로 나누므로, 분해가 덜 끝나면 r_max · T_α50 · T_α90 · f₄₀₀ 같은 정규화 지표가 어디서 끊느냐(종점)에 따라 달라집니다. T_onset과 T_peak는 정규화 상수에 영향을 받지 않습니다.",
    "- 재측정은 공기 중 EPS 가연성 하나입니다. 질량감소가 0.16 · 0.30 mg로 너무 작아서 부력에 의한 겉보기 증가가 최대 17 %까지 나옵니다.",
    "- DSC는 교정·기준선 보정 없이 기록돼 전 조합이 정성 자료이고, 두 반복에서 재현된 사건은 2건뿐입니다.",
    "- 그래서 추가 실험은 필요하지만 전면 재시험은 아닙니다. 필수 10회와 기록 확인 4건이면 재측정·조건부 판정이 풀립니다(15쪽).",
    "",
    "[질문 대비]",
    "- 판정 기준은 이 분석의 자체 기준이고 표준이 아닙니다(부록 16쪽). 기준을 조금 바꾸면 경계에 있는 조합(예: EPS 준불연 N₂의 질량감소 0.66 mg)만 바뀝니다.",
    "- 왜 전면 재시험이 아닌가: 공기 3종과 대부분의 온도 지표는 이미 기준을 충족하고, 문제가 특정 조합(EPS 가연성 공기)과 기록 누락(초기 질량, 블랭크)에 몰려 있기 때문입니다.",
    "- 종점 의존이 '데이터 불량'인가: 아닙니다. 측정은 정상이고, 800 °C에서 끝난 승온 프로그램이 N₂의 PF·PIR 분해를 다 담지 못한 것입니다. 800 °C 등온 유지를 붙이면 해소됩니다.",
  ].join("\n"));
}

// ---------------------------------------------------------------- 3 conclusions x reliability
{
  const s = content("결론", "결론 — 기존 분석 결론의 신뢰도",
    "기존 결론 7개 중 4개는 그대로 유효하고, 중공층 순서 일치(ρ = 1.0)와\nN₂ PF 동일성은 조건부, EPS 가연성 공기 값은 재측정 후 확정");
  const VERD = { ok: ["reliable", "유효"], cond: ["conditional", "조건부"], hold: ["remeasure", "보류"] };
  const rows = [
    ["PF 내부용은 공기 중 474 °C 부근에서 급속 전소 — 외부용 대비 r_max 2.9배, T_α90 117 °C 낮음, 같은 온도에 DSC 발열", "공기 PF 2종: 신뢰 · DSC 발열 재현 (474 °C)", "ok", ""],
    ["PIR(PUR) 가연성은 TGA로 설명되지 않는 예외 — 공기 r_max 0.75(중간)인데 중공층 30 s 만에 전체 화염", "공기 PIR(PUR) 가연성: 신뢰 (r_max 반복 차이 0.5 %, 분리도 47)", "ok", "예외는 자료 문제가 아님"],
    ["PIR(PUR)은 2단계 분해 — N₂ 316–328 °C 주분해, 공기 525–532 °C 탄화층 산화", "피크 온도 반복 차이 ≤ 7.5 °C (T_peak는 정규화에 무관)", "ok", ""],
    ["EPS 준불연은 가연성보다 주분해가 27 °C 늦고 최대속도는 절반 이하 (N₂)", "N₂ EPS 2종: 조건부 (4회 중 3회가 질량감소 1 mg 미만) — 차이는 반복 차이의 4배 이상", "ok", ""],
    ["준불연 4종: 공기 r_max 순서 = 중공층 확산 순서 (ρ = 1.0)", "EPS 준불연 공기: 종점 의존 · PIR(PUR) 준불연 공기: 조건부 · 점 4개", "cond", "700 °C 종점 ρ 0.8"],
    ["N₂에서는 두 PF의 열분해가 같음 (T_α50 546 · 544 °C)", "N₂ PF 2종: 종점 의존 (790 °C에서 피크의 53–102 %)", "cond", "800 °C 유지 후 확정"],
    ["EPS 가연성은 공기에서 주분해가 약 100 °C 낮아짐 (N₂ 397 → 공기 296 °C)", "공기 EPS 가연성: 재측정 (질량감소 0.16 · 0.30 mg) — 피크 온도는 반복 차이 5.5 °C", "hold", "재측정 후 확정"],
  ];
  const hdr = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 }, valign: "middle" } });
  const tbl = [[hdr("#"), hdr("기존 결론 (분석 README · 통합 평가식)"), hdr("근거 자료의 판정"), hdr("결론 판정")]];
  rows.forEach(([c1, c2, v, note], i) => {
    const [st, lab] = VERD[v];
    tbl.push([
      { text: String(i + 1), options: { align: "center", color: C.text2 } },
      { text: c1 },
      { text: c2, options: { color: C.text2 } },
      { text: [{ text: lab, options: { bold: true, breakLine: !!note } }, ...(note ? [{ text: note, options: { fontSize: 10.5, color: C.text2 } }] : [])], options: { fill: { color: STATUS[st].tint } } },
    ]);
  });
  table(s, tbl, { x: MX, y: 1.85, w: CW, colW: [0.45, 5.6, 4.38, 1.9], fontSize: 12, color: C.text1, valign: "middle",
    border: { type: "solid", pt: 0.75, color: "D1D5DB" }, margin: [0.05, 0.09, 0.05, 0.09], rowH: 0.38 });
  text(s, "판정: 유효 = 근거 자료가 신뢰이거나, 조건부라도 결론의 차이가 반복 차이보다 충분히 큼 · 조건부 = 근거 자료가 조건부/종점 의존 · 보류 = 근거 자료가 재측정 대상",
    { x: MX, y: 6.45, w: CW, h: 0.4, fontSize: 10.5, color: C.text2 });
  s.addNotes([
    "[스크립트]",
    "- 앞쪽 판정을 기존 결론(tga_dsc README, 통합 평가식)에 대입한 표입니다.",
    "- 1번 PF 내부용 급속 전소: 공기 PF 2종이 신뢰 판정이고 DSC 발열도 두 반복에서 같은 474 °C에 나타나 그대로 유효합니다.",
    "- 2번 PIR(PUR) 가연성 예외: 공기 PIR(PUR) 가연성은 반복 차이 0.5 %, 분리도 47로 가장 안정된 자료입니다. 그래서 '예외'는 측정 문제가 아니라 TGA가 재지 못하는 요인에서 나온 것입니다.",
    "- 3번 PIR(PUR) 2단계 분해, 4번 EPS 준불연 vs 가연성: 피크 온도 지표는 정규화와 무관하고 차이가 반복 차이보다 훨씬 커서 유효합니다.",
    "- 5번 중공층 순서 일치(ρ = 1.0): 근거인 EPS 준불연(종점 의존)과 PIR(PUR) 준불연(조건부)이 약해서 조건부입니다. 700 °C로 끊으면 ρ가 0.8로 떨어집니다(11쪽).",
    "- 6번 N₂ PF 동일성: 790 °C에서도 분해가 진행 중이라 800 °C 유지 측정 후 확정합니다.",
    "- 7번 EPS 가연성 공기: 피크 온도는 두 반복이 5.5 °C 차이로 맞지만, 신호가 0.16 · 0.30 mg라 재측정 후 확정합니다.",
    "",
    "[질문 대비]",
    "- 점 4개에서 Spearman ρ = 1.0이 우연히 나올 확률은 1/24 ≈ 4 %입니다(4개 순서의 경우의 수 24). 5종(PIR(PUR) 가연성 포함)으로는 0.7입니다.",
    "- 4번이 '유효'인 이유: N₂ EPS 2종의 질량감소가 0.66–0.99 mg로 '보통'이지만, 가연성–준불연 차이(T_peak 27 °C, r_max 2.65 vs 1.15)가 반복 차이(분리도 4.1 · 16)보다 훨씬 큽니다.",
  ].join("\n"));
}

// ---------------------------------------------------------------- 4 data and formulas
{
  pres.addSection({ title: "분석 결과" });
  const s = content("분석 결과", "자료와 산정 식",
    "원자료에 질량 절대값과 블랭크가 없어,\n모든 TGA 지표를 ‘총 질량감소 기준 전환율 α’로 정규화함");
  badge(s, 1, MX, 1.85, "측정 조건");
  const kv = [
    ["장비", "PerkinElmer STA (TGA–DSC 동시측정)"],
    ["시료", "6종 × N₂ · 공기(O) × 2회 = 24개"],
    ["승온", "5 °C/min (실측 4.94–4.98), N₂ 20 → 800 °C, O 30 → 800 °C"],
    ["가스", "20 mL/min (실측 평균 19.8)"],
    ["기록", "0.125 s 간격, 승온 전 등온 약 5.1분"],
  ];
  table(s, kv.map(([k, v]) => [{ text: k, options: { bold: true, fill: { color: C.background2 } } }, { text: v }]),
    { x: MX, y: 2.27, w: 5.6, colW: [0.85, 4.75], fontSize: 12, color: C.text1, valign: "middle", border: { type: "solid", pt: 0.75, color: "D1D5DB" }, margin: [0.04, 0.09, 0.04, 0.09], rowH: 0.33 });
  badge(s, 2, MX, 4.22, "원자료의 열과 빠진 정보");
  const rec = [
    ["good", "있음", "시간, 프로그램 · 시료 온도, 미보정 질량, 미보정 열류, 가스 유량"],
    ["fair", "값 없음", "열류 교정 계수 = 1, 기준선 질량 · 열류 = 0 (보정 없이 내보냄)"],
    ["poor", "없음", "시료 초기 질량, 블랭크(빈 도가니) 측정, O 가스 종류, 교정 기록"],
  ];
  rec.forEach(([st, lab, t], i) => {
    const y = 4.66 + i * 0.68;
    chip(s, MX, y + 0.02, 1.05, st, lab);
    text(s, t, { x: MX + 1.2, y, w: 4.4, h: 0.6, fontSize: 12, valign: "top" });
  });
  badge(s, 3, 6.35, 1.85, "산정 식");
  const F = [
    ["전환율", "α = (W_ref − W) / (W_ref − W_end)", "W_ref 승온 시작, W_end 790 °C 부근 평균 → 총 감소 중 진행된 몫"],
    ["DTG", "r = dα/dT (10 °C 평활)", "r_max = 100–750 °C 최대값 (%/°C), T_peak = 그 온도"],
    ["특성 온도", "T_α5 · T_α50 · T_α90", "α = 0.05 · 0.5 · 0.9인 온도 (초기 질량 기준 T5%와 다름)"],
    ["시작 온도", "T_onset = T_peak − [α(T_peak) − α(T_lo)] / r_max", "주 피크 접선 외삽 — α와 r이 같은 비율로 변해 정규화에 무관"],
    ["고온 몫", "f₄₀₀ = 1 − α(400 °C)", "400 °C 이상 감소 비율 (탄화층 분해 · 산화)"],
    ["국소 열류", "h = H − SG₁₀₀(H)", "100 °C 창 추세 제거, 흡열 + → 사건의 온도 · 부호만"],
    ["순위 분리도", "S = 이웃 재료와의 차이 / 반복 차이", "≥ 3이면 반복해도 순서 유지, < 1.5면 구분 불가"],
    ["종점 민감도", "α를 700 · 750 · 790 °C 감소로 재정규화", "지표 변화 = 분해 미완의 영향"],
  ];
  const fh = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 } } });
  table(s, [[fh("지표"), fh("식"), fh("의미")], ...F.map(([a, b, c]) => [{ text: a, options: { bold: true } }, { text: b }, { text: c, options: { color: C.text2 } }])],
    { x: 6.35, y: 2.27, w: 6.48, colW: [1.12, 2.66, 2.7], fontSize: 11, color: C.text1, valign: "middle", border: { type: "solid", pt: 0.75, color: "D1D5DB" }, margin: [0.04, 0.07, 0.04, 0.07], rowH: 0.42 });
  s.addNotes([
    "[스크립트]",
    "- ① 측정 조건: 5 °C/min, N₂는 20 °C, O는 30 °C에서 800 °C까지 올렸고, 가스는 20 mL/min입니다. 실측 승온속도는 4.94–4.98 °C/min로 설정과 맞습니다.",
    "- ② 원자료에는 미보정 질량과 미보정 열류가 있습니다. 교정 계수 열은 1, 기준선 열은 0으로 비어 있어 보정 없이 내보낸 값입니다. 시료 초기 질량, 블랭크 측정, O 가스 종류, 교정 기록은 없습니다.",
    "- ③ 그래서 질량은 절대값 대신 '승온 시작부터 790 °C까지의 총 질량감소'를 1로 놓은 전환율 α로 다룹니다. 모든 TGA 지표는 α와 그 미분 r에서 나옵니다.",
    "- DSC는 기준선이 없어 100 °C 창의 추세를 빼서 사건의 온도와 부호만 봅니다.",
    "- 신뢰성 점검에 쓰는 두 식: 순위 분리도 S(이웃 재료와의 차이 ÷ 반복 차이)와 종점 민감도(700 · 750 · 790 °C로 다시 정규화).",
    "",
    "[질문 대비]",
    "- 왜 초기 질량 대비 %(T5% 등)가 아닌가: 질량 열의 시작값이 −1.22 ~ +2.70 mg로 절대 영점이 없어 초기 질량을 알 수 없기 때문입니다(10쪽).",
    "- T_α5와 T5%: T_α5는 총 질량감소의 5 %가 일어난 온도라서, 초기 질량의 5 %가 줄어든 온도(T5%)와 다릅니다. 혼용하면 안 됩니다.",
    "- T_onset이 정규화에 무관한 이유: α와 r이 모두 같은 상수(총 감소량)로 나뉘므로 [α(T_peak) − α(T_lo)] / r_max에서 상수가 지워집니다. T_peak도 최대값의 위치라 상수와 무관합니다.",
    "- 국소 열류의 한계: 추세를 빼는 필터 때문에 큰 발열 피크 양옆에 반대 부호의 곁봉우리가 생깁니다. 면적(J/g)은 의미가 없습니다.",
  ].join("\n"));
}

// ---------------------------------------------------------------- 5, 6 small multiples
function smallMultiples(s, key, opt) {
  const gx = 0.85, gw = W - MX - gx, colGap = 0.18, ch_h = 2.3, rowGap = 0.1, gy = 2.08;
  const cw = (gw - 2 * colGap) / 3;
  const L = { x: 0.12, y: 0.16, w: 0.84, h: 0.68 };
  ORDER.forEach((m, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const own = opt.own && opt.own[m];
    const ch = { x: gx + col * (cw + colGap), y: gy + row * (ch_h + rowGap), w: cw, h: ch_h, L, title: opt.title ? opt.title(m) : LABEL[m] + (own ? "  (세로축 다름)" : ""),
      xMin: 0, xMax: 800, xMajor: 200, ...(own || opt.y) };
    const c = D.curves[m];
    curves(s, ch, D.grid, [
      { name: "N₂ #1", ys: c["N2#1"][key], color: HEX.n2 },
      { name: "N₂ #2", ys: c["N2#2"][key], color: HEX.n2, style: { dash: "dash" } },
      { name: "공기(O) #1", ys: c["O#1"][key], color: HEX.air },
      { name: "공기(O) #2", ys: c["O#2"][key], color: HEX.air, style: { dash: "dash" } },
    ], { objectName: `${key} ${m}` });
  });
  vlabel(s, opt.ylab, 0.36, gy + ch_h + rowGap / 2, 4.6);
  text(s, "온도 (°C, 시료 온도)", { x: gx, y: 6.8, w: gw, h: 0.22, fontSize: 11, color: C.text2, align: "center", valign: "middle" });
  text(s, opt.caption, { x: MX, y: 1.72, w: 6.6, h: 0.3, fontSize: 11, color: C.text2, valign: "middle" });
  legend(s, [
    { label: "N₂ #1", color: HEX.n2, w: 0.55 }, { label: "N₂ #2", color: HEX.n2, dash: true, w: 0.55 },
    { label: "공기(O) #1", color: HEX.air, w: 0.95 }, { label: "공기(O) #2", color: HEX.air, dash: true, w: 0.95 },
  ], 7.3, 1.73);
}
{
  const s = content("분석 결과", "분석 ① 질량 잔류 곡선",
    "공기에서는 EPS의 주분해가 80–100 °C 앞당겨지고,\nPIR(PUR) · PF는 400 °C 이후 탄화층 산화 단계가 더해짐");
  smallMultiples(s, "rem", {
    y: { yMin: 0, yMax: 110, yMajor: 20 }, own: { EPS_C: { yMin: -20, yMax: 120, yMajor: 20 } },
    ylab: "잔류 질량 100·(1 − α) (%)", caption: "승온 시작 ~ 790 °C 질량감소 기준 — 초기 질량 대비 %가 아님",
  });
  s.addNotes([
    "[스크립트]",
    "- 여섯 칸은 재료별로 N₂(파랑)와 공기(O)(주황)를 두 반복씩 겹친 것입니다. 실선이 1회, 점선이 2회입니다.",
    "- EPS 2종은 한 단계로 거의 전부 사라집니다. 공기에서는 주분해 피크가 가연성 397 → 296 °C, 준불연 425 → 343 °C로 80–100 °C 낮아집니다.",
    "- PIR(PUR) 2종은 300 °C 부근의 1단계 뒤 N₂에서는 800 °C까지 천천히 줄고, 공기에서는 525–532 °C에서 탄화층이 산화되며 끝납니다.",
    "- PF 2종은 N₂에서 800 °C까지 계속 줄어 분해가 끝나지 않고, 공기에서는 내부용이 474 °C 부근에서 급격히 끝나는 반면 외부용은 590 °C 부근까지 이어집니다.",
    "- EPS 가연성 칸만 세로축이 −20 ~ 120 %입니다. 공기 시료의 질량감소가 0.16 · 0.30 mg로 작아 겉보기 증가(100 % 초과)가 크게 보이기 때문입니다.",
    "",
    "[질문 대비]",
    "- 100 %를 넘는 구간: 질량이 늘어난 것이 아니라 부력(가스 밀도 변화)이 보정되지 않은 겉보기 증가입니다. 신호가 작을수록 상대적으로 커집니다(10쪽).",
    "- 0 %가 '재가 없다'는 뜻인가: 아닙니다. 790 °C까지의 감소량을 100으로 놓았으므로 0 %는 '790 °C 시점'일 뿐이고, 실제 잔류율은 초기 질량이 있어야 계산됩니다.",
  ].join("\n"));
}
{
  const s = content("분석 결과", "분석 ② 질량감소속도 (DTG)",
    `공기 중 최대 질량감소속도 순서: EPS 가연성 ≫ PF 내부용 > EPS 준불연\n> PIR(PUR) 가연성 > PIR(PUR) 준불연 > PF 외부용 (${fmt(rmaxAir.EPS_C)} → ${fmt(rmaxAir.PF_out)} %/°C)`);
  smallMultiples(s, "dtg", {
    y: { yMin: -0.25, yMax: 1.25, yMajor: 0.25 }, own: { EPS_C: { yMin: -0.5, yMax: 3.0, yMajor: 0.5 } },
    ylab: "dα/dT (%/°C)", caption: "칸 제목 = r_max 2회 평균 (공기 · 질소) — EPS 가연성만 세로축 −0.5 ~ 3.0",
    title: (m) => `${LABEL[m]}  ${fmt(rmaxAir[m])} · ${fmt(rmaxN2[m])}`,
  });
  s.addNotes([
    "[스크립트]",
    "- DTG는 총 질량감소 중 1 °C마다 사라지는 몫입니다. 피크 높이가 r_max, 피크 위치가 T_peak입니다.",
    `- 공기 기준 순서: EPS 가연성 ${fmt(rmaxAir.EPS_C)} ≫ PF 내부용 ${fmt(rmaxAir.PF_in)} > EPS 준불연 ${fmt(rmaxAir.EPS_SC)} > PIR(PUR) 가연성 ${fmt(rmaxAir.PIR_C)} > PIR(PUR) 준불연 ${fmt(rmaxAir.PIR_SC)} > PF 외부용 ${fmt(rmaxAir.PF_out)} %/°C.`,
    "- EPS 가연성은 다른 재료보다 2–6배 빨라 세로축을 따로 잡았습니다(0 ~ 3.0). 나머지 다섯 칸은 같은 축(−0.25 ~ 1.25)이라 높이를 그대로 비교할 수 있습니다.",
    "- PF 내부용은 공기에서 474 °C의 좁고 높은 피크 하나로 끝나고, PF 외부용은 넓고 낮은 피크입니다. N₂에서는 두 PF 모두 800 °C를 향해 계속 올라가는 꼬리만 있습니다.",
    "",
    "[질문 대비]",
    "- 평활 폭(10 °C)이 피크를 깎는가: 평활 없이 10 °C 중앙차분으로 다시 구하면 r_max가 0.7–3.8 % 낮게 나와 순서는 같습니다(부록).",
    "- 음의 DTG: 질량이 잠깐 늘어난 구간(부력, 잡음)입니다. EPS 가연성 공기에서 가장 큽니다.",
  ].join("\n"));
}

// ---------------------------------------------------------------- 7 key metrics
{
  const s = content("분석 결과", "분석 ③ 핵심 지표",
    `N₂에서는 두 PF의 r_max가 같지만(${fmt(rmaxN2.PF_in)} · ${fmt(rmaxN2.PF_out)}),\n공기에서는 PF 내부용만 ${fmt(rmaxAir.PF_in / rmaxN2.PF_in, 1)}배로 커져 PF 외부용의 ${fmt(rmaxAir.PF_in / rmaxAir.PF_out, 1)}배가 됨`);
  const hd = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 }, align: "center" } });
  const rows = [[hd("재료"), hd("분위기"), hd("T_onset"), hd("T_peak"), hd("r_max"), hd("T_α50"), hd("T_α90"), hd("f₄₀₀")]];
  ORDER.forEach((m) => [ATM_N2, ATM_AIR].forEach((atm, k) => {
    const t = T(m, atm), ov = V(m, atm).overall;
    const dep = ov === "endpoint", all = ov === "remeasure";
    const cell = (v, n, grey, txt) => ({ text: txt || fmt(v, n), options: { align: "right", color: grey || all ? "9CA3AF" : C.text1 } });
    const peakTxt = m === "PF_out" && atm === ATM_N2 ? `(${fmt(t.T_peak_C, 0)})*` : null;
    rows.push([
      k === 0 ? { text: LABEL[m], options: { rowspan: 2, bold: true, valign: "middle" } } : null,
      { text: atm === ATM_N2 ? "N₂" : "공기", options: { align: "center", color: all ? "9CA3AF" : C.text1 } },
      cell(t.T_onset_C, 0), cell(t.T_peak_C, 0, peakTxt !== null, peakTxt), cell(t.r_max_pct_C, 2, dep),
      cell(t.T_a50_C, 0, dep), cell(t.T_a90_C, 0, dep), cell(t.f_above400, 2, dep),
    ].filter((c) => c !== null));
  }));
  table(s, rows, { x: MX, y: 1.85, w: 7.7, colW: [1.75, 0.75, 0.88, 0.88, 0.86, 0.86, 0.86, 0.86], fontSize: 11, color: C.text1, valign: "middle",
    border: { type: "solid", pt: 0.75, color: "D1D5DB" }, margin: [0.03, 0.08, 0.03, 0.08], rowH: 0.315 });
  text(s, "2회 평균, 온도 °C · r_max %/°C. 회색 = 종점 의존 지표(790 °C 정규화 값)와 재측정 대상(EPS 가연성 공기). (745)* = 탐색 상한 750 °C 부근 값으로 실제 피크가 아님",
    { x: MX, y: 6.08, w: 7.7, h: 0.62, fontSize: 10.5, color: C.text2 });
  // r_max N2 vs air
  const ch = { x: 8.55, y: 2.36, w: 4.28, h: 3.7 };
  text(s, "최대 질량감소속도 r_max (%/°C)", { x: 8.55, y: 1.8, w: 4.28, h: 0.3, fontSize: 11, bold: true, valign: "middle" });
  legend(s, [{ kind: "box", label: "N₂", color: HEX.n2, w: 0.4 }, { kind: "box", label: "공기(O)", color: HEX.air, w: 0.8 }], 8.55, 2.08);
  s.addChart(pres.ChartType.bar, [
    { name: "N₂", labels: ORDER.map((m) => CAT[m]), values: ORDER.map((m) => rmaxN2[m]) },
    { name: "공기(O)", labels: ORDER.map((m) => CAT[m]), values: ORDER.map((m) => rmaxAir[m]) },
  ], {
    x: ch.x, y: ch.y, w: ch.w, h: ch.h, ...axisStyle(), barDir: "bar", barGrouping: "clustered", barGapWidthPct: 60,
    catAxisOrientation: "maxMin", chartColors: [HEX.n2, HEX.air],
    showTitle: false, showLegend: false,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.00",
    valAxisMinVal: 0, valAxisMaxVal: 3.2, valAxisMajorUnit: 0.5, valAxisLabelFormatCode: "0.0", objectName: "r_max N2 vs air",
  });
  const rr = (m) => fmt(rmaxAir[m] / rmaxN2[m]);
  text(s, paras([
    [["공기 ÷ N₂   ", { bold: true, color: C.text1 }], [`EPS 가연성 ${rr("EPS_C")} · 준불연 ${rr("EPS_SC")}`, {}]],
    `PIR(PUR) 가연성 ${rr("PIR_C")} · 준불연 ${rr("PIR_SC")}`,
    [[`PF 외부용 ${rr("PF_out")} · `, {}], [`내부용 ${rr("PF_in")}`, { bold: true, color: C.text1 }]],
  ], { bullet: false, space: 1 }), { x: 8.55, y: 6.12, w: 4.28, h: 0.8, fontSize: 10.5, color: C.text2 });
  s.addNotes([
    "[스크립트]",
    "- 왼쪽 표는 12개 조합의 핵심 지표(2회 평균)입니다. 회색 글자는 판정상 종점 의존(r_max · T_α50 · T_α90 · f₄₀₀)이거나 재측정 대상인 값입니다.",
    `- 오른쪽은 r_max의 N₂ · 공기 비교입니다. 공기 ÷ N₂ 비는 EPS 2종 0.9 정도, PIR(PUR) 2종 1.1 정도로 분위기 영향이 작고, PF 외부용 ${fmt(rmaxAir.PF_out / rmaxN2.PF_out)}, PF 내부용 ${fmt(rmaxAir.PF_in / rmaxN2.PF_in)}입니다.`,
    "- 즉 PF 내부용과 외부용은 N₂에서는 같은 재료처럼 분해하지만(분리도 0.9), 공기에서는 내부용만 탄화층이 빠르게 산화됩니다. 중공층에서 PF 중 내부용만 III등급이 나온 것과 같은 방향입니다.",
    "",
    "[질문 대비]",
    "- PF N₂의 r_max는 진짜 피크인가: 아닙니다. 790 °C에서도 감소속도가 피크의 53–102 %라 '750 °C 이하에서 찾은 최대값'입니다. PF 외부용 N₂의 T_peak(745 °C)는 탐색 상한에 걸린 값입니다.",
    "- 반복 범위(±)는 results/tga_summary.csv의 _halfrange 열에 있습니다. 공기 PF 2종 · PIR(PUR) 가연성은 r_max ±0.01 이하입니다.",
  ].join("\n"));
}

// ---------------------------------------------------------------- 8 DSC
{
  const s = content("분석 결과", "분석 ④ DSC (정성)",
    "DSC에서 두 반복이 같은 온도 · 부호로 재현된 사건은 2건뿐 —\nPF 내부용 공기 474 °C 발열과 EPS 가연성 N₂ 398 °C 흡열");
  const gd = D.grid_dsc, din = D.dsc.PF_in, dout = D.dsc.PF_out;
  const ch = { x: MX, y: 2.08, w: 7.7, h: 4.6, L: { x: 0.1, y: 0.05, w: 0.86, h: 0.84 }, xMin: 100, xMax: 750, xMajor: 100, yMin: -20, yMax: 15, yMajor: 5 };
  curves(s, ch, gd, [
    { name: "PF 외부용 공기 #1", ys: dout["O#1"], color: HEX.grey },
    { name: "PF 외부용 공기 #2", ys: dout["O#2"], color: HEX.grey, style: { dash: "dash" } },
    { name: "PF 내부용 공기 #1", ys: din["O#1"], color: HEX.air },
    { name: "PF 내부용 공기 #2", ys: din["O#2"], color: HEX.air, style: { dash: "dash" } },
  ], { objectName: "DSC PF air" });
  legend(s, [
    { label: "PF 내부용 공기 #1", color: HEX.air, w: 1.45 }, { label: "#2", color: HEX.air, dash: true, w: 0.3 },
    { label: "PF 외부용 공기 #1", color: HEX.grey, w: 1.45 }, { label: "#2", color: HEX.grey, dash: true, w: 0.3 },
  ], MX, 1.72);
  const m = plotMap(ch);
  const r1 = runs("PF_in", ATM_AIR);
  text(s, [
    { text: `474 °C 발열  ${fmt(r1[0].dsc_event_mW, 1)} · ${fmt(r1[1].dsc_event_mW, 1)} mW`, options: { bold: true, breakLine: true } },
    { text: `크기 차이 2배 = 총 질량감소 ${fmt(r1[0].dm_tot_mg)} · ${fmt(r1[1].dm_tot_mg)} mg와 같은 비율 (시료량 차이)` },
  ], { x: m.X(492), y: m.Y(-11.5), w: 3.1, h: 0.75, fontSize: 11, color: C.text1 });
  text(s, "국소 열류 = 미보정 열류 − 100 °C 창 추세 (흡열 +). 큰 발열 양옆의 + 곁봉우리는 필터 효과",
    { x: m.px, y: m.py + m.ph + 0.28, w: m.pw, h: 0.22, fontSize: 10, color: C.text2, align: "center" });
  vlabel(s, "국소 열류 (mW, 흡열 +)", 0.3, m.py + m.ph / 2, 3.6);
  // EPS_C N2
  const ce = { x: 8.55, y: 2.12, w: 4.28, h: 2.3, L: { x: 0.12, y: 0.05, w: 0.84, h: 0.78 }, xMin: 300, xMax: 500, xMajor: 50, yMin: -1, yMax: 1.5, yMajor: 0.5 };
  text(s, "EPS 가연성 N₂ #1 · #2 (300–500 °C)", { x: 8.55, y: 1.8, w: 4.28, h: 0.3, fontSize: 11, bold: true, valign: "middle" });
  const sel = gd.map((t, i) => (t >= 300 && t <= 500 ? i : -1)).filter((i) => i >= 0);
  curves(s, ce, sel.map((i) => gd[i]), [
    { name: "EPS 가연성 N₂ #1", ys: sel.map((i) => D.dsc.EPS_C["N2#1"][i]), color: HEX.n2 },
    { name: "EPS 가연성 N₂ #2", ys: sel.map((i) => D.dsc.EPS_C["N2#2"][i]), color: HEX.n2, style: { dash: "dash" } },
  ], { objectName: "DSC EPS_C N2" });
  const re = runs("EPS_C", ATM_N2), me = plotMap(ce);
  text(s, `398 °C 흡열\n+${fmt(re[0].dsc_event_mW, 1)} · +${fmt(re[1].dsc_event_mW, 1)} mW`, { x: me.X(410), y: me.Y(1.4), w: 1.6, h: 0.45, fontSize: 10.5, bold: true });
  const others = D.verdict.filter((v) => v.v_dsc !== "fair").map((v) => `${v.label.replace("PIR(PUR) ", "PIR ")} ${v.atm === ATM_N2 ? "N₂" : "공기"}`);
  text(s, paras([
    [["재현 기준  ", { bold: true }], ["두 반복의 최대 사건이 150–720 °C에서 온도 차 ≤ 10 °C, 같은 부호, 크기 ≥ 0.5 mW", {}]],
    [[`비재현 ${others.length}개 조합  `, { bold: true }], ["온도 차가 크거나 크기가 0.5 mW 미만 → 판정 근거로 쓰지 않음", {}]],
    [["정량 불가  ", { bold: true }], ["열류 교정 계수 1 · 기준선 0, 기준선 이동 111–180 mW (13쪽)", {}]],
  ], { space: 6 }), { x: 8.55, y: 4.62, w: 4.28, h: 2.2, fontSize: 12 });
  s.addNotes([
    "[스크립트]",
    "- DSC는 교정·기준선 보정이 없어서, 100 °C 창의 추세를 뺀 '국소 열류'로 사건의 온도와 부호만 봅니다. 흡열이 +입니다(EPS 해중합이 N₂에서 +로 나오는 것으로 확인).",
    `- 왼쪽: 공기에서 PF 내부용은 두 반복 모두 474 °C에 큰 발열(${fmt(r1[0].dsc_event_mW, 1)} · ${fmt(r1[1].dsc_event_mW, 1)} mW)이 있고, 이 온도는 질량감소 피크(472 °C)와 같습니다. PF 외부용은 같은 범위에 뚜렷한 사건이 없습니다.`,
    "- 발열 크기가 두 배 다른 것은 시료량 차이로 설명됩니다: 총 질량감소가 1.27 · 2.81 mg로 같은 비율입니다.",
    `- 오른쪽 위: EPS 가연성 N₂의 398 °C 흡열(+${fmt(re[0].dsc_event_mW, 1)} · +${fmt(re[1].dsc_event_mW, 1)} mW)도 두 반복에서 재현됩니다. 해중합 피크(397 °C)와 같은 온도입니다.`,
    `- 나머지 ${others.length}개 조합은 최대 사건의 온도가 반복 간 다르거나 0.5 mW 미만이라 판정 근거로 쓰지 않습니다: ${others.join(", ")}.`,
    "",
    "[질문 대비]",
    "- J/g(엔탈피)를 낼 수 있나: 아닙니다. 기준선이 100 → 700 °C 사이 111–180 mW 움직이고(가장 큰 사건의 6배 이상), 초기 질량도 없습니다. 블랭크 측정과 초기 질량, 열류 교정(인듐 등)이 있어야 합니다.",
    "- 곁봉우리(+14 mW 등): 추세 제거 필터가 큰 발열 양옆에 만드는 반대 부호 신호입니다. 흡열 사건이 아닙니다.",
  ].join("\n"));
}

// ---------------------------------------------------------------- 9 cavity link
{
  const s = content("분석 결과", "분석 ⑤ 중공층 화재시험 연계",
    "준불연 4종은 공기 r_max가 클수록 3000 mm 도달이 빠름(ρ = 1.0)\n— 단 점이 4개이고 700 °C 종점에서는 ρ = 0.8, PIR(PUR) 가연성은 예외");
  const cav = Object.fromEntries(D.cavity.map((c) => [c.material, c]));
  const sc = ["PF_in", "EPS_SC", "PIR_SC", "PF_out"];
  const fitX = []; for (let r = 0.35; r <= 1.301; r += 0.05) fitX.push(Number(r.toFixed(2)));
  const xs = [...sc.map((k) => cav[k].r_max_air), cav.PIR_C.r_max_air, ...fitX];
  const n = xs.length, blank = () => new Array(n).fill(null);
  const ySc = blank(), yC = blank(), yFit = blank();
  sc.forEach((k, i) => (ySc[i] = cav[k].t600));
  yC[4] = cav.PIR_C.t600;
  fitX.forEach((r, i) => (yFit[5 + i] = Number((248 + 423 / r).toFixed(1))));
  const ch = { x: MX, y: 2.1, w: 7.0, h: 4.65, L: { x: 0.13, y: 0.03, w: 0.83, h: 0.83 }, xMin: 0.3, xMax: 1.3, xMajor: 0.1, yMin: 0, yMax: 1500, yMajor: 250 };
  curves(s, ch, xs, [
    { name: "통합 평가식", ys: yFit, color: "B8BCC4", style: { noMarker: true } },
    { name: "준불연 4종", ys: ySc, color: HEX.air, style: { noLine: true } },
    { name: "가연성 PIR(PUR)", ys: yC, color: HEX.ink, style: { noLine: true } },
  ], { lineDataSymbol: "circle", lineDataSymbolSize: 10, objectName: "cavity link" });
  const m = plotMap(ch);
  // label boxes: [dx, dy, width, align] from the point
  const lab = { PF_in: [-2.25, 0.08, 2.1, "right", "PF 내부용 514 s"], EPS_SC: [0.14, -0.36, 2.0, "left", "EPS 준불연 975 s"],
    PIR_SC: [0.14, -0.36, 2.4, "left", "PIR(PUR) 준불연 1080 s"], PF_out: [0.16, -0.13, 2.0, "left", "PF 외부용 1350 s"] };
  Object.entries(lab).forEach(([k, [dx, dy, w, al, t]]) => text(s, t, { x: m.X(cav[k].r_max_air) + dx, y: m.Y(cav[k].t600) + dy, w, h: 0.26, fontSize: 11, bold: true, align: al }));
  text(s, "PIR(PUR) 가연성 30 s — 예외", { x: m.X(cav.PIR_C.r_max_air) + 0.14, y: m.Y(cav.PIR_C.t600) - 0.36, w: 2.6, h: 0.26, fontSize: 11, bold: true });
  legend(s, [{ kind: "dot", label: "준불연 4종", color: HEX.air, w: 0.95 }, { kind: "dot", label: "PIR(PUR) 가연성", color: HEX.ink, w: 1.35 },
    { label: "통합 평가식 t₆₀₀ = 248 + 423 / r_max", color: "B8BCC4", w: 2.75 }], MX, 1.72);
  text(s, "공기 r_max (%/°C, 790 °C 정규화)", { x: m.px, y: m.py + m.ph + 0.28, w: m.pw, h: 0.24, fontSize: 11, color: C.text2, align: "center" });
  vlabel(s, "3000 mm 600 K 도달 t₆₀₀ (s, 중앙값)", 0.2, m.py + m.ph / 2, 3.8);
  badge(s, 1, 7.85, 1.95, "종점별 순위상관 (Spearman ρ)", 4.5);
  const rho = (chk, e) => tab(D.spearman, { check: chk, endpoint: `endpoint ${e} °C` }).rho;
  const k4 = "Spearman r_max(air) vs spread speed, sc 4 materials", k5 = "Spearman r_max(air) vs spread speed, 5 materials incl. PIR_C";
  const hh = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 }, align: "center" } });
  const rc = (v, hi) => ({ text: fmt(v, 1), options: { align: "center", bold: hi, fill: { color: hi ? STATUS.fair.tint : "FFFFFF" } } });
  table(s, [
    [hh("대상"), hh("700 °C"), hh("750 °C"), hh("790 °C")],
    [{ text: "준불연 4종" }, rc(rho(k4, 700), true), rc(rho(k4, 750)), rc(rho(k4, 790))],
    [{ text: "PIR(PUR) 가연성 포함" }, rc(rho(k5, 700)), rc(rho(k5, 750)), rc(rho(k5, 790))],
  ], { x: 7.85, y: 2.4, w: 4.98, colW: [2.28, 0.9, 0.9, 0.9], fontSize: 12, color: C.text1, valign: "middle", border: { type: "solid", pt: 0.75, color: "D1D5DB" }, margin: [0.04, 0.08, 0.04, 0.08], rowH: 0.36 });
  const e700 = (k) => EP(k, ATM_AIR, 700).r_max_pct_C;
  text(s, paras([
    `700 °C로 끊으면 EPS 준불연 ${fmt(e700("EPS_SC"))} > PF 내부용 ${fmt(e700("PF_in"))}로 1 · 2위가 바뀌어 ρ 0.8 (11쪽)`,
    "점 4개에서 ρ = 1.0이 우연히 나올 확률 1/24 ≈ 4 % — 공기 3번째 반복으로 확정 필요",
    "PIR(PUR) 가연성의 공기 자료는 ‘신뢰’(분리도 47) → 예외는 자료 문제가 아니라 TGA가 재지 못하는 요인",
    "곡선 = 통합 평가식 t₆₀₀ = 248 s + 423 / r_max (준불연 4종에 맞춘 식, integrated_model)",
  ], { space: 7 }), { x: 7.85, y: 3.75, w: 4.98, h: 3.0, fontSize: 12.5 });
  s.addNotes([
    "[스크립트]",
    "- 가로축은 공기 중 r_max, 세로축은 중공층 시험에서 3000 mm 높이가 600 K(상승 600 °C)에 닿은 시간 t₆₀₀의 중앙값입니다.",
    "- 준불연 4종(주황)은 r_max가 클수록 t₆₀₀가 짧습니다: PF 내부용 514 s < EPS 준불연 975 s < PIR(PUR) 준불연 1080 s < PF 외부용 1350 s. 순위상관 ρ = 1.0입니다.",
    "- 그러나 이 관계는 조건부입니다. 1 · 2위(PF 내부용 vs EPS 준불연)의 r_max 차이가 반복 차이의 1.8배뿐이고, EPS 준불연 공기는 790 °C에서도 분해 중이라 700 °C로 끊으면 순서가 바뀌어 ρ가 0.8이 됩니다.",
    "- 가연성 PIR(PUR)(검정)은 r_max 0.75로 중간인데 30 s 만에 전체 화염이 됐습니다. 공기 PIR(PUR) 가연성 자료는 가장 안정적인 자료(분리도 47)라서, 이 예외는 측정 문제가 아닙니다.",
    "",
    "[질문 대비]",
    "- 회색 곡선은 통합 평가식(integrated_model)입니다. 같은 4점으로 맞춘 식이라 이 그림이 식의 독립 검증은 아닙니다.",
    "- 종점을 바꿔도 PF 외부용 < PIR(PUR) 준불연 순서(아래쪽 두 점)는 유지되고, 바뀌는 것은 위쪽 두 점의 순서입니다.",
    "- ρ = 1.0의 우연 확률 4 %는 단측 기준입니다. 점이 4개라 통계적 결론으로는 약하고, 메커니즘 설명(산화 취약성)과 함께 써야 합니다.",
  ].join("\n"));
}

// ---------------------------------------------------------------- 10 reliability: mass signal
{
  pres.addSection({ title: "신뢰성 평가" });
  const nOk = D.runs.filter((r) => r.dm_tot_mg >= 1.0).length;
  const okMin = Math.min(...D.runs.filter((r) => r.dm_tot_mg >= 1.0).map((r) => r.dm_tot_mg));
  const okMax = Math.max(...D.runs.map((r) => r.dm_tot_mg));
  const s = content("신뢰성 평가", "신뢰성 ① 질량 신호",
    `24개 중 ${nOk}개는 질량감소 ${fmt(okMin, 1)}–${fmt(okMax, 1)} mg로 충분하지만,\nEPS 가연성 공기 2개는 0.16 · 0.30 mg로 부족 — 질량 영점이 없어 잔류율은 산정 불가`);
  const ser = [["N₂ #1", ATM_N2, 1, HEX.n2], ["N₂ #2", ATM_N2, 2, HEX.n2Light], ["공기(O) #1", ATM_AIR, 1, HEX.air], ["공기(O) #2", ATM_AIR, 2, HEX.airLight]];
  const ch = { x: MX, y: 2.12, w: 7.95, h: 4.75, L: { x: 0.075, y: 0.03, w: 0.81, h: 0.82 }, yMin: 0, yMax: 3.0 };
  text(s, "총 질량감소 (mg, 승온 시작 ~ 790 °C)", { x: MX, y: 1.8, w: 3.4, h: 0.3, fontSize: 11, bold: true, valign: "middle" });
  legend(s, ser.map(([nm, , , c], i) => ({ kind: "box", label: i % 2 ? "#2" : nm, color: c, w: i % 2 ? 0.28 : nm.length > 6 ? 0.85 : 0.5 })), 3.75, 1.8);
  s.addChart(pres.ChartType.bar, ser.map(([nm, atm, rep]) => ({ name: nm, labels: ORDER.map((m) => CAT[m]), values: ORDER.map((m) => runs(m, atm)[rep - 1].dm_tot_mg) })), {
    x: ch.x, y: ch.y, w: ch.w, h: ch.h, ...axisStyle(), layout: { ...ch.L }, barDir: "col", barGrouping: "clustered", barGapWidthPct: 45,
    chartColors: ser.map((x) => x[3]), showTitle: false, showLegend: false,
    valAxisMinVal: 0, valAxisMaxVal: 3.0, valAxisMajorUnit: 0.5, valAxisLabelFormatCode: "0.0", objectName: "mass loss per run",
  });
  const m = plotMap(ch);
  refLine(s, m, 1.0, "양호\n≥ 1.0 mg", STATUS.good.color);
  refLine(s, m, 0.5, "미흡\n< 0.5 mg", STATUS.poor.color);
  const st = D.stats;
  const callouts = [
    [`${fmt(st.w_start_min)} ~ +${fmt(st.w_start_max)} mg`, "질량 열의 시작 읽음값", `절대 영점이 없음 — 24개 중 ${st.n_negative_end}개가 음수로 끝남 → 잔류율(%) · J/g 산정 불가`],
    ["0.16 · 0.30 mg", "EPS 가연성 공기의 총 질량감소", "200 °C 이전 겉보기 증가 17 · 10 % (부력, 블랭크 없음) · r_max 반복 차이 33 % → 재측정"],
    ["2.0–5.1 %", "PF의 150 °C 이전 감소 (총 감소 대비)", "수분 · 휘발분 — PF의 T_α5(151–198 °C)는 분해 시작이 아님. 공기 시료는 승온 전 등온에서도 21–48 µg 감소"],
  ];
  callouts.forEach(([big, lab, desc], i) => {
    const y = 1.85 + i * 1.6;
    s.addShape(pres.ShapeType.roundRect, { x: 8.75, y, w: 4.08, h: 1.45, rectRadius: 0.08, fill: { color: C.background2 }, line: { color: C.background2, width: 0.5 } });
    text(s, big, { x: 8.95, y: y + 0.1, w: 3.7, h: 0.45, fontSize: 21, bold: true });
    text(s, lab, { x: 8.95, y: y + 0.55, w: 3.7, h: 0.28, fontSize: 12, bold: true, color: C.text2 });
    text(s, desc, { x: 8.95, y: y + 0.84, w: 3.75, h: 0.6, fontSize: 10.5, color: C.text2 });
  });
  text(s, `잡음 ${fmt(st.noise_min)}–${fmt(st.noise_max)} µg · 신호/잡음 ≥ ${st.snr_min} (잡음은 문제 아님)`,
    { x: 8.75, y: 6.68, w: 4.08, h: 0.24, fontSize: 10, color: C.text2 });
  s.addNotes([
    "[스크립트]",
    "- 막대는 24개 시험의 총 질량감소(승온 시작 ~ 790 °C)입니다. 녹색 점선이 '양호' 기준 1.0 mg, 빨간 점선이 '미흡' 경계 0.5 mg입니다.",
    `- ${nOk}개 시험은 ${fmt(okMin, 1)} mg 이상으로 충분합니다. EPS N₂ 3개(0.66 · 0.87 · 0.99 mg)는 '보통'이고, EPS 가연성 공기 2개(0.30 · 0.16 mg)는 '미흡'입니다.`,
    "- EPS는 밀도가 매우 낮아 도가니에 들어가는 질량 자체가 작습니다. 그 결과 부력에 의한 겉보기 증가가 총 감소의 10–17 %까지 커지고, r_max 반복 차이도 33 %가 됩니다.",
    `- 질량 열에는 절대 영점이 없습니다. 시작 읽음값이 ${fmt(st.w_start_min)} ~ +${fmt(st.w_start_max)} mg이고 24개 중 ${st.n_negative_end}개가 음수로 끝납니다. 초기 질량 기록이 없으면 잔류율(%)과 J/g은 계산할 수 없습니다.`,
    "- PF는 150 °C까지 총 감소의 2.0–5.1 %가 일어납니다(수분 · 휘발분). PF의 T_α5를 '분해 시작'으로 읽으면 안 됩니다.",
    "",
    "[질문 대비]",
    `- 잡음은 충분히 작은가: 등온 구간 잡음은 ${fmt(st.noise_min)}–${fmt(st.noise_max)} µg, 신호/잡음은 ${st.snr_min} 이상입니다. 문제는 잡음이 아니라 부력과 영점입니다.`,
    "- 1.0 mg 기준의 근거: 이 분석의 자체 기준입니다. 겉보기 증가(부력)가 수십 µg 수준이라 1 mg이면 그 영향이 몇 % 이내로 줄어듭니다.",
    "- PF 공기 시료의 등온 중 감소(21–48 µg)는 30 °C 등온에서 일어난 것으로, 수분 증발로 보입니다(승온 전이므로 α 계산에는 들어가지 않음).",
  ].join("\n"));
}

// ---------------------------------------------------------------- 11 reliability: endpoint
{
  const s = content("신뢰성 평가", "신뢰성 ② 종점 완결성",
    "N₂의 PIR(PUR) · PF와 공기의 EPS 준불연은 790 °C에서도 분해 중이라\n정규화 지표가 종점에 따라 바뀌고, 700 °C로 끊으면 중공층 순위상관이 1.0 → 0.8");
  const ch = { x: MX, y: 2.12, w: 6.8, h: 4.75, L: { x: 0.08, y: 0.05, w: 0.78, h: 0.8 }, yMin: 0, yMax: 110 };
  text(s, "790 °C의 남은 감소속도 (피크 대비 %, 두 반복 중 큰 값)", { x: MX, y: 1.8, w: 4.6, h: 0.3, fontSize: 11, bold: true, valign: "middle" });
  legend(s, [{ kind: "box", label: "N₂", color: HEX.n2, w: 0.35 }, { kind: "box", label: "공기(O)", color: HEX.air, w: 0.75 }], 5.05, 1.8);
  s.addChart(pres.ChartType.bar, [ATM_N2, ATM_AIR].map((atm) => ({ name: atm, labels: ORDER.map((m) => CAT[m]), values: ORDER.map((m) => Math.max(0, V(m, atm).end_rate_pct)) })), {
    x: ch.x, y: ch.y, w: ch.w, h: ch.h, ...axisStyle(), layout: { ...ch.L }, barDir: "col", barGrouping: "clustered", barGapWidthPct: 50,
    chartColors: [HEX.n2, HEX.air], showTitle: false, showLegend: false,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0",
    valAxisMinVal: 0, valAxisMaxVal: 110, valAxisMajorUnit: 20, valAxisLabelFormatCode: "0", objectName: "end rate",
  });
  const m = plotMap(ch);
  refLine(s, m, 5, "양호 ≤ 5 %", STATUS.good.color);
  refLine(s, m, 20, "미흡 > 20 %", STATUS.poor.color);
  badge(s, 1, 7.6, 1.85, "공기 r_max의 종점 민감도 (%/°C)", 4.8);
  const hh = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 }, align: "center" } });
  const mats = ["PF_in", "EPS_SC", "PIR_C", "PIR_SC", "PF_out"];
  const rows = [[hh("재료 (공기)"), hh("700 °C"), hh("750 °C"), hh("790 °C")]];
  mats.forEach((k) => rows.push([{ text: LABEL[k] }, ...[700, 750, 790].map((e) => {
    const flip = e === 700 && (k === "EPS_SC" || k === "PF_in");
    return { text: fmt(EP(k, ATM_AIR, e).r_max_pct_C, 3), options: { align: "center", bold: flip, fill: { color: flip ? STATUS.fair.tint : "FFFFFF" } } };
  })]));
  table(s, rows, { x: 7.6, y: 2.3, w: 5.23, colW: [2.2, 1.01, 1.01, 1.01], fontSize: 12, color: C.text1, valign: "middle", border: { type: "solid", pt: 0.75, color: "D1D5DB" }, margin: [0.04, 0.08, 0.04, 0.08], rowH: 0.34 });
  const ta90 = [700, 750, 790].map((e) => fmt(EP("EPS_SC", ATM_AIR, e).T_a90_C, 0)).join(" → ");
  text(s, paras([
    "노란 칸: 700 °C로 끊으면 EPS 준불연이 PF 내부용을 앞질러 중공층 순위상관이 0.8로 떨어짐",
    `EPS 준불연 공기 T_α90: ${ta90} °C (종점 700 → 750 → 790 °C)`,
    "T_onset · T_peak는 정규화 상수와 무관 → 종점 의존 조합은 온도 지표 위주로 사용",
    "해결: 800 °C 등온 유지(또는 1000 °C까지)로 감소를 끝까지 담기",
  ], { space: 6 }), { x: 7.6, y: 4.5, w: 5.23, h: 2.3, fontSize: 12.5 });
  s.addNotes([
    "[스크립트]",
    "- 승온이 끝나는 790 °C에서 질량감소 속도가 피크의 몇 %나 남아 있는지를 봤습니다. 0에 가까워야 '다 끝난' 시험입니다.",
    "- N₂의 PF 외부용은 102 %로, 끝에서 오히려 가장 빠르게 줄고 있습니다. PF 내부용 65 %, PIR(PUR) 준불연 41 %, 가연성 25 %입니다. 공기에서는 EPS 준불연이 40 %로 남습니다.",
    "- 이 경우 '총 질량감소'가 승온 종료 시점에 좌우되므로, 같은 원자료를 700 · 750 · 790 °C로 끊어 다시 정규화해 봤습니다.",
    "- 공기 r_max는 PF 2종 · PIR(PUR) 가연성에서 거의 안 바뀌지만, EPS 준불연은 1.24 → 1.01로 바뀌고 700 °C에서는 PF 내부용(1.18)보다 커집니다. 그래서 중공층 순위상관이 0.8로 떨어집니다.",
    "",
    "[질문 대비]",
    "- 왜 T_onset · T_peak는 괜찮은가: α와 r이 같은 상수로 나뉘므로 접선 외삽 식에서 상수가 지워지고, 피크 위치도 상수와 무관합니다.",
    "- EPS 준불연이 공기에서 늦게까지 줄어드는 이유: 난연 처리된 잔류물(탄화층)이 고온까지 서서히 산화되는 것으로 보입니다. 확인하려면 800 °C 등온 유지 측정이 필요합니다(추정).",
  ].join("\n"));
}

// ---------------------------------------------------------------- 12 reliability: replicates
{
  const s = content("신뢰성 평가", "신뢰성 ③ 반복성 · 순위 분리도",
    "공기 r_max 순위에서 PF 2종 · PIR(PUR) 가연성은 분리도 9.7–47로 확실하지만,\nEPS 준불연(1.8) · PIR(PUR) 준불연(1.5)은 반복 차이와 비슷함");
  const sep = (m, atm) => tab(D.separation, { material: m, atm }).sep_r_max;
  const ch = { x: MX, y: 2.12, w: 7.25, h: 4.75, L: { x: 0.075, y: 0.05, w: 0.79, h: 0.8 }, yMin: 0.1, yMax: 100, yLog: true };
  text(s, "r_max 순위 분리도 S (로그 축)", { x: MX, y: 1.8, w: 3.1, h: 0.3, fontSize: 11, bold: true, valign: "middle" });
  legend(s, [{ kind: "box", label: "N₂", color: HEX.n2, w: 0.35 }, { kind: "box", label: "공기(O)", color: HEX.air, w: 0.75 }], 3.65, 1.8);
  s.addChart(pres.ChartType.bar, [ATM_N2, ATM_AIR].map((atm) => ({ name: atm, labels: ORDER.map((m) => CAT[m]), values: ORDER.map((m) => sep(m, atm)) })), {
    x: ch.x, y: ch.y, w: ch.w, h: ch.h, ...axisStyle(), layout: { ...ch.L }, barDir: "col", barGrouping: "clustered", barGapWidthPct: 50,
    chartColors: [HEX.n2, HEX.air], showTitle: false, showLegend: false,
    valAxisMinVal: 0.1, valAxisMaxVal: 100, valAxisLogScaleBase: 10, valAxisCrossesAt: 0.1, valAxisLabelFormatCode: "0.#", objectName: "separation",
  });
  const m = plotMap(ch);
  refLine(s, m, 3, "양호 ≥ 3", STATUS.good.color);
  refLine(s, m, 1.5, "미흡 < 1.5", STATUS.poor.color);
  const items = [
    [["공기 — 확실  ", { bold: true }], [`PIR(PUR) 가연성 ${fmt(sep("PIR_C", ATM_AIR), 0)} · PF 외부용 ${fmt(sep("PF_out", ATM_AIR), 0)} · PF 내부용 ${fmt(sep("PF_in", ATM_AIR), 1)}`, {}]],
    [["공기 — 약함  ", { bold: true }], [`EPS 준불연 ${fmt(sep("EPS_SC", ATM_AIR), 1)} (PF 내부용과) · PIR(PUR) 준불연 ${fmt(sep("PIR_SC", ATM_AIR), 1)} (PF 외부용과) · EPS 가연성 ${fmt(sep("EPS_C", ATM_AIR), 1)} → 3번째 반복 필요`, {}]],
    [["N₂ — 확실  ", { bold: true }], [`PIR(PUR) 가연성 ${fmt(sep("PIR_C", ATM_N2), 0)} · EPS 준불연 ${fmt(sep("EPS_SC", ATM_N2), 0)} · EPS 가연성 ${fmt(sep("EPS_C", ATM_N2), 1)} · PIR(PUR) 준불연 ${fmt(sep("PIR_SC", ATM_N2), 1)}`, {}]],
    [["N₂ — PF 2종  ", { bold: true }], [`${fmt(sep("PF_out", ATM_N2), 1)} · ${fmt(sep("PF_in", ATM_N2), 1)}: N₂에서는 두 PF를 구분할 수 없음 (같은 열분해와 일치)`, {}]],
    [["온도 반복 차이  ", { bold: true }], ["대부분 ≤ 10 °C. 큰 값은 모두 T_onset: PF 외부용 N₂ 42 °C · PIR(PUR) 준불연 공기 23 °C · PF 내부용 N₂ 23 °C", {}]],
  ];
  text(s, paras(items, { space: 9 }), { x: 8.1, y: 1.85, w: 4.73, h: 5.0, fontSize: 12.5 });
  s.addNotes([
    "[스크립트]",
    "- 반복이 2회뿐이라 표준편차 대신 '두 반복의 차이'를 쓰고, 순위가 믿을 만한지는 분리도 S = (순위상 이웃 재료와의 r_max 차이) ÷ (반복 차이)로 봤습니다. S가 3 이상이면 반복을 해도 순서가 바뀌지 않는다고 봅니다.",
    "- 공기에서 PIR(PUR) 가연성 47, PF 외부용 32, PF 내부용 9.7은 확실합니다.",
    "- 반면 EPS 준불연(1.8, 이웃은 PF 내부용)과 PIR(PUR) 준불연(1.5, 이웃은 PF 외부용)은 반복 차이와 크게 다르지 않습니다. 중공층 연계에서 바로 이 두 쌍의 순서가 쓰이므로, 공기 3번째 반복이 필요합니다.",
    "- N₂에서 PF 2종의 S가 1 미만인 것은 측정 불량이 아니라 두 PF의 열분해가 실제로 같다는 결과와 일치합니다.",
    "",
    "[질문 대비]",
    "- 반복 2회로 신뢰구간을 낼 수 있나: 2회로는 의미 있는 표준편차를 내기 어렵습니다. 그래서 3번째 반복을 필수로 넣었습니다(15쪽).",
    "- 로그 축을 쓴 이유: 분리도가 0.9–47로 50배 범위라서 선형 축에서는 경계(1.5, 3) 근처 막대가 보이지 않습니다. 막대는 0.1에서 시작하고, 값은 오른쪽 글에 적었습니다.",
  ].join("\n"));
}

// ---------------------------------------------------------------- 13 reliability: records
{
  const st = D.stats;
  const s = content("신뢰성 평가", "신뢰성 ④ 온도 · 가스 · 열류 기록",
    "온도 · 승온 · 가스 기록은 정상이지만, DSC는 교정 · 기준선 보정 없이 기록돼\n기준선 이동(111–180 mW)이 가장 큰 사건보다 6배 이상 큼");
  const sh = D.runs.filter((r) => r.material === "PF_in" && r.atm === ATM_AIR).map((r) => fmt(r.self_heating_C, 1)).join(" · ");
  const shMax = Math.max(...D.runs.filter((r) => !(r.material === "PF_in" && r.atm === ATM_AIR)).map((r) => r.self_heating_C));
  const tiles = [
    ["good", `${fmt(st.rate_min)}–${fmt(st.rate_max)} °C/min`, "실제 승온속도 (설정 5)", "24개 모두 설정의 98.9 % 이상, 시험 간 차이 0.03 °C/min"],
    ["good", `${fmt(st.lag_max, 1)} ~ ${fmt(st.lag_min, 1)} °C`, "시료 − 프로그램 온도차 (중앙값)", "시료가 프로그램보다 늦음 — 모든 온도 지표는 시료 온도 기준"],
    ["good", `≤ ${fmt(shMax, 1)} °C`, "자기발열 (150 °C 이상)", `PF 내부용 공기만 ${sh} °C (474 °C, 발열 사건과 같은 온도)`],
    ["good", `${fmt(st.flow_mean_min, 1)} mL/min`, "가스 유량 평균", "단발 이탈 3회(26.1 · 10.9 · 24.0) 외 19.7–19.9. O 가스 종류(공기/산소)는 기록 없음"],
    ["poor", "계수 1 · 기준선 0", "열류 교정 · 기준선 열", "교정 · 블랭크 보정 없이 내보낸 값 → DSC는 사건의 온도 · 부호만"],
    ["poor", `${fmt(st.hf_drift_min, 0)}–${fmt(st.hf_drift_max, 0)} mW`, "열류 기준선 이동 (100 → 700 °C)", "가장 큰 사건(17.9 mW)의 6배 이상 → 엔탈피(J/g) 산정 불가"],
  ];
  const tw = (CW - 2 * 0.25) / 3, th = 2.3;
  tiles.forEach(([stt, big, lab, desc], i) => {
    const col = i % 3, row = Math.floor(i / 3), x = MX + col * (tw + 0.25), y = 1.9 + row * (th + 0.25);
    s.addShape(pres.ShapeType.roundRect, { x, y, w: tw, h: th, rectRadius: 0.08, fill: { color: C.background2 }, line: { color: C.background2, width: 0.5 }, objectName: `tile ${lab}` });
    chip(s, x + 0.25, y + 0.22, 0.95, stt);
    text(s, big, { x: x + 0.25, y: y + 0.66, w: tw - 0.5, h: 0.55, fontSize: 24, bold: true });
    text(s, lab, { x: x + 0.25, y: y + 1.22, w: tw - 0.5, h: 0.3, fontSize: 13, bold: true, color: C.text2 });
    text(s, desc, { x: x + 0.25, y: y + 1.55, w: tw - 0.45, h: 0.68, fontSize: 11.5, color: C.text2 });
  });
  s.addNotes([
    "[스크립트]",
    `- 위 줄은 온도와 가스 기록입니다. 승온속도는 ${fmt(st.rate_min)}–${fmt(st.rate_max)} °C/min로 설정 5 °C/min과 맞고, 시료 온도는 프로그램보다 ${fmt(-st.lag_max, 1)}–${fmt(-st.lag_min, 1)} °C 늦습니다. 분석의 모든 온도는 시료 온도라 이 지연은 이미 반영돼 있습니다.`,
    `- 자기발열(시료가 자기 반응열로 앞서 가는 현상)은 150 °C 이상에서 ${fmt(shMax, 1)} °C 이하이고, PF 내부용 공기만 ${sh} °C가 474 °C에서 나타납니다. DSC 발열 사건과 같은 온도라 앞뒤가 맞습니다.`,
    "- 가스 유량은 평균 19.8 mL/min으로 정상이고 단발성 이탈 3번만 있습니다. 다만 O 분위기의 가스가 공기인지 산소인지는 기록에 없습니다.",
    "- 아래 줄은 DSC입니다. 열류 교정 계수가 1, 기준선 열이 0으로 비어 있어 보정 없이 내보낸 값이고, 기준선이 100 → 700 °C에서 111–180 mW 움직입니다. 가장 큰 사건(17.9 mW)의 6배 이상이라 면적(J/g)을 낼 수 없습니다.",
    "",
    "[질문 대비]",
    "- 온도 교정은 확인했나: 원자료에는 교정 기록이 없습니다. 반복 간 피크 온도 차가 대부분 몇 °C로 작아 '상대 비교'에는 문제없지만, 절대 정확도는 교정 기록(표준물질 · 교정일)으로 확인해야 합니다.",
    "- 자기발열 값의 산정: 시료−프로그램 온도차에서 10분 이동 중앙값을 빼고, 150 °C 이상에서 최대 이탈을 봤습니다(승온 시작의 과도 응답 제외). 공통으로 235 °C 부근에 0.2 °C 정도의 작은 이탈이 있는데, 모든 시험에서 같은 위치라 장비 쪽 요인으로 봅니다.",
  ].join("\n"));
}

// ---------------------------------------------------------------- 14 verdict matrix
{
  pres.addSection({ title: "판정과 추가 실험" });
  const s = content("판정과 추가 실험", "판정 매트릭스",
    `12개 조합 판정 — 신뢰 ${counts.reliable} · 조건부 ${counts.conditional} · 종점 의존 ${counts.endpoint} · 재측정 ${counts.remeasure},\nDSC는 재현 2건 외에는 정성 참고`);
  const hh = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 }, align: "center", valign: "middle" } });
  const rows = [[hh("재료"), hh("분위기"), hh("신호량\n(mg)"), hh("겉보기 증가\n(%)"), hh("종점 잔여\n(% 피크)"), hh("온도 반복\n(°C)"), hh("속도 반복\n(%)"), hh("순위 분리\n(S)"), hh("DSC\n재현"), hh("종합")]];
  const cellv = (val, st) => ({ text: [{ text: val, options: { bold: true } }, { text: "  " + statusOf(st).label, options: { fontSize: 9.5, color: C.text2 } }], options: { align: "center", fill: { color: statusOf(st).tint } } });
  const evT = (mat, atm) => fmt(runs(mat, atm).reduce((a, r) => a + r.dsc_event_C, 0) / 2, 0);
  ORDER.forEach((mat) => [ATM_N2, ATM_AIR].forEach((atm, k) => {
    const v = V(mat, atm);
    const endTxt = v.end_rate_pct < 1 ? "<1" : fmt(v.end_rate_pct, 0);
    const dscTxt = v.v_dsc === "fair" ? `${evT(mat, atm)} °C` : "—";
    rows.push([
      k === 0 ? { text: LABEL[mat], options: { rowspan: 2, bold: true, valign: "middle" } } : null,
      { text: atm === ATM_N2 ? "N₂" : "공기", options: { align: "center" } },
      cellv(fmt(v.signal_mg, 2), v.v_signal), cellv(fmt(v.gain_pct, 1), v.v_gain), cellv(endTxt, v.v_end),
      cellv(fmt(v.temp_rep_C, 0), v.v_temp_rep), cellv(fmt(100 * v.rate_rep, 0), v.v_rate_rep), cellv(fmt(v.sep_r_max, 1), v.v_rank),
      { text: v.v_dsc === "fair" ? [{ text: "재현", options: { bold: true, breakLine: true } }, { text: dscTxt, options: { fontSize: 9.5 } }] : "비재현", options: { align: "center", color: v.v_dsc === "fair" ? C.text1 : C.text2, fill: { color: v.v_dsc === "fair" ? STATUS.fair.tint : "FFFFFF" } } },
      { text: statusOf(v.overall).label, options: { align: "center", bold: true, fill: { color: statusOf(v.overall).tint } } },
    ].filter((c) => c !== null));
  }));
  table(s, rows, { x: MX, y: 1.78, w: CW, colW: [1.75, 0.7, 1.22, 1.22, 1.22, 1.22, 1.22, 1.22, 1.08, 1.48], fontSize: 11, color: C.text1, valign: "middle",
    border: { type: "solid", pt: 0.75, color: "FFFFFF" }, margin: [0.02, 0.05, 0.02, 0.05], rowH: [0.5, ...new Array(12).fill(0.32)] });
  text(s, paras([
    [["칸 판정  ", { bold: true }], ["양호 · 보통 · 미흡 (기준: 16쪽, 이 분석의 자체 기준). DSC 재현 = 두 반복 최대 사건 온도 차 ≤ 10 °C · 같은 부호 · ≥ 0.5 mW", {}]],
    [["종합 (TGA)  ", { bold: true }], ["재측정 = 신호량 · 겉보기 증가 · 속도 반복 중 미흡 / 종점 의존 = 종점 잔여 미흡 / 조건부 = 보통 포함 / 신뢰 = 6개 모두 양호", {}]],
  ], { bullet: false, space: 3 }), { x: MX, y: 6.42, w: CW, h: 0.56, fontSize: 10.5, color: C.text2 });
  s.addNotes([
    "[스크립트]",
    "- 12개 조합 × 7개 점검의 판정입니다. 칸 안의 숫자가 점검값, 옆 글자가 판정입니다.",
    "- 맨 오른쪽 종합은 TGA 6개 점검으로만 정했습니다. 순서대로 '재측정'(측정 자체가 약함) → '종점 의존'(정규화 지표가 종점에 좌우) → '조건부'(보통 항목 있음) → '신뢰'입니다.",
    "- 신뢰: 공기 PF 내부용 · PF 외부용 · PIR(PUR) 가연성. 조건부: N₂ EPS 2종, 공기 PIR(PUR) 준불연. 종점 의존: N₂ PIR(PUR) 2종 · PF 2종, 공기 EPS 준불연. 재측정: 공기 EPS 가연성.",
    "- DSC는 12개 중 2개만 재현됩니다(EPS 가연성 N₂ 398 °C 흡열, PF 내부용 공기 474 °C 발열).",
    "",
    "[질문 대비]",
    "- 순위 분리가 '미흡'인 N₂ PF 2종: 두 PF가 N₂에서 같은 재료처럼 분해한다는 결과와 일치하는 것이지 측정 불량이 아닙니다. 종합 판정은 종점 잔여 때문에 '종점 의존'입니다.",
    "- 종점 잔여가 '<1'인 칸: 790 °C에서 감소속도가 피크의 1 % 미만(일부는 측정 잡음으로 약간 음수)이라는 뜻입니다.",
    "- 판정 기준의 민감도: 모든 경계를 20 % 느슨하게 하면 판정이 하나도 바뀌지 않고, 20 % 엄격하게 하면 PIR(PUR) 준불연 공기(r_max 반복 차이 19.5 %)가 '재측정'으로 바뀝니다. '신뢰 3'은 어느 쪽에서도 유지됩니다. PIR(PUR) 준불연 공기는 3번째 반복(필수)에 들어 있습니다.",
  ].join("\n"));
}

// ---------------------------------------------------------------- 15 additional experiments
{
  const s = content("판정과 추가 실험", "추가 실험 판단",
    "추가 실험은 필요하나 전면 재시험은 아님 —\n필수 10회(블랭크 2 · EPS 가연성 공기 3 · 공기 3번째 반복 5)와 기록 확인 4건");
  const hh = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 }, valign: "middle" } });
  const grp = (t, sub, st, n) => ({ text: [{ text: t, options: { bold: true, breakLine: true } }, { text: sub, options: { fontSize: 10, color: C.text2 } }], options: { rowspan: n, valign: "middle", fill: { color: STATUS[st].tint } } });
  const r = (a, b, c, d) => [{ text: a, options: { bold: true } }, { text: b }, { text: c, options: { align: "center", bold: true } }, { text: d, options: { color: C.text2 } }];
  const rows = [
    [hh("구분"), hh("항목"), hh("조건"), hh("횟수"), hh("해결되는 문제")],
    [grp("필수", "10회", "remeasure", 3), ...r("블랭크(빈 도가니)", "N₂ · 공기 각 1회, 본 시험과 같은 프로그램 · 도가니", "2", "부력 · 열류 기준선 보정 → 겉보기 증가 제거, DSC 기준선")],
    r("EPS 가연성 공기 재측정", "시료량을 늘려 총 질량감소 ≥ 1 mg (현재 0.16 · 0.30), 초기 질량 기록", "3", "재측정 판정 해소 (신호 부족, 겉보기 증가 17 %, r_max 반복 차이 33 %)"),
    r("공기 3번째 반복", "나머지 5종, 같은 5 °C/min + 800 °C 등온 30분, 초기 질량 기록", "5", "순위 분리도 확인 (EPS 준불연 1.8 · PIR(PUR) 준불연 1.5) → 중공층 순위상관 확정, 종점 완결"),
    [grp("기록 확인", "시험 없음", "fair", 4), ...r("시료 초기 질량 24개", "측정 기록지 · 장비 로그", "—", "잔류율(%) · 초기 질량 기준 T5% · J/g 계산 가능")],
    r("O 가스 종류", "공기(O₂ 21 %)인지 순수 산소인지", "—", "산화 조건 해석, 화재 조건과의 대응"),
    r("교정 기록", "온도 · 질량 · 열류 교정일과 표준물질", "—", "온도 지표의 절대 정확도"),
    r("시료 채취", "심재 / 면재 포함 여부, 시료 형태 · 도가니", "—", "재료 대표성, 반복 간 차이의 원인"),
    [grp("권장", "목적별", "reliable", 4), ...r("N₂ 완결 측정", "PIR(PUR) 2종 · PF 2종, 800 °C 등온 30분 또는 1000 °C까지", "4–8", "N₂ 종점 의존 해소 (현재 790 °C에서 피크의 23–102 %)")],
    r("승온속도 변화", "공기, 10 · 20 °C/min (기존 5 °C/min 포함 3속도), 6종", "12–24", "활성화에너지(등전환율법) → 화재 가열속도로 외삽"),
    r("비열 측정", "사파이어법 DSC (ASTM E1269), 6종", "6–12", "c_p → 열관성(kρc) 역산 결과 검증"),
    r("발열량 측정", "MCC (ASTM D7309) 또는 콘칼로리미터 (ISO 5660-1)", "—", "J/g · 열방출률 — DSC로는 불가"),
  ];
  table(s, rows, { x: MX, y: 1.8, w: CW, colW: [1.05, 2.25, 4.15, 0.75, 4.13], fontSize: 11, color: C.text1, valign: "middle",
    border: { type: "solid", pt: 0.75, color: "D1D5DB" }, margin: [0.03, 0.08, 0.03, 0.08], rowH: [0.34, ...new Array(11).fill(0.42)] });
  s.addNotes([
    "[스크립트]",
    "- 질문하신 '추가 실험이 필요한가'에 대한 답은 '필요하지만 표적 보완'입니다. 필수 10회와 기록 확인 4건이면 재측정 · 조건부 판정이 풀리고, 중공층 연계(순위상관)를 확정할 수 있습니다.",
    "- 필수 ① 블랭크 2회: 빈 도가니로 같은 프로그램을 돌려 부력과 열류 기준선을 빼면 겉보기 증가가 사라지고 DSC 기준선도 생깁니다.",
    "- 필수 ② EPS 가연성 공기 3회: 시료량을 늘려 질량감소를 1 mg 이상으로 맞춥니다. 3회로 반복 범위를 냅니다.",
    "- 필수 ③ 공기 3번째 반복 5회: 중공층 연계에 쓰는 공기 r_max의 순위(특히 EPS 준불연 vs PF 내부용, PIR(PUR) 준불연 vs PF 외부용)를 확정합니다. 800 °C 등온 30분을 붙여 종점 문제도 같이 없앱니다.",
    "- 기록 확인 4건은 시험 없이 됩니다. 특히 초기 질량 24개만 있으면 잔류율(%)과 T5%를 바로 계산할 수 있습니다.",
    "- 권장 항목은 목적이 있을 때입니다: N₂ 지표 확정, 반응속도론(활성화에너지), 열물성(c_p), 발열량.",
    "",
    "[질문 대비]",
    "- 필수 10회의 장비 시간: 1회당 등온 5분 + 승온 약 155분(+ 등온 30분)으로 약 3시간, 10회면 약 30시간 수준입니다(냉각 · 준비 제외).",
    "- 왜 공기만 3번째 반복인가: 중공층 연계와 PF 비교의 결론이 공기 자료에서 나오고, N₂는 종점 문제가 먼저라 '완결 측정'(권장)으로 다룹니다.",
    "- 승온속도 3속도: 등전환율 반응속도 해석에는 3개 이상의 승온속도가 필요합니다(ICTAC 권고).",
    "- 블랭크는 장비 · 도가니 · 가스가 같아야 의미가 있으므로 본 시험과 같은 조건으로 합니다.",
  ].join("\n"));
}

// ---------------------------------------------------------------- 16 appendix
{
  pres.addSection({ title: "부록" });
  const s = content("부록", "부록 — 검산 · 판정 기준 · 재현",
    "모든 수치는 원자료 24개에서 스크립트 3개로 다시 만들어지며,\n요약표 · 미분 방법 · 자기발열 · 순위상관을 독립 계산으로 확인함");
  badge(s, 1, MX, 1.85, "판정 기준 (이 분석의 자체 기준, 표준 아님)", 6);
  const hh = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 }, align: "center" } });
  const crit = [
    ["신호량 (작은 반복의 총 감소)", "≥ 1.0 mg", "0.5–1.0", "< 0.5"],
    ["겉보기 증가 (200 °C 이전, 총 감소 대비)", "≤ 2 %", "2–5 %", "> 5 %"],
    ["종점 잔여 (790 °C 감소속도 ÷ 피크)", "≤ 5 %", "5–20 %", "> 20 %"],
    ["온도 반복 (T_onset · T_peak · T_α50 중 최대)", "≤ 10 °C", "10–25 °C", "> 25 °C"],
    ["속도 반복 (r_max 반복 차이 ÷ 평균)", "≤ 10 %", "10–20 %", "> 20 %"],
    ["순위 분리도 S", "≥ 3", "1.5–3", "< 1.5"],
  ];
  const tint = (st) => ({ align: "center", fill: { color: STATUS[st].tint } });
  table(s, [[hh("점검"), hh("양호"), hh("보통"), hh("미흡")], ...crit.map(([a, b, c, d]) => [{ text: a }, { text: b, options: tint("good") }, { text: c, options: tint("fair") }, { text: d, options: tint("poor") }])],
    { x: MX, y: 2.27, w: 6.3, colW: [3.3, 1.0, 1.0, 1.0], fontSize: 11, color: C.text1, valign: "middle", border: { type: "solid", pt: 0.75, color: "D1D5DB" }, margin: [0.03, 0.07, 0.03, 0.07], rowH: 0.36 });
  badge(s, 2, MX, 4.95, "재현 (tga_dsc 폴더에서)", 6);
  text(s, [
    { text: "python analysis.py          # 요약표 · 중공층 연계", options: { breakLine: true } },
    { text: "python qa_reliability.py    # 신뢰성 점검 results/qa_*.csv", options: { breakLine: true } },
    { text: "python report/deck_data.py  # deck_data.json", options: { breakLine: true } },
    { text: "node report/build_deck.js   # 이 PPT (pptxgenjs)" },
  ], { x: MX, y: 5.38, w: 6.3, h: 1.3, fontSize: 10.5, fontFace: "Consolas", color: C.text1, fill: { color: C.background2 }, margin: [0.08, 0.12, 0.08, 0.12] });
  badge(s, 3, 7.15, 1.85, "검산", 5);
  const xc = (k) => D.crosscheck.find((c) => c.check.startsWith(k));
  const rows = [
    ["요약표 평균 = 개별 시험 평균", `차이 ${fmt(xc("summary mean r_max").min_pct, 2)} ~ +${fmt(xc("summary mean r_max").max_pct, 2)} % (반올림)`],
    ["r_max: 평활 미분 vs 10 °C 중앙차분", `${fmt(xc("r_max:").min_pct, 1)} ~ ${fmt(xc("r_max:").max_pct, 1)} % — 순서 동일`],
    ["T_peak: 두 방법", `≤ 2.3 °C (PF 내부용 N₂만 ${fmt(xc("T_peak:").max_abs_diff, 0)} °C, 넓고 평탄한 피크)`],
    ["자기발열 (150 °C 이상)", "≤ 0.2 °C, PF 내부용 공기 0.4 · 1.0 °C"],
    ["종점별 순위상관 (준불연 4종)", "700 · 750 · 790 °C → 0.8 · 1.0 · 1.0"],
    ["DSC 부호", "EPS 해중합(흡열)이 N₂에서 + → 흡열 = +"],
  ];
  table(s, [[hh("항목"), hh("결과")], ...rows.map(([a, b]) => [{ text: a, options: { bold: true } }, { text: b }])],
    { x: 7.15, y: 2.27, w: 5.68, colW: [2.55, 3.13], fontSize: 11, color: C.text1, valign: "middle", border: { type: "solid", pt: 0.75, color: "D1D5DB" }, margin: [0.04, 0.07, 0.04, 0.07], rowH: 0.5 });
  text(s, "결과 파일: tga_dsc/results/qa_*.csv · 원자료(data/N, data/O)는 저장소에 포함하지 않음", { x: 7.15, y: 5.95, w: 5.68, h: 0.5, fontSize: 10.5, color: C.text2 });
  s.addNotes([
    "[스크립트]",
    "- ① 판정 기준입니다. 표준에서 가져온 값이 아니라 이 자료의 크기와 반복 차이를 보고 정한 자체 기준입니다. 숫자는 qa_reliability.py 맨 위에 상수로 있습니다.",
    "- ② 재현 순서입니다. 원자료(data/N, data/O)를 넣고 네 줄을 실행하면 이 PPT의 모든 숫자와 그래프가 다시 만들어집니다.",
    "- ③ 검산: 요약표 평균과 개별 시험 평균이 같고, 평활 미분 대신 단순 중앙차분으로 r_max를 다시 구해도 0.7–3.8 % 낮을 뿐 순서는 같습니다. T_peak는 2.3 °C 이내로 맞고, PF 내부용 N₂만 피크가 넓고 평탄해 12 °C 차이가 납니다.",
    "",
    "[질문 대비]",
    "- 자체 기준의 민감도: 모든 경계를 20 % 느슨하게 하면 판정 변화가 없고, 20 % 엄격하게 하면 PIR(PUR) 준불연 공기가 조건부 → 재측정(r_max 반복 차이 19.5 % > 16 %)으로 바뀝니다. 신뢰 3은 그대로입니다.",
    "- 왜 표준 기준이 없나: TGA 반복성 기준은 시험법마다 다르고(재료 · 목적별), 이 자료는 초기 질량이 없어 표준 방식(초기 질량 대비 %)으로 계산할 수도 없습니다.",
  ].join("\n"));
}

pres.writeFile({ fileName: OUT }).then(async (f) => {
  await finish(f);
  console.log("wrote", f);
});
