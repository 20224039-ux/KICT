#!/usr/bin/env node
// 외벽 정량적 화재위험성평가 보고서 서술 양식 (docx-js): NFPA 551(2022) 7장의 문서 요건을 항목마다
// "쓸 내용 · 작성 방법 · 예문"으로 풀어 쓴 양식. 예문의 수치는 qfra_data.json(../qfra.py)과 verification.json에서 가져옴.
//
//   node build_qfra_docx.js [out.docx]          (docx는 전역 npm 패키지)
"use strict";
const fs = require("fs");
const path = require("path");
const {
  AlignmentType, BorderStyle, Document, Footer, Header, HeadingLevel, ImageRun, LevelFormat, Packer, PageBreak, PageNumber,
  Paragraph, ShadingType, Table, TableCell, TableRow, TextRun, VerticalAlign, WidthType,
} = require("docx");

const D = JSON.parse(fs.readFileSync(path.join(__dirname, "qfra_data.json"), "utf8"));
const VER = JSON.parse(fs.readFileSync(path.join(__dirname, "verification.json"), "utf8"));
const FIG = path.join(__dirname, "..", "figures");
const OUT = process.argv[2] || path.join(__dirname, "외벽_QFRA_서술양식.docx");

// ------------------------------------------------------------------------------------------------ data
const MATS = D.materials, ORDER = ["NC", ...MATS], LABEL = D.label;
const PRM = Object.fromEntries(D.params.map((p) => [p.key, p]));
const best = (k) => PRM[k].best;
const PT = (m, b, g) => D.point.find((r) => r.material === m && r.building === b && r.mitigation === g);
const MC = (m, b, g) => D.mc.find((r) => r.material === m && r.building === b && r.mitigation === g);
const CO = (m, ie, bar) => D.conditional.find((r) => r.material === m && r.ie === ie && r.barrier === bar);
const SE = (g, m, k) => D.sensitivity.find((r) => r.mitigation === g && r.material === m && r.parameter === k);
const units = best("N") * best("units");

// ------------------------------------------------------------------------------------------------ text helpers
const FONT = "Malgun Gothic";
const INK = "23303D", INK2 = "4B5563", GREEN = "0A7A0A", EXFILL = "EEF5EC", NOTEFILL = "FEF2D3", HEADFILL = "23303D";
const SUPD = ["⁰", "¹", "²", "³", "⁴", "⁵", "⁶", "⁷", "⁸", "⁹"];
const SUPMAP = { "ʲ": "j", "⁻": "−", "⁰": "0", "¹": "1", "²": "2", "³": "3", "⁴": "4", "⁵": "5", "⁶": "6", "⁷": "7", "⁸": "8", "⁹": "9" };
// 1.26×10⁻⁴ (Unicode superscripts; rich() turns them into real superscript runs)
function sci(v, d = 2) {
  if (v === 0) return "0";
  let e = Math.floor(Math.log10(Math.abs(v)));
  let m = v / 10 ** e;
  if (Math.abs(Number(m.toFixed(d - 1))) >= 10) { m /= 10; e += 1; }
  return `${m.toFixed(d - 1)}×10${e < 0 ? "⁻" : ""}${String(Math.abs(e)).split("").map((c) => SUPD[+c]).join("")}`;
}
const f = (v, d = 2) => Number(v).toFixed(d);
const pct = (v, d = 0) => `${(100 * v).toFixed(d)} %`;
const rho = (v) => `${v < 0 ? "−" : "+"}${Math.abs(v).toFixed(2)}`;
// subscript tokens written as base_sub (or with Unicode subscript digits)
const SUBS = {
  R_S: ["R", "S"], R_P: ["R", "P"], R_t: ["R", "t"], RR_S: ["RR", "S"], F_i: ["F", "i"], C_i: ["C", "i"], C_ij: ["C", "ij"], C_L: ["C", "L"],
  "F₁": ["F", "1"], "F₂": ["F", "2"], "C₁": ["C", "1"], "C₂": ["C", "2"], "λ₁": ["λ", "1"], "λ₂": ["λ", "2"], p_s: ["p", "s"],
  p_FO: ["p", "FO"], p_v: ["p", "v"], p_x: ["p", "x"], p_e0: ["p", "e0"], p_e1: ["p", "e1"], p_e: ["p", "e"], p_cb: ["p", "cb"],
  p_L: ["p", "L"], p_B: ["p", "B"], p_d: ["p", "d"], T_FS: ["T", "FS"], T_stop: ["T", "stop"], "τ_F": ["τ", "F"], "τ_L": ["τ", "L"],
  n_F: ["n", "F"], n_L: ["n", "L"], w_j: ["w", "j"], a_j: ["a", "j"], b_j: ["b", "j"], t_e: ["t", "e"], t_ref: ["t", "ref"],
  n_a: ["n", "a"], n_u: ["n", "u"], t_ig: ["t", "ig"], t_sp: ["t", "sp"], r_max: ["r", "max"], "t₆₀₀": ["t", "600"], t_b: ["t", "b"],
  "α₆₀₀": ["α", "600"], "f₄₀₀": ["f", "400"],
};
const TOK = Object.keys(SUBS).sort((a, b) => b.length - a.length).map((t) => t.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
const RICH_RE = new RegExp(`(?<![A-Za-z_])(${TOK.join("|")})|([ʲ⁻⁰¹²³⁴⁵⁶⁷⁸⁹]+)`, "g");
// text -> TextRuns with real subscripts / superscripts
function runs(str, o = {}) {
  const base = { font: FONT, ...o };
  const out = [];
  let last = 0, m;
  RICH_RE.lastIndex = 0;
  while ((m = RICH_RE.exec(str)) !== null) {
    if (m.index > last) out.push(new TextRun({ text: str.slice(last, m.index), ...base }));
    if (m[1]) {
      const [b, s] = SUBS[m[1]];
      out.push(new TextRun({ text: b, ...base }), new TextRun({ text: s, ...base, subScript: true }));
    } else {
      out.push(new TextRun({ text: [...m[2]].map((c) => SUPMAP[c]).join(""), ...base, superScript: true }));
    }
    last = m.index + m[0].length;
  }
  if (last < str.length) out.push(new TextRun({ text: str.slice(last), ...base }));
  return out;
}
// parts: string, or [text, runOptions] pairs
const R = (parts, o = {}) => (Array.isArray(parts) ? parts : [parts]).flatMap((p) => (Array.isArray(p) ? runs(p[0], { ...o, ...p[1] }) : runs(p, o)));
const P = (parts, o = {}) => new Paragraph({ children: R(parts, o.run || {}), spacing: { after: 100, line: 300 }, ...o.para });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: runs(t), pageBreakBefore: true });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: runs(t) });
const H3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: runs(t) });
const bullets = (items, ref = "bullets") => items.map((it) => new Paragraph({ numbering: { reference: ref, level: 0 }, children: R(it), spacing: { after: 60, line: 290 } }));
const label = (t, color = GREEN) => new Paragraph({ children: [new TextRun({ text: t, bold: true, color, font: FONT, size: 20 })], spacing: { before: 140, after: 60 } });
// example box: shaded paragraphs with a green left rule
function example(paras, title = "예문") {
  const border = { left: { style: BorderStyle.SINGLE, size: 18, color: "0CA30C", space: 8 } };
  const shade = { type: ShadingType.CLEAR, fill: EXFILL, color: "auto" };
  return [
    new Paragraph({ children: [new TextRun({ text: title, bold: true, color: GREEN, font: FONT, size: 18 })], shading: shade, border, spacing: { before: 120, after: 40 }, indent: { left: 160, right: 160 } }),
    ...paras.map((p, i) => new Paragraph({ children: R(p), shading: shade, border, spacing: { after: i === paras.length - 1 ? 160 : 60, line: 300 }, indent: { left: 160, right: 160 } })),
  ];
}
function note(parts) {
  return new Paragraph({ children: R(parts, { size: 18 }), shading: { type: ShadingType.CLEAR, fill: NOTEFILL, color: "auto" },
    border: { left: { style: BorderStyle.SINGLE, size: 18, color: "FAB219", space: 8 } }, spacing: { before: 80, after: 140, line: 280 }, indent: { left: 160, right: 160 } });
}
const W = 9638;   // A4 with 2 cm margins, DXA
const BORDER = { style: BorderStyle.SINGLE, size: 4, color: "D1D5DB" };
const BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };
function table(header, rows, widths, o = {}) {
  const sum = widths.reduce((a, b) => a + b, 0);
  const ws = widths.map((w) => Math.round((w * W) / sum));
  ws[ws.length - 1] += W - ws.reduce((a, b) => a + b, 0);
  const cell = (c, i, head, fill) => new TableCell({
    width: { size: ws[i], type: WidthType.DXA }, borders: BORDERS, verticalAlign: VerticalAlign.CENTER,
    shading: head ? { type: ShadingType.CLEAR, fill: HEADFILL, color: "auto" } : fill ? { type: ShadingType.CLEAR, fill, color: "auto" } : undefined,
    margins: { top: 50, bottom: 50, left: 90, right: 90 },
    children: [new Paragraph({ children: R(c, head ? { bold: true, color: "FFFFFF", size: 18 } : { size: 18 }), alignment: head || (o.center || []).includes(i) ? AlignmentType.CENTER : AlignmentType.LEFT, spacing: { line: 260 } })],
  });
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: ws,
    rows: [new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, i, true)) }),
      ...rows.map((r, k) => new TableRow({ children: r.map((c, i) => cell(c, i, false, (o.fill && o.fill(k, i)) || null)) }))],
  });
}
const caption = (t) => new Paragraph({ children: runs(t, { size: 18, color: INK2, bold: true }), spacing: { before: 60, after: 160 } });
function image(file, widthPx) {
  const buf = fs.readFileSync(file);
  const w = buf.readUInt32BE(16), h = buf.readUInt32BE(20);          // PNG IHDR
  return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 40 },
    children: [new ImageRun({ type: "png", data: buf, transformation: { width: widthPx, height: Math.round((widthPx * h) / w) } })] });
}
// a section: what NFPA 551 asks, how to write it for a façade, example text, optional extras, checks
function section(title, clause, { what, how, ex, extra = [], check = [] }) {
  const out = [H2(title), label(`쓸 내용 — NFPA 551 ${clause}`), ...bullets(what), label("작성 방법 — 외벽(중공층 외단열)에 적용"), ...bullets(how)];
  if (ex) out.push(...example(ex));
  out.push(...extra);
  if (check.length) out.push(label("점검", INK2), ...bullets(check, "checks"));
  return out;
}

// ------------------------------------------------------------------------------------------------ example numbers
const pf = PT("PF_in", "B1", "M0"), pfmc = MC("PF_in", "B1", "M0");
const c1 = CO("PF_in", "IE1", false), cl = CO("NC", "IE1", false), c2 = CO("PF_in", "IE2", false);
const F1 = pf.F_IE1_ext, F2 = pf.F_IE2_ext;
const verdictOf = (m, g) => MC(m, "B1", g).verdict;
const confOf = (m, g) => MC(m, "B1", g).conf_meets_F2;
const scMats = MATS.filter((m) => D.grade[m] === "sc");
const minConfM1 = Math.min(...scMats.map((m) => confOf(m, "M1")));

// ================================================================================================ content
const body = [];

// ---------------------------------------------------------------------------------- cover
body.push(
  new Paragraph({ spacing: { before: 1800, after: 120 }, children: [new TextRun({ text: "외벽 화재위험성평가 · NFPA 551 정량적 방법", font: FONT, size: 22, bold: true, color: INK2 })] }),
  new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "외벽(중공층 외단열) 정량적 화재위험성평가", font: FONT, size: 44, bold: true, color: INK })] }),
  new Paragraph({ spacing: { after: 360 }, children: [new TextRun({ text: "보고서 서술 양식 — 항목별 쓸 내용 · 작성 방법 · 예문", font: FONT, size: 32, bold: true, color: INK })] }),
  P("NFPA 551(2022) 7장(문서)의 요건을 외벽 정량 평가에 맞춰 항목마다 풀어 쓴 양식입니다. 예문의 수치는 중공층 시험 5종과 15층 공동주택 예제로 계산한 값입니다(예시 입력)."),
  P([["작성일: ", { bold: true }], "2026-10-23    ", ["함께 쓰는 파일: ", { bold: true }], "외벽_QFRA_계산서.xlsx (계산), 외벽_정량적_화재위험성평가_절차.pptx (절차 설명)"]),
  new Paragraph({ spacing: { before: 360 } }),
  ...example([
    [["이 양식의 구성  ", { bold: true }], "각 절은 ① 쓸 내용(NFPA 551 요건과 조항) → ② 작성 방법(외벽 적용) → ③ 예문(예제 값으로 쓴 문장) → ④ 점검 순서입니다."],
    [["예문의 수치  ", { bold: true }], "빈도·분기확률·시간은 절차를 보이기 위한 예시값이며, 측정값은 중공층 시험의 t₆₀₀(3000 mm)과 플래그 B뿐입니다. 실제 평가에서는 국내 통계·실대형 시험값으로 바꾼 뒤 다시 계산하고 예문의 숫자를 고쳐 쓰십시오."],
    [["[ ] 표시  ", { bold: true }], "평가자가 채울 자리입니다."],
  ], "읽는 법"),
);

// ---------------------------------------------------------------------------------- 0 overview
body.push(H1("0. 개요 — NFPA 551과 이 양식"));
body.push(
  P("NFPA 551은 관할기관(AHJ)이 화재위험성평가(FRA)를 검토·승인할 때 쓰는 지침입니다(1.1). 평가 방법을 정하지 않고(1.2) 검토 절차와 문서 요건을 제시하므로, 외벽에 맞는 시나리오·모델·입력값과 서술 방법은 평가자가 구성해야 합니다. 이 양식은 그 문서 요건을 다음 세 묶음으로 정리합니다."),
  ...bullets([
    [["1장 개념보고서(7.2) ", { bold: true }], "— 분석 전에 관할기관과 합의할 8항목: 참여자, 범위·목적, 목표, 허용기준, 위험요소, 시나리오, 방법, 자료 출처"],
    [["2장 정량 결과 문서(7.3.3.4) ", { bold: true }], "— 분석 후 제출할 8항목: 빈도, 결과, 위험도, 불확실성, 소프트웨어·모델 평가, 한계, 결론, 참고문헌"],
    [["3장 운영·유지관리 매뉴얼(7.4)과 이행(7.5) ", { bold: true }], "— 평가가 유효하려면 지켜야 할 조건, 변경 관리, 점검·시험·유지관리"],
    [["4장 검토 대응(8장) ", { bold: true }], "— 검토 방식, 검증·확인, 검토 질문 9개의 답 위치"],
  ]),
  label("표 0-1  12단계 절차와 이 양식의 절", INK),
  table(["단계", "내용", "NFPA 551", "산출물", "이 양식"], [
    ["1", "사전협의 · 문제 정의", "4.3, 4.4.1", "목적·대상·범위", "1.1–1.2"],
    ["2", "위험요소 정의", "4.4.2", "노출 대상·자극·전달·반응", "1.5"],
    ["3", "허용기준 설정", "4.4.3, A.5.1.2", "지표·기준값·신뢰수준", "1.4"],
    ["4", "방법 선정", "4.4.4, 5.1.3", "방법·적합성 근거", "1.7"],
    ["5", "시나리오 선정·군집", "5.1.1.5, 5.5.1.1", "초기사건·군집·제외 근거", "1.6"],
    ["6", "빈도 분석", "5.3", "사건수목, 경로별 F", "2.1"],
    ["7", "결과 분석", "5.4, 5.4.9", "층간 확산 C", "2.2"],
    ["8", "위험도 산정", "5.5.1", "R = Σ F·C, 초과빈도", "2.3"],
    ["9", "불확실성 분석", "4.5, 5.5.6", "분포·신뢰수준·민감도", "2.4"],
    ["10", "검증·확인", "5.5.9, 8.2", "재현 계산·시험 비교", "2.5, 4.2"],
    ["11", "문서화", "7.2, 7.3.3.4, 7.4", "개념보고서·결과문서·매뉴얼", "1–3장"],
    ["12", "검토·승인", "8.1, 8.3", "검토 질문 대응", "4장"],
  ], [6, 22, 16, 30, 12], { center: [0, 2, 4] }),
  new Paragraph({ spacing: { after: 120 } }),
  P("평가 방법의 범주는 NFPA 551 표 5.1.2.1의 ‘정량적 방법’(가능성과 결과를 모두 수치로 결합)입니다. 기존의 R = P × S(확산 등급 × 심각도 등급)는 두 인자가 모두 결과 쪽 지표이므로 위험 지수(정성 범주)에 해당하며, 정량 평가에서는 여기에 사건 빈도 F를 사건수목으로 더합니다(위험 = 확률과 결과의 쌍, 3.3.14)."),
);

// ---------------------------------------------------------------------------------- 1 concept report
body.push(H1("1. 개념보고서 (NFPA 551 7.2) — 분석 전에 합의할 내용"));
body.push(P("개념보고서의 목적은 제안하는 평가 접근 방법에 대해 관할기관과 합의하는 것입니다(7.2.1). 분석이 끝난 뒤의 종합 문서에는 이 내용을 변경 사항까지 반영해 다시 넣습니다(7.3.2)."));

body.push(...section("1.1 참여자와 자격", "7.2.2.1, 1.4, 4.2", {
  what: ["평가 준비에 참여한 이해관계자와 그 자격(학력, 화재위험성평가 경력, 전문 자격)을 적습니다.",
    "평가자의 자격은 관할기관이 확인할 수 있도록 문서로 남깁니다(1.4)."],
  how: ["평가자 · 자료 제공자(시험기관, 소방서, 관리주체) · 검토자를 나눠 표로 적습니다.",
    "외벽 정량 평가는 화재공학(확산 모델), 시험(중공층·실대형), 통계(빈도)의 전문성이 함께 필요하므로 팀으로 구성하고 분담을 적습니다(A.5.1.3.6).",
    "동료 검토자는 평가에 관여하지 않은 같은 수준의 전문가로 정합니다(8.1.2.1)."],
  ex: ["본 평가는 [기관명] [팀명]이 수행하였다. 평가 책임자 [성명]은 [학위], 화재위험성평가 경력 [n]년, [자격]을 보유하며 중공층 화재시험 자료의 해석과 확률 모델 계산을 맡았다. 화재 통계는 [기관], 소방 출동 시간 자료는 [소방서]에서 제공받았으며, 동료 검토는 평가에 참여하지 않은 [성명, 소속]이 수행한다."],
  extra: [label("표 1-1  참여자 (작성 틀)", INK), table(["역할", "성명", "소속", "학력 · 경력 · 자격", "담당"], [
    ["평가 책임", "[ ]", "[ ]", "[ ]", "방법·모델·결론"], ["시험 자료", "[ ]", "[ ]", "[ ]", "중공층·실대형 시험"], ["통계 자료", "[ ]", "[ ]", "[ ]", "화재 빈도"], ["동료 검토", "[ ]", "[ ]", "[ ]", "독립 검토(8.1.2.1)"],
  ], [14, 12, 16, 34, 24])],
  check: ["역할마다 자격 근거가 있는가", "동료 검토자가 평가에 관여하지 않았는가"],
}));

body.push(...section("1.2 범위와 목적", "7.2.2.2, A.7.2.2.2, 4.4.1", {
  what: ["분석의 한계(건물, 건물의 일부, 구성요소)와 평가를 하는 목적을 적습니다(A.7.2.2.2).",
    "목적은 기존 건물의 위험 수준 파악, 위험을 낮추는 방법 도출, 신축·개축에서 수용 가능한 위험의 대안 제시 가운데 무엇인지 밝힙니다(4.4.1.1). 범위 밖의 부수 문제가 평가 목적을 흐리지 않게 합니다(4.4.1.2)."],
  how: ["목적을 한 문장으로 씁니다: 대안 비교 / 기존 건물 우선순위 / 동등성 입증.",
    "대상: 건물(층수, 높이, 용도, 세대 수, 재실자), 외벽 시스템(마감재, 중공층 깊이, 단열재 종류·등급, 개구부 차단재, 층간 캐비티 배리어).",
    "포함·제외: 사용 단계에서 외벽이 화염에 노출된 뒤 외벽을 통해 위층으로 번지는 경로만 다루고, 실내 연기 확산·공사 중 화재·인접 건물은 제외합니다(제외 근거는 1.6).",
    "위험의 단위(동·년)를 정해 둡니다."],
  ex: [`본 평가의 목적은 ${best("N")}층 공동주택(${units}세대, 재실 약 ${D.occupants}명)의 외벽에 중공층 100 mm 외단열 시스템을 적용할 때, 단열재 5종(EPS 준불연, PIR(PUR) 가연성·준불연, PF 준불연 외부용·내부용)과 개구부 차단재·층간 캐비티 배리어 적용 여부에 따른 외벽 화재 확산 위험을 정량적으로 비교하는 것이다.`,
    "평가 범위는 건물 사용 단계에서 외벽이 화염에 노출된 뒤 외벽을 통해 발화층 위로 번지는 경로로 한정한다. 실내 연기 확산, 공사 중 화재, 인접 건물로의 확산은 범위에서 제외한다. 위험은 건물 1동, 1년 기준으로 나타낸다."],
  check: ["목적이 하나로 분명한가(대안 비교 · 우선순위 · 동등성)", "외벽 시스템의 구성이 시험 시편과 같은지 적었는가"],
}));

body.push(...section("1.3 목표", "7.2.2.3, A.7.2.2.3, 4.4.2.1", {
  what: ["화재안전 목표를 분명히 적습니다: 생명(재실자·소방관), 재산, 운영, 환경 가운데 무엇을 지키는가.",
    "목표는 일반인도 이해할 수 있는 정성적 표현으로 쓰고(A.7.2.2.3), 1.4의 지표로 측정합니다."],
  how: ["외벽 평가의 1차 목표는 ‘발화층 위 재실자가 외벽을 통해 번진 화재에 피난 전 노출되지 않게 하는 것’입니다.",
    "재산 목표(여러 층으로 번지는 것을 제한)와 소방대 안전(외벽 화재의 고층 진압 한계)을 덧붙입니다.",
    "목표마다 측정 지표를 한 줄로 연결합니다(목표 → 지표 → 기준)."],
  ex: ["본 평가의 화재안전 목표는 (1) 외벽을 통한 화재 확산으로 발화층 위 거주자가 피난 전에 화염에 노출되는 일을 최소화하고, (2) 외벽 화재가 여러 층으로 번져 건물 피해가 커지는 것을 제한하는 것이다. 목표 (1)은 노출 인원 R_P와 개인위험도로, 목표 (2)는 발화층 위 확산 층 수의 초과빈도 F(n)로 측정한다."],
}));

body.push(...section("1.4 허용기준", "7.2.2.4, 4.4.3, A.5.1.2", {
  what: ["평가 결과를 판단할 지표와 기준값을 적습니다. 기준은 상대(기준안과 비교) 또는 절대(예: 연간 사망자)일 수 있고, 정성 또는 정량일 수 있습니다(4.4.3.2).",
    "기준은 사전협의 단계에서 정하며(4.4.3.4), 규정·성능 기준·합의값·표준 가운데 근거를 밝힙니다(4.4.3.5).",
    "정량 기준의 예: 위험의 기대값(금액), 단위 면적당 부상, 점수 체계, 생명 위험 기대값, 화재 손실 비율, 화재 확산 범위(A.5.1.2)."],
  how: ["주 지표 하나와 보조 지표를 정합니다. 외벽은 ‘화재 확산 범위’(A.5.1.2의 (6))가 가장 직접적인 지표입니다.",
    "불확실성을 판정에 넣는 방식을 함께 정합니다(예: 인식론적 표본 중 기준을 만족하는 비율 ≥ 90 %).",
    "기준안(불연 단열재 등)과의 상대 비교(RR_S)를 함께 보고하면 대안 비교 목적에 맞습니다(5.6.1.4).",
    "결과를 본 뒤 기준을 바꾸지 않습니다. 기준을 바꿔야 하면 개념보고서를 고쳐 다시 합의합니다."],
  ex: [`허용기준은 관할기관과의 사전협의에서 다음과 같이 정하였다. 주 지표는 외벽 화재가 발화층 위 2개 층 이상으로 번지는 연간 빈도 F(n ≥ 2)이며, 기준값은 ${sci(best("crit_f2"), 1)}/동·년이다(예시값). 인식론적 불확실성을 반영하기 위해 몬테카를로 표본 가운데 기준을 만족하는 비율이 90 % 이상이면 ‘충족’, 10 % 이하이면 ‘불충족’, 그 사이는 ‘판단 유보’로 판정한다.`,
    "보조 지표로 개인위험도(참고: 영국 HSE R2P2의 10⁻⁶/년 ‘일반적으로 허용’, 10⁻⁴/년 ‘일반인 허용 한계’)와 불연 단열재 기준안 대비 상대위험도 RR_S를 함께 보고한다."],
  check: ["기준값의 근거(규정·합의·표준)를 적었는가", "불확실성을 판정에 넣는 규칙이 분석 전에 정해졌는가"],
}));

body.push(...section("1.5 위험요소", "7.2.2.5, 4.4.2", {
  what: ["평가가 다루는 위험요소(hazard)를 적습니다. 위험의 요소는 노출 대상(4.4.2.1), 화재 자극(열·연기·가스·폭발, 4.4.2.2), 자극을 대상까지 옮기는 전달 현상(4.4.2.3), 자극에 대한 대상의 반응과 기준(4.4.2.4)입니다."],
  how: ["외벽은 표 1-2처럼 네 요소를 한 줄씩 적습니다.",
    "중공층 외단열의 특징(중공층의 굴뚝 효과, 단열재 열분해 가스, 용융 낙하물, 버너 소화 후 자립 연소)을 전달 현상에 적습니다."],
  extra: [label("표 1-2  외벽 위험요소 (예)", INK), table(["요소", "외벽(중공층 외단열)에서의 내용"], [
    ["노출 대상 (4.4.2.1)", "발화층 위 재실자, 소방대, 건물(외벽·세대 내부)"],
    ["화재 자극 (4.4.2.2)", "화염·복사열, 용융·연소 낙하물, 연기·유독가스"],
    ["전달 현상 (4.4.2.3)", "창 분출화염 또는 외부 화원 → 외장재·중공층 → 중공층 수직 확산(굴뚝 효과) → 위층 창으로 재유입; 창→창 도약 확산"],
    ["반응·기준 (4.4.2.4)", "피난 완료(RSET) 전에 화염이 층에 닿는가, 발화층 위로 몇 개 층이 번지는가"],
  ], [24, 76])],
}));

body.push(...section("1.6 시나리오와 군집", "7.2.2.6, 5.1.1.5, 5.5.1.1, 5.5.1.4", {
  what: ["평가에 쓸 시나리오 또는 시나리오 군집을 적고, 비슷한 시나리오를 묶었다면 묶은 근거를 적습니다.",
    "선택한 시나리오가 건물이 겪을 수 있는 모든 시나리오를 대표하는 이유를 적습니다.",
    "비현실적으로 가혹하거나 충분히 드물어 다루지 않은 시나리오는 이유와 함께 적습니다(5.5.1.1). 시나리오를 너무 줄이면 위험을 낮게 평가하게 됩니다(5.5.8)."],
  how: ["외벽을 화염에 노출시키는 초기사건으로 군집합니다: IE1 세대 화재 → 창 분출화염, IE2 외부 발화원(필로티·주차 차량, 쓰레기, 외부 설비) → 외벽 하부. 필요하면 IE3 발코니 화재를 추가합니다.",
    "군집 안에서는 발화층(IE1은 층별 균등), 방호 상태(스프링클러, 차단재, 배리어), 재료 반복시험으로 나눕니다. 군집 빈도는 군집에 든 시나리오 빈도의 합입니다(5.1.1.5.3).",
    "EFFECT(NFPA·Arup) 같은 정성 도구로 먼저 대상 건물을 선별했다면 그 결과와 선별 기준을 함께 적습니다."],
  ex: ["외벽을 화염에 노출시키는 사건을 두 초기사건 군집으로 나누었다. IE1은 세대 화재가 스프링클러로 제어되지 않고 최성기에 이르러 창으로 화염이 분출하는 사건으로, 발화층은 1–15층에 균등하게 분포한다고 보았다. IE2는 필로티·주차 차량, 쓰레기 등 외벽 하부의 외부 발화원이 외벽을 직접 노출시키는 사건으로, 1층에서 시작하며 세대 스프링클러의 영향을 받지 않는다.",
    "공사 중 화재는 사용 단계 평가 범위 밖이므로, 방화(다중 발화)는 비현실적으로 가혹한 시나리오로 보아 제외하였다(5.5.1.1). 발코니 화재는 IE3로 추가할 수 있으나 본 평가에서는 다루지 않았으며, 그 영향은 2.6 한계에 적었다."],
  extra: [label("표 1-3  제외·보류한 시나리오 (예)", INK), table(["시나리오", "판단", "근거"], [
    ["공사 중 화재(단열재 노출)", "범위 밖", "사용 단계만 평가(1.2) → 별도 평가"],
    ["방화 · 다중 발화", "제외", "비현실적으로 가혹(5.5.1.1)"],
    ["발코니 화재", "보류", "IE3로 추가 가능(발화층 균등, 스프링클러 무관)"],
    ["인접 건물 복사열 · 강풍", "보류", "별도 군집으로 추가 가능"],
    ["실내 연기 확산", "범위 밖", "외벽 경로 밖 → 실내 FRA와 결합"],
  ], [30, 14, 56])],
  check: ["제외한 시나리오마다 이유가 있는가", "군집의 대표성(발화층 분포, 반복시험 가중)을 적었는가"],
}));

body.push(...section("1.7 평가 방법", "7.2.2.7, 4.4.4, 5.1.3, 5.5", {
  what: ["사용할 방법과 그 방법이 이 평가에 적합한 이유를 적습니다(7.2.2.7).",
    "방법의 해법, 계산(단위 포함), 일반적이지 않은 식의 출처나 유도를 적습니다(4.4.4.1)."],
  how: ["범주: NFPA 551 표 5.1.2.1의 정량적 방법(사건수목 + 화재 모델).",
    "구성: ① 사건수목(빈도 F) ② 층간 확산 결과 모델(결과 C, 중공층 시험 t₆₀₀ = 층당 확산시간) ③ 위험도 R = Σ F·C(식 5.5.1.2·5.5.1.3) ④ 2중 몬테카를로(불확실성).",
    "적합성: 시험 자료(층당 확산시간, 자립 연소)와 화재 통계(빈도)를 같은 식에서 직접 결합하고, 대안(재료·대책)을 같은 조건에서 비교할 수 있음.",
    "다른 방법을 쓰지 않은 이유(예: 위험 지수·EFFECT는 정성 선별용이라 확정 판정에 쓰지 않음, 5.1.3.1)를 적습니다.",
    "식은 부록 A처럼 기호·단위와 함께 모두 적습니다."],
  ex: ["본 평가는 NFPA 551 표 5.1.2.1의 정량적 방법을 사용하였다. 외벽 노출 사건의 빈도는 사건수목으로, 노출 1회당 결과는 중공층 시험의 t₆₀₀(3000 mm에서 ΔT ≥ 600 K가 시작되는 시각)을 ‘중공층 화재가 한 개 층을 오르는 시간’으로 보는 층간 확산 모델로 계산하였으며, 위험도는 R = Σ F·C로 합산하였다. 입력의 불확실성은 인식론적 2,000 × 우연적 400 표본의 2중 몬테카를로로 전파하였다. 이 방법은 시험 자료와 화재 통계를 같은 식에서 결합하고 재료·대책 대안을 같은 조건에서 비교할 수 있어 평가 목적(1.2)에 적합하다."],
}));

body.push(...section("1.8 자료 출처와 가정", "7.2.2.8, 4.4.5, 6.2", {
  what: ["자료의 출처, 가정과 그 근거를 적습니다(7.2.2.8). 자료가 없어 가정값을 쓴 경우 그렇다고 분명히 밝힙니다(4.4.5.4, 6.2.1.3).",
    "자료가 이 용도·지역에 적용되는지(6.2.2), 자료의 불확실성 원천(6.2.3: 표본 크기, 편향, 수집 기간 등)을 검토합니다."],
  how: ["입력마다 값 · 출처 · 종류(측정 · 통계 · 문헌 · 가정)를 표로 적습니다(부록 B).",
    "외국 통계(예: 미국의 스프링클러 작동 통계)를 쓰면 지역·문화적 편향(6.2.2.3)을 적고 국내 자료로 바꿀 계획을 밝힙니다.",
    "반복시험이 2–3회뿐인 시험값은 표본이 작다는 점과 그 처리 방법(예: 자립 연소 확률을 Jeffreys 추정 (k + ½)/(n + 1)로 둠)을 적습니다."],
  ex: ["재료 입력은 중공층 화재시험(중공층 100 mm, 외단열 100 mm, 30 kW 가스버너 20분, 30분 평가)의 3000 mm t₆₀₀과 버너 소화 후 자립 연소 판정(플래그 B)이며, EPS 가연성은 기록이 4분(240·262 s)뿐이어서 t₆₀₀을 구할 수 없어 정량 평가에서 제외하였다(6.2.1.3). 세대 화재 빈도, 외부 발화원 빈도, 분기 확률, 소방 진압·피난 시각은 절차 시연을 위한 예시값이며, 부록 B에 각각을 대체할 국내 자료(국가화재정보시스템 통계, 관할 소방서 출동 자료 등)를 적었다."],
  check: ["가정값마다 ‘가정’ 표시와 대체 계획이 있는가", "외국 자료의 적용성을 검토했는가"],
}));

// ---------------------------------------------------------------------------------- 2 quantitative results
body.push(H1("2. 정량 결과 문서 (NFPA 551 7.3.3.4) — 분석 후 제출할 내용"));
body.push(P("정량 모델을 쓴 평가의 종합 문서에는 다음 8항목을 넣습니다(7.3.3.4.1–7.3.3.4.8). 각 절의 표·그림은 계산서(xlsx)와 파이썬 결과(facade_fra/results)에서 그대로 가져옵니다."));

const etRows = [
  ["IE1", "세대 화재", "λ₁ × 세대 수", sci(best("lam1") * units, 2), sci(best("lam1") * units, 2)],
  ["", "스프링클러 제어 실패", "× (1 − p_s) (B0: × 1)", sci(best("lam1") * units * (1 - best("p_spr")), 2), sci(best("lam1") * units, 2)],
  ["", "최성기 도달", "× p_FO", sci(best("lam1") * units * (1 - best("p_spr")) * best("p_fo"), 2), sci(best("lam1") * units * best("p_fo"), 2)],
  ["", "창 분출화염 = F₁", "× p_v", sci(F1, 3), sci(PT("PF_in", "B0", "M0").F_IE1_ext, 3)],
  ["", "중공층 노출 → 외벽 + 도약 (M0)", "× p_e0", sci(F1 * best("p_exp0"), 2), sci(PT("PF_in", "B0", "M0").F_IE1_ext * best("p_exp0"), 2)],
  ["IE2", "외부 발화원 화재", "λ₂", sci(best("lam2"), 2), sci(best("lam2"), 2)],
  ["", "외벽 화염 노출 = F₂", "× p_x", sci(F2, 2), sci(F2, 2)],
  ["", "중공층 노출 → 외벽 (M0)", "× p_e0", sci(F2 * best("p_exp0"), 2), sci(F2 * best("p_exp0"), 2)],
];
body.push(...section("2.1 빈도 · 확률 분석 결과", "7.3.3.4.1", {
  what: ["시나리오 또는 군집마다 확률이나 빈도를 적습니다. 확률이면 그 기간을 밝힙니다."],
  how: ["사건수목의 분기와 끝 상태별 빈도를 표 2-1처럼 적고, 단위(/동·년)를 붙입니다.",
    "분기를 서로 독립으로 곱했다면 그 가정을 적습니다(5.3.1.2.2). 독립이 아니면 조건부 확률을 씁니다.",
    "건물 조건(스프링클러 유무 B1/B0)과 대책(M0/M1)에 따라 바뀌는 분기를 구분해 적습니다."],
  ex: [`외벽이 화염에 노출되는 연간 빈도는 IE1(세대 화재 → 창 분출화염) ${sci(F1, 2)}/동·년, IE2(외부 발화원) ${sci(F2, 2)}/동·년이다(스프링클러 설치 건물, 예시 입력). 세대 스프링클러는 IE1만 줄이므로, 스프링클러가 설치된 건물에서는 외부 발화원이 외벽 노출의 더 큰 몫을 차지한다. 스프링클러가 없는 건물(B0)에서는 IE1이 ${sci(PT("PF_in", "B0", "M0").F_IE1_ext, 2)}/동·년으로 10배가 된다. 분기 확률은 서로 독립이라고 보고 곱하였다(5.3.1.2.2).`],
  extra: [label("표 2-1  사건수목 분기와 빈도 (/동·년, 예시 입력)", INK), table(["초기사건", "분기", "분기 확률", "B1 (설치)", "B0 (미설치)"], etRows, [10, 34, 22, 17, 17], { center: [0, 3, 4] })],
}));

const consRows = ORDER.map((m) => {
  if (m === "NC") return [LABEL.NC, "—", "—", f(cl.E_storeys), "—", f(cl.P_ge2), f(cl.E_persons, 1)];
  const ts = D.tests.filter((t) => t.material === m && t.used).map((t) => (t.t600_s === null ? "미도달" : t.t600_s.toFixed(0))).join(" / ");
  const fb = D.flag_b[m];
  return [LABEL[m], ts, `${fb.k}/${fb.n} → ${f(fb.p_best, 3)}`, f(CO(m, "IE1", false).E_storeys), D.grade[m] === "c" ? "불인정" : f(CO(m, "IE1", true).E_storeys), f(CO(m, "IE1", false).P_ge2), f(CO(m, "IE1", false).E_persons, 1)];
});
body.push(...section("2.2 결과(피해) 분석 결과", "7.3.3.4.2", {
  what: ["시나리오 또는 군집마다 결과를 적습니다. 군집을 쓴 경우 그 군집을 대표하는 결과를 어떻게 정했는지 밝힙니다."],
  how: ["재료 입력이 모델에 들어가는 방식을 먼저 적습니다: τ_F = s · t₆₀₀ · h / 3 m (s: 시험 → 실물 보정), 플래그 B → 자립 연소 확률 p_B.",
    "결과 모델의 식(부록 A)과 멈춤 조건(진압 T_FS, 층간 배리어 p_cb, 비자립 재료의 화원 소멸 T_stop)을 적습니다.",
    "군집 대표성: 반복시험 하나하나를 같은 무게의 사건으로, 발화층은 층별 균등(w_j = (N − j)/N)으로 반영했음을 적습니다.",
    "결과는 ‘외벽 노출 1회당 기대값’으로 표 2-2처럼 정리합니다."],
  ex: [`외벽이 노출된 IE1 사건 1회당 발화층 위로 번지는 층 수의 기대값은 PIR(PUR) 가연성 ${f(CO("PIR_C", "IE1", false).E_storeys, 1)}개 층, PF 준불연 내부용 ${f(c1.E_storeys, 1)}개 층, 그 밖의 준불연 ${f(CO("EPS_SC", "IE1", false).E_storeys, 2)}–${f(CO("PF_out", "IE1", false).E_storeys, 2)}개 층이며, 불연 단열재 기준안(창→창 도약만)은 ${f(cl.E_storeys, 2)}개 층이다(점추정: 진압·피난 완료 15분, 분출 지속 20분, s = ${best("s")}). 개구부 차단재와 층간 배리어를 적용하면 준불연 재료는 ${f(CO("EPS_SC", "IE1", true).E_storeys, 2)}–${f(CO("PF_in", "IE1", true).E_storeys, 2)}개 층으로 줄어든다. 가연성 등급은 3 m 중공층이 30 s 만에 전면 화염이 되었고 배리어 효과를 입증하는 시험이 없으므로 대책 효과를 인정하지 않았다(5.1.4.5).`],
  extra: [label("표 2-2  외벽 노출 1회당 결과 (IE1, 점추정)", INK), table(["재료", "시험 t₆₀₀ (s)", "자립 연소 k/n → p_B", "E[확산 층 수] 대책 없음", "배리어·차단재", "P(n ≥ 2)", "E[노출 인원]"], consRows, [22, 15, 15, 13, 11, 10, 12], { center: [1, 2, 3, 4, 5, 6] })],
}));

const riskRows = ORDER.map((m) => {
  const a = PT(m, "B1", "M0"), b = MC(m, "B1", "M0");
  return [LABEL[m], sci(a.R_S, 2), sci(b.R_S_mean, 2), `${sci(b.R_S_p05, 1)} – ${sci(b.R_S_p95, 1)}`, sci(b.F_n2_mean, 2), m === "NC" ? "1" : f(b.RR_S_p50, 0), pct(a.share_IE2)];
});
body.push(...section("2.3 위험도", "7.3.3.4.3", {
  what: ["계산한 위험도는 시나리오 또는 군집마다 확률·빈도와 결과를 곱해 더한 값입니다.",
    "선택한 시나리오가 일어날 수 있는 시나리오 전체를 대표한다고 보는 이유를 함께 적습니다."],
  how: ["식을 적습니다: R_S = F₁ [p_e C₁ + (1 − p_e) C_L] + F₂ p_e C₂ (확산 층 수), 같은 식으로 R_P(노출 인원)와 초과빈도 F(n).",
    "한 재료에 대해 손계산을 보이면 검토자가 각 단계를 확인하기 쉽습니다(8.2.1.2).",
    "표 2-3처럼 점추정과 몬테카를로 평균·구간을 함께 적고, 초과빈도 곡선(그림 2-1)과 위험도 비교(그림 2-2)를 넣습니다.",
    "초기사건별 기여(IE1 · IE2)를 적어 어떤 대책이 효과적인지 보입니다."],
  ex: [`PF 준불연 내부용(스프링클러 설치, 대책 없음)의 외벽 확산 위험도는 R_S = ${sci(F1, 3)} × (${best("p_exp0")} × ${f(c1.E_storeys, 3)} + ${f(1 - best("p_exp0"), 1)} × ${f(cl.E_storeys, 3)}) + ${sci(F2, 2)} × ${best("p_exp0")} × ${f(c2.E_storeys, 3)} = ${sci(pf.R_S, 3)}층/동·년(점추정)이며, 이 가운데 ${pct(pf.share_IE2)}가 외부 발화원(IE2)에서 온다. 발화층 위 2개 층 이상 확산 빈도는 ${sci(pf.F_n2, 2)}/동·년으로 예시 기준 ${sci(best("crit_f2"), 1)}/동·년의 ${Math.round(pf.F_n2 / best("crit_f2"))}배이다.`,
    `두 초기사건은 외벽을 화염에 노출시키는 사용 단계의 경로를 모두 포함하도록 정하였고(1.6), 제외한 시나리오는 표 1-3에 근거와 함께 적었다. 몬테카를로 평균은 ${sci(pfmc.R_S_mean, 2)}층/동·년(90 % 구간 ${sci(pfmc.R_S_p05, 1)} – ${sci(pfmc.R_S_p95, 1)})로, 진압이 늦어지는 꼬리 사건 때문에 점추정보다 크다.`],
  extra: [label("표 2-3  위험도 (스프링클러 설치 건물, 대책 없음)", INK),
    table(["재료", "R_S 점추정", "R_S MC 평균", "R_S 90 % 구간", "F(n ≥ 2) MC 평균", "RR_S (중앙값)", "IE2 비중"], riskRows, [22, 12, 12, 19, 13, 11, 11], { center: [1, 2, 3, 4, 5, 6] }),
    caption("단위: R_S 층/동·년, F /동·년. RR_S = R_S ÷ 불연 기준안의 R_S."),
    image(path.join(FIG, "fig1_exceedance.png"), 640), caption("그림 2-1  초과빈도 곡선 F(발화층 위 확산 ≥ n) — 재료별, M0(대책 없음)과 M1(차단재 + 배리어)"),
    image(path.join(FIG, "fig2_risk.png"), 560), caption("그림 2-2  외벽 확산 위험도 R_S의 평균과 90 % 구간")],
}));

const uncRows = [
  ["불확실성과 변동성 (1)", "빈도 λ₁ · λ₂, 분기 확률(인식론적) / 진압·피난 시각의 사건별 차이(우연적)", "인식론적: 바깥 루프 분포 · 우연적: 안쪽 루프 분포"],
  ["이론·모델 (2)", "3 m 시험 = 1개 층, 층마다 같은 τ_F; 창→창 도약 모델", "보정 계수 s (0.25–1, 로그균등), 실대형 시험으로 확인"],
  ["자료·입력 (3)", "반복시험 2–3회, 예시 통계", "p_B ~ Beta(k + ½, n − k + ½), 반복시험을 경험분포로"],
  ["계산 한계 (4)", "분기 독립, 해석적 기대값", "엑셀·파이썬 독립 구현으로 검증(2.5)"],
  ["시나리오 선택 (5)", "IE1 · IE2만, 발코니·바람 제외", "제외 목록과 영향(2.6)"],
  ["인간 행동 (6)", "피난 개시·속도, 재실 유지", "RSET 분포(중앙값 15분, 로그 σ 0.5)"],
  ["위험 인식·가치 (7)", "허용기준값과 신뢰수준", "예시값으로 명시, 사전협의로 확정"],
];
const keys = ["s", "lam2", "p_x2", "t_fs", "p_cb", "p_exp1"];
const sym = { s: "s", lam2: "λ₂", p_x2: "p_x", t_fs: "T_FS", p_cb: "p_cb", p_exp1: "p_e1" };
const sensRows = keys.map((k) => [sym[k], ...["M0", "M1"].flatMap((g) => ["PIR_SC", "PF_in"].map((m) => { const v = SE(g, m, k).rho_R_S; return (v >= 0 ? "+" : "−") + Math.abs(v).toFixed(2); }))]);
body.push(...section("2.4 불확실성 분석", "7.3.3.4.4, 4.5, 5.5.6, A.4.5, A.5.5.6", {
  what: ["확률·빈도·결과 예측의 불확실성 원천과 그 처리 방법을 적습니다.",
    "불확실성 원천의 예(A.4.5): 불확실성과 변동성, 이론·모델, 자료·입력, 계산 한계, 시나리오 선택, 인간 행동, 위험 인식·가치.",
    "우연적 불확실성(무작위 변동)과 인식론적 불확실성(지식 부족)을 구분합니다(A.5.5.6)."],
  how: ["표 2-4처럼 원천별 내용과 처리를 적습니다.",
    "전파 방법(2중 몬테카를로: 바깥 = 인식론적, 안쪽 = 우연적), 표본 수, 난수 씨앗을 적습니다.",
    "민감도(순위상관 등)로 결론을 좌우하는 입력을 밝히고, 자료 보강 우선순위로 연결합니다(그림 2-3, 표 2-5).",
    "판정은 1.4에서 정한 신뢰수준 규칙으로 합니다(표 2-6)."],
  ex: [`불확실성은 2중 몬테카를로로 전파하였다. 바깥 루프는 빈도·분기 확률·보정 계수 s·시간 중앙값·자립 연소 확률 p_B의 인식론적 불확실성 ${D.n_outer.toLocaleString()}개 표본, 안쪽 루프는 진압·피난 시각, 도약 시간, 분출 지속, 반복시험, 자립 연소 여부의 우연적 변동 ${D.n_inner}개 표본(라틴 하이퍼큐브)이다. 결론에 가장 큰 영향을 준 입력은 대책이 없을 때 s(ρ = ${rho(SE("M0", "PIR_SC", "s").rho_R_S)}, PIR(PUR) 준불연)와 외부 발화원 빈도 λ₂(ρ = ${rho(SE("M0", "PF_in", "lam2").rho_R_S)}, PF 내부용)이며, 대책을 적용하면 배리어 효과 p_cb(ρ = ${rho(SE("M1", "PF_in", "p_cb").rho_R_S)}, PF 내부용)가 커졌다. 따라서 실대형 시험(s 보정), 외부 발화원 화재 통계, 배리어를 넣은 중공층 시험이 판정을 확정하는 데 가장 필요한 자료이다.`],
  extra: [label("표 2-4  불확실성 원천과 처리 (A.4.5의 7개 원천)", INK), table(["원천", "이 평가에서의 내용", "처리"], uncRows, [20, 44, 36]),
    new Paragraph({ spacing: { after: 80 } }),
    image(path.join(FIG, "fig3_sensitivity.png"), 600), caption("그림 2-3  인식론적 입력과 log R_S의 순위상관 ρ (스프링클러 설치 건물)"),
    label("표 2-5  민감도 상위 입력 (ρ, log R_S)", INK),
    table(["입력", "M0 · PIR(PUR) 준불연", "M0 · PF 내부용", "M1 · PIR(PUR) 준불연", "M1 · PF 내부용"], sensRows, [16, 21, 21, 21, 21], { center: [0, 1, 2, 3, 4] }),
    new Paragraph({ spacing: { after: 80 } }),
    label("표 2-6  허용기준 판정 (F(n ≥ 2) ≤ 기준값을 만족하는 인식론적 표본 비율, 신뢰수준 90 %)", INK),
    table(["재료", "M0 충족 비율", "M0 판정", "M1 충족 비율", "M1 판정"], ORDER.map((m) => [LABEL[m], pct(confOf(m, "M0"), 1), verdictOf(m, "M0"), pct(confOf(m, "M1"), 1), verdictOf(m, "M1")]), [30, 17, 18, 17, 18], { center: [1, 2, 3, 4],
      fill: (k, i) => { const v = i === 2 ? verdictOf(ORDER[k], "M0") : i === 4 ? verdictOf(ORDER[k], "M1") : null; return v === "충족" ? "E2F3E2" : v === "불충족" ? "F8DEDE" : v ? "FEF2D3" : null; } })],
}));

body.push(...section("2.5 소프트웨어와 모델 평가", "7.3.3.4.5, 6.2.4.2, 5.5.9, 8.2", {
  what: ["사용한 모델이 이 상황에 적합한 이유를 적습니다.",
    "계산 코드는 작성자·버전·입출력 형식·실행 환경과, 검증·확인(V&V) 여부를 밝힙니다(6.2.4.2). 결과·빈도 추정에 쓴 방법의 확인 근거를 제출합니다(5.5.9)."],
  how: ["코드와 버전: facade_fra/qfra.py(Python, NumPy·pandas), 외벽_QFRA_계산서.xlsx(엑셀 수식).",
    "검증(8.2.1): 두 구현의 결과 비교(대체 계산 재현), 손계산 확인, 몬테카를로가 변동 0에서 점추정과 같은지.",
    "확인(8.2.2): 재료 입력은 시험값 자체이므로 3 m 안에서는 일치; ‘3 m = 1개 층’ 외삽은 실대형 시험·PyroSim(FDS) 계산과 비교해야 함.",
    "적용 범위: 준불연·탄화형과 용융형(EPS)은 반복시험 값으로, 가연성은 시험값만(대책 효과 불인정), 층고는 h/3 m 비율로."],
  ex: [`계산은 파이썬 스크립트 qfra.py와 엑셀 계산서로 따로 구현하였다. 엑셀 수식 ${VER.n_formulas.toLocaleString()}개로 계산한 점추정 ${VER.n_cases}경우 × 5지표(${VER.n_values}개 값)는 파이썬 결과와 최대 상대차 ${sci(VER.max_rel_diff, 2)}로 일치하였다(8.2.1.1). 몬테카를로의 우연적 변동을 0으로 두면 결과가 점추정과 같음을 확인하였다. 재료 입력은 중공층 시험값이므로 3 m 범위에서는 시험과 일치하지만, 한 층씩 같은 시간으로 오른다는 외삽은 아직 실대형 시험으로 확인되지 않았다(2.6).`],
}));

body.push(...section("2.6 한계", "7.3.3.4.6", {
  what: ["모델과 분석 범위에서 오는 한계를 모두 적습니다."],
  how: ["가장 영향이 큰 한계부터 적고, 각 한계가 결과를 높게 또는 낮게 만드는 방향을 함께 적습니다.",
    "한계마다 해소 방법(추가 시험·자료)을 짝지어 적습니다."],
  ex: ["본 평가의 주요 한계는 다음과 같다. ① 3 m 중규모 중공층 시험을 ‘1개 층 확산’으로 보고 층마다 같은 시간으로 외삽하였다(보정 계수 s = 0.25–1 가정). ② 빈도·분기 확률·시간은 예시값으로 국내 통계를 반영하지 않았다. ③ 분기를 독립으로 보았고 수평 확산, 바람, 낙하물에 의한 하부 착화는 다루지 않았다. ④ 노출은 피난 완료 전 화염 도달로 정의하여 연기 영향은 포함하지 않았다. ⑤ 반복시험이 2–3회로 적고, EPS 가연성은 자료가 없다."],
}));

const concl = ORDER.filter((m) => m !== "NC").map((m) => `${LABEL[m]} ${verdictOf(m, "M0")} → ${verdictOf(m, "M1")}`).join(", ");
body.push(...section("2.7 결론", "7.3.3.4.7", {
  what: ["평가 결과를 허용기준과 비교한 판정(해당하면 합격·불합격)을 요약합니다.",
    "평가 목적과 목표를 어느 정도 달성했는지, 결과가 이 목적에 적절하고 충분한지 적습니다."],
  how: ["판정은 1.4의 규칙 그대로 적습니다(표 2-6).",
    "조건부 결론은 조건(대책, 관리 조건)을 함께 적고, 그 조건을 3장 운영·유지관리에 넘깁니다.",
    "판단 유보가 있으면 판정을 확정하는 데 필요한 자료를 적습니다."],
  ex: [`예시 입력과 예시 기준(F(n ≥ 2) ≤ ${sci(best("crit_f2"), 1)}/동·년, 신뢰수준 90 %)으로 판정한 결과(대책 없음 → 차단재 + 배리어)는 ${concl}이다. 준불연 4종은 개구부 불연 차단재와 층간 캐비티 배리어를 함께 적용할 때 기준을 충족하며(충족 비율 ${pct(minConfM1)} 이상), 이 결론은 배리어·차단재가 설계대로 시공·유지되는 조건(3장)에서 유효하다. 가연성 PIR(PUR)은 대책 효과를 입증하는 시험 자료가 없어 기준을 충족하지 못한다. 대책 없는 준불연 3종의 ‘판단 유보’를 확정하려면 실대형 시험에 의한 s 보정과 외부 발화원 화재 통계가 필요하다.`],
}));

body.push(...section("2.8 참고문헌", "7.3.3.4.8, 6.2.1", {
  what: ["입력 자료의 출처와 그 자료가 이 평가에 적합한 이유를 적습니다. 도면, 보고서, 매뉴얼, 간행물, 규정, 표준의 판·발행일을 적습니다."],
  how: ["공개 자료는 제목, 저자, 쪽·표·그림 번호, 발행처, 발행일을 적습니다(6.2.1.1). 비공개 자료는 제공자와 접근 방법을 적습니다(6.2.1.2).",
    "계산 파일은 버전과 날짜를 붙여 보관합니다(6.2.1.4)."],
  ex: ["NFPA 551, Guide for the Evaluation of Fire Risk Assessments, 2022 Edition, National Fire Protection Association.",
    "HSE, Reducing risks, protecting people (R2P2), Health and Safety Executive, 2001.",
    "NFPA · Arup, EFFECT (Exterior Façade Fire Evaluation Comparison Tool), 2018 — [사용자 가이드 판·발행일].",
    "[기관명], 중공층 외단열 화재시험 보고서(중공층 100 mm, 외단열 100 mm, 30 kW 20분), [발행일].",
    "KS F 8414, [외벽 실물모형 화재시험 표준명], [판].",
    "소방청, 국가화재정보시스템 화재 통계, [조회 기간·조회일]."],
}));

// ---------------------------------------------------------------------------------- 3 O&M
body.push(H1("3. 운영 · 유지관리 매뉴얼 (NFPA 551 7.4)과 이행 (7.5)"));
body.push(P("운영·유지관리 매뉴얼은 평가에서 내린 결정이 계속 유효하도록 지켜야 할 조건(사용 제한, 점검·시험·유지관리 요건)을 건물 소유자·관리자가 쓸 수 있게 적은 문서입니다(7.4.1). 외벽 평가는 평가 때의 재료·차단재·배리어·발화원 관리가 그대로일 때만 유효합니다."));
body.push(...section("3.1 한계 · 가정 목록과 관리 대책", "7.4.2", {
  what: ["평가에 쓴 단순화와 그에 따른 한계·가정을 목록으로 적고(7.4.2.1), 그 가정을 지키는 관리 대책(행정 프로그램, 설계 요소)을 적습니다(7.4.2.2)."],
  how: ["가정마다 ‘지켜야 할 조건’으로 바꿔 씁니다: 재료, 차단재·배리어, 발화원 관리, 스프링클러.",
    "조건이 깨지면 바뀌는 입력(p_e, p_cb, λ₂, p_s)을 적어 두면 재평가 범위가 분명해집니다."],
  ex: ["본 평가는 다음 조건에서 유효하다. ① 외장재·단열재는 평가한 제품과 같은 구성일 것. ② 개구부 불연 차단재와 층간 캐비티 배리어가 설계대로 시공되고 관통부가 없을 것(p_e1, p_cb). ③ 필로티·외벽 하부에 차량·쓰레기 등 가연물을 두지 않을 것(λ₂·p_x). ④ 세대 스프링클러가 작동 상태로 유지될 것(p_s)."],
}));
body.push(...section("3.2 변경 관리", "7.4.3", {
  what: ["평가는 사용한 입력 조건에서만 유효하므로, 어떤 변경이 새 평가를 필요로 하는지와 평가가 유효한 조건을 적습니다(7.4.3.2).",
    "변경으로 위험이 모르게 바뀌지 않도록 교육, 절차, 변경 승인, 감사 같은 통제를 둡니다(7.4.3.3)."],
  how: ["재평가 계기를 목록으로: 외장 교체, 리모델링·창호 교체, 설비 관통부 추가, 발코니 확장, 필로티 용도 변경.",
    "변경 전 영향 검토와 기록 방법을 적습니다."],
  ex: ["외장재 교체, 창호 교체, 외벽 관통부 추가, 발코니 확장, 필로티 주차장 용도 변경은 이 평가의 입력을 바꾸므로, 시행 전에 [관리주체]가 영향 검토를 요청하고 필요하면 재평가한다. 변경 내용과 검토 결과는 평가 문서의 변경 이력에 기록한다."],
}));
body.push(...section("3.3 점검 · 시험 · 유지관리", "7.4.4", {
  what: ["평가가 근거로 삼은 점검·시험·유지관리(ITM) 요건을 적습니다. ITM은 인정한 설비·구성요소의 작동성과 가용성에 영향을 줍니다(7.4.4.2)."],
  how: ["대상별 주기와 담당: 배리어·차단재(관통부·시공 하자), 스프링클러(작동 시험), 필로티(적치물·주차).",
    "점검 결과로 p_s · p_e · p_cb 가정을 갱신하는 절차를 적습니다."],
  ex: ["층간 캐비티 배리어와 개구부 차단재는 [주기]마다 관통부와 시공 하자를 점검하고, 세대 스프링클러는 [주기]마다 작동 시험을 한다. 필로티 주차·적치물은 [주기]마다 점검한다. 점검 결과는 [기관]에 제출하며, 결함률이 가정을 넘으면 해당 분기 확률을 고쳐 재계산한다."],
}));
body.push(...section("3.4 이행", "7.5", {
  what: ["관리 조건과 설비가 제대로 유지되도록 하는 이행 수단을 적습니다(7.5.1). 사후 발견(제품 리콜, 예상하지 못한 거동)이 있으면 평가를 갱신하는 방법을 적습니다(7.5.4)."],
  how: ["점검 기록 제출 주기·담당, 미이행 시 조치(관할 규정 연계), 재평가 계기를 관리 규약·계약에 반영."],
  ex: ["점검 기록은 [주기]마다 관할기관에 제출하며, 미이행 시 [관할 규정]에 따른다. 외장 제품의 리콜이나 시공 하자가 발견되면 [기간] 안에 재평가한다."],
}));

// ---------------------------------------------------------------------------------- 4 review
body.push(H1("4. 검토 대응 (NFPA 551 8장)"));
body.push(...section("4.1 검토 방식", "8.1", {
  what: ["관할기관은 평가 문서를 직접 검토하거나(8.1.1), 제3자에게 맡깁니다(8.1.2): 동료 검토(제3자가 의견을 내고 관할기관이 판단, 8.1.2.1) 또는 계약 검토(제3자가 승인·보완·반려까지 결정, 8.1.2.2)."],
  how: ["개념보고서 단계에서 검토 방식과 검토자를 정해 둡니다(4.3.2.1).", "동료 검토자는 평가에 관여하지 않은 같은 수준의 전문가여야 합니다."],
}));
body.push(...section("4.2 검증 · 확인 요약", "8.2", {
  what: ["검증(8.2.1)은 수식과 계산이 맞는지: 대체 계산에 의한 재현, 단계별 확인, 표본 감사.",
    "확인(8.2.2)은 결과가 실제 위험을 대표하는지: 대체 모델과 비교, 시험 결과와 비교, 완성된 건물에서의 실증."],
  how: ["2.5의 검증 결과와 확인 계획을 요약하고, 남은 확인 항목(실대형 시험, PyroSim 비교)의 일정을 적습니다."],
}));
body.push(H2("4.3 검토 질문 대응표 (8.3)"));
body.push(P("검토자는 다음 9개 질문으로 평가가 제대로 되었는지 확인합니다. 답이 문서의 어디에 있는지 미리 표로 붙여 두면 검토가 빨라집니다."));
body.push(table(["8.3", "검토 질문", "답하는 위치"], [
  ["(1)", "평가 목적이 명시되었는가", "1.2 범위와 목적"], ["(2)", "평가 범위가 명시되었는가", "1.2 범위와 목적"],
  ["(3)", "방법과 그 적합성 근거가 있는가", "1.7 평가 방법"], ["(4)", "한계가 명시되었는가", "2.6 한계"],
  ["(5)", "분석 결과가 포함되었는가", "2.1–2.4, 계산서 5_위험도 · 6_불확실성(MC)"], ["(6)", "결론이 포함되었는가", "2.7 결론"],
  ["(7)", "운영·유지관리 매뉴얼이 있는가", "3.1"], ["(8)", "변경 관리 지침이 있는가", "3.2"], ["(9)", "점검·시험·유지관리 프로그램이 있는가", "3.3"],
], [8, 52, 40], { center: [0] }));

// ---------------------------------------------------------------------------------- appendix A equations
body.push(H1("부록 A. 계산식"));
body.push(...bullets([
  [["위험도 (5.5.1.2·5.5.1.3)  ", { bold: true }], "R_t = Σ C_i F_i (i: 시나리오),  결과가 여럿이면 R_t = Σ Σ C_ij F_i (j: 결과 종류)"],
  [["외벽 노출 빈도  ", { bold: true }], "F₁ = λ₁ · U · (1 − p_s) · p_FO · p_v,   F₂ = λ₂ · p_x   (U: 세대 수, 스프링클러가 없으면 p_s = 0)"],
  [["층당 확산시간  ", { bold: true }], "τ_F = s · t₆₀₀ · h / 3 m   (t₆₀₀: 3000 mm에서 ΔT ≥ 600 K 시작 시각, 600 K 미도달이면 τ_F = ∞)"],
  [["중공층 경로의 층 수  ", { bold: true }], "n_F = 내림[ min(T_FS, T_stop) ÷ τ_F ],  T_stop = t_e (비자립) 또는 ∞ (자립, 확률 p_B)"],
  [["배리어 · 도약  ", { bold: true }], "a_j = (1 − p_cb)ʲ (j ≤ n_F),  b_j = p_Lʲ (j ≤ n_L),  n_L = 내림[T_FS ÷ τ_L]"],
  [["확산 확률  ", { bold: true }], "P(n ≥ j) = w_j [1 − (1 − a_j)(1 − b_j)],  w_j = (N − j)/N (IE1, 발화층 균등) 또는 1 (IE2)"],
  [["노출 인원  ", { bold: true }], "층 j에 먼저 닿은 경로의 도달시각(j τ_F 또는 j τ_L) < RSET이면 n_a × 세대당 재실자"],
  [["위험도 (외벽)  ", { bold: true }], "R_S = F₁ [p_e C₁ + (1 − p_e) C_L] + F₂ p_e C₂  (C₁: IE1 외벽 + 도약, C_L: 도약만, C₂: IE2 외벽), R_P·F(n)도 같은 식"],
  [["자립 연소 확률  ", { bold: true }], "p_B = (k + ½)/(n + 1)  (k: 플래그 B 시험 수, n: 반복시험 수), 불확실성은 Beta(k + ½, n − k + ½)"],
  [["개인위험도  ", { bold: true }], "IR = p_d · R_P ÷ 건물 재실자"],
]));

// ---------------------------------------------------------------------------------- appendix B inputs
body.push(H1("부록 B. 입력값과 대체할 자료"));
const pkeys = ["N", "units", "occ", "n_aff", "lam1", "lam2", "p_spr", "p_fo", "p_vent", "p_x2", "p_exp0", "p_exp1", "p_cb", "p_L", "tau_L", "t_fs", "rset", "s", "p_d", "crit_f2"];
const val = (v) => (v === null || v === undefined ? "—" : Math.abs(v) < 0.01 && v !== 0 ? sci(v, 2) : String(v));
body.push(table(["기호", "항목", "값", "범위", "근거 · 대체할 자료"], pkeys.map((k) => {
  const p = PRM[k];
  return [p.symbol, `${p.item} (${p.unit})`, val(p.best), p.low === null || p.low === undefined ? "고정" : `${val(p.low)} – ${val(p.high)}`, p.basis];
}), [8, 30, 10, 15, 37], { center: [0, 2, 3] }));
body.push(caption("범위는 몬테카를로의 인식론적 범위(확률: 삼각분포, 빈도·s: 로그균등). 사건별 변동: 진압·피난 시각 로그 σ 0.5, 도약 시간 0.3, 분출 지속 IE1 10–30분, IE2 5–20분."));

// ================================================================================================ document
const doc = new Document({
  creator: "KICT 중공층 외단열 화재 분석",
  title: "외벽 정량적 화재위험성평가 — 보고서 서술 양식",
  description: "NFPA 551(2022) 7장 문서 요건을 외벽(중공층 외단열) 정량 평가에 맞춘 항목별 작성 방법과 예문",
  styles: {
    default: { document: { run: { font: FONT, size: 20, color: INK } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 30, bold: true, color: INK }, paragraph: { spacing: { before: 120, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 25, bold: true, color: INK }, paragraph: { spacing: { before: 320, after: 100 }, outlineLevel: 1,
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "D1D5DB", space: 4 } } } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 22, bold: true, color: INK }, paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 240 } } } }] },
      { reference: "checks", levels: [{ level: 0, format: LevelFormat.BULLET, text: "□", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 260 } } } }] },
    ],
  },
  sections: [{
    properties: { page: { margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "외벽 정량적 화재위험성평가 — 보고서 서술 양식 (NFPA 551 7장)", font: FONT, size: 16, color: INK2 })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: ["– ", PageNumber.CURRENT, " –"], font: FONT, size: 16, color: INK2 })] })] }) },
    children: body,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log("wrote", OUT);
});
