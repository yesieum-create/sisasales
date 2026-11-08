const pptxgen = require('pptxgenjs');
const fs = require('fs');
const SK = '/root/.claude/skills/synced/76ce9fd3-5eb1-44df-8594-3459122bf824_7dd8e0a0-5dbb-4f73-8e85-9190543b3f81/pptx';
const { applyTheme } = require(SK + '/scripts/apply_theme.js');
const S = '/tmp/claude-0/-home-user-sisasales/82ad116b-b749-5857-911d-f0676cb5351b/scratchpad';
const OUT = '/home/user/sisasales/reports/unaccredited-intl-schools-2026-deck.pptx';

const F = 'Pretendard';
const THEME = {
  name: 'SISA Company Style', headFontFace: F, bodyFontFace: F,
  colors: { dk1: '172536', lt1: 'FFFFFF', dk2: '102B49', lt2: 'EAF4FC',
    accent1: '2447C9', accent2: '83C9F2', accent3: '5D6B78', accent4: 'C0392B', accent5: 'D68910', accent6: '1E8449',
    hlink: '2447C9', folHlink: '5D6B78' },
};
const H = { deep: '102B49', royal: '2447C9', sky: '83C9F2', lblue: 'EAF4FC', lgray: 'F4F6F8', text: '172536', muted: '5D6B78', div: 'D9E2EB', red: 'C0392B', amber: 'D68910', green: '1E8449', white: 'FFFFFF' };

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5
pres.theme = { headFontFace: F, bodyFontFace: F };
pres.title = '한국 비인가 국제학교의 현재와 미래';
pres.author = 'SISA Company';
const C = pres.SchemeColor;
const LOGO = 'image/png;base64,' + fs.readFileSync(S + '/logo.png').toString('base64');
const W = 13.333, MX = 0.6, CW = W - 2 * MX;

// ---------- Layouts ----------
pres.defineSlideMaster({
  title: 'COVER', background: { color: H.deep },
  objects: [
    { image: { x: W - 0.6 - 1.99, y: 0.45, w: 1.99, h: 0.5, data: LOGO } },
    { placeholder: { options: { name: 'title', type: 'title', x: MX, y: 2.3, w: 9.5, h: 1.8, fontFace: F, fontSize: 40, bold: true, color: C.background1, valign: 'top', align: 'left', margin: 0 }, text: '' } },
    { placeholder: { options: { name: 'body', type: 'body', x: MX, y: 4.25, w: 9.5, h: 0.8, fontFace: F, fontSize: 18, color: 'D6E6F5', valign: 'top', align: 'left', margin: 0 }, text: '' } },
  ],
});
pres.defineSlideMaster({
  title: 'DIVIDER', background: { color: H.deep },
  objects: [
    { image: { x: W - 0.6 - 1.99, y: 0.45, w: 1.99, h: 0.5, data: LOGO } },
    { placeholder: { options: { name: 'title', type: 'title', x: MX, y: 2.9, w: 11, h: 1.2, fontFace: F, fontSize: 34, bold: true, color: C.background1, valign: 'top', align: 'left', margin: 0 }, text: '' } },
    { placeholder: { options: { name: 'body', type: 'body', x: MX, y: 4.1, w: 11, h: 0.8, fontFace: F, fontSize: 18, color: 'D6E6F5', valign: 'top', align: 'left', margin: 0 }, text: '' } },
  ],
});
pres.defineSlideMaster({
  title: 'CONTENT', background: { color: H.white },
  margin: [0.5, 0.6, 0.6, 0.6],
  objects: [
    { text: { text: 'SISA COMPANY · 한국 비인가 국제학교의 현재와 미래 · v1.0', options: { x: MX, y: 7.05, w: 8, h: 0.3, fontFace: F, fontSize: 10, color: C.accent3, margin: 0 } } },
    { placeholder: { options: { name: 'title', type: 'title', x: MX, y: 0.62, w: CW, h: 0.9, fontFace: F, fontSize: 26, bold: true, color: C.text2, valign: 'middle', align: 'left', margin: 0 }, text: '' } },
    { placeholder: { options: { name: 'body', type: 'body', x: MX, y: 1.52, w: CW, h: 0.38, fontFace: F, fontSize: 14, color: C.accent3, valign: 'middle', align: 'left', margin: 0 }, text: '' } },
  ],
  slideNumber: { x: W - 0.6 - 0.8, y: 7.05, w: 0.8, h: 0.3, fontFace: F, fontSize: 10, color: H.muted, align: 'right' },
});

// ---------- helpers ----------
let N = 0;
const name = (s) => `${s}-${++N}`;
const T = (slide, text, o) => slide.addText(text, Object.assign({ isTextBox: true, fontFace: F, color: C.text1, margin: 0, valign: 'middle', objectName: name('text') }, o));
const R = (slide, o, shape) => slide.addShape(shape || pres.shapes.RECTANGLE, Object.assign({ line: { type: 'none' }, objectName: name('shape') }, o));

function content(section, tag, title, sub) {
  const s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: section });
  T(s, tag, { x: MX, y: 0.3, w: CW, h: 0.3, fontSize: 11, bold: true, color: C.accent1, charSpacing: 1 });
  s.addText(title, { placeholder: 'title' });
  if (sub) s.addText(sub, { placeholder: 'body' });
  return s;
}
function src(s, text) { T(s, '출처: ' + text, { x: MX, y: 6.62, w: CW, h: 0.36, fontSize: 10, color: C.accent3, valign: 'top' }); }
function insight(s, y, runs, h) {
  R(s, { x: MX, y, w: CW, h: h || 0.62, fill: { color: C.background2 } });
  T(s, runs, { x: MX + 0.3, y, w: CW - 0.6, h: h || 0.62, fontSize: 15, bold: true, color: C.text2 });
}
const pill = (t, col) => ({ text: t, options: { bold: true, color: col } });

// table helpers
const hdr = (arr, align) => arr.map((t, i) => ({ text: t, options: { bold: true, color: H.white, fill: { color: H.royal }, align: (align && align[i]) || 'left' } }));
function table(s, rows, o) {
  s.addTable(rows, Object.assign({ x: MX, fontFace: F, fontSize: 13, color: H.text, valign: 'middle', border: { type: 'solid', pt: 0.75, color: H.div }, margin: [0.06, 0.1, 0.06, 0.1], objectName: name('table') }, o));
}
const cell = (t, opt) => ({ text: t, options: opt || {} });

// ================= 1. COVER =================
pres.addSection({ title: '표지' });
let s = pres.addSlide({ masterName: 'COVER', sectionTitle: '표지' });
T(s, 'SISA COMPANY · STRATEGY DECK', { x: MX, y: 1.7, w: 8, h: 0.35, fontSize: 13, bold: true, color: C.accent2, charSpacing: 2 });
s.addText('한국 비인가 국제학교의\n현재와 미래', { placeholder: 'title' });
s.addText('단속 국면의 ‘학교 아닌 국제학교’ 시장 진단, 2027–2030 시나리오, 시사컴퍼니 시사점', { placeholder: 'body' });
R(s, { x: MX, y: 5.45, w: 1.4, h: 0.07, fill: { color: C.accent1 } });
R(s, { x: MX + 1.4, y: 5.45, w: 0.5, h: 0.07, fill: { color: C.accent2 } });
T(s, 'Version 1.0   ·   2026. 11. 8   ·   대표이사 보고', { x: MX, y: 5.7, w: 8, h: 0.35, fontSize: 13, color: 'B9CDE0' });
s.addNotes('표지. 본 덱은 같은 제목의 A4 보고서(v1.0)를 경영진 보고용으로 요약한 것입니다.');

// ================= 2. EXEC SUMMARY =================
pres.addSection({ title: 'Executive Summary' });
s = content('Executive Summary', 'EXECUTIVE SUMMARY', '학교형은 줄지만, 영어 몰입·해외진학 수요는 합법 채널로 옮겨 갈 뿐 사라지지 않는다', '비인가 국제학교는 ‘성장 시장’에서 ‘규제 리스크 시장’으로 전환되었다');
const kp = [
  ['97곳', '교육부 점검에서 확인된\n미인가 국제학교', '2025.12–2026.3, 점검 213곳 중', true],
  ['74%', '수도권 집중도\n(경기 39 · 서울 33)', '교육부 파악, 2026.6 보도'],
  ['4,150만원', '주요 30곳 고등부\n연평균 학비', '사교육걱정없는세상, 2026.2'],
  ['0건', '폐쇄명령\n집행 실적', '2026.10 보도 기준'],
];
const kw = (CW - 0.3 * 3) / 4;
kp.forEach((k, i) => {
  const x = MX + i * (kw + 0.3), y = 2.15, dark = k[3];
  R(s, { x, y, w: kw, h: 1.95, fill: { color: dark ? C.text2 : 'F4F6F8' } });
  T(s, k[0], { x, y: y + 0.15, w: kw, h: 0.75, fontSize: 36, bold: true, align: 'center', color: dark ? C.background1 : C.accent1 });
  T(s, k[1], { x, y: y + 0.9, w: kw, h: 0.65, fontSize: 14, align: 'center', color: dark ? 'D6E6F5' : C.text1 });
  T(s, k[2], { x, y: y + 1.55, w: kw, h: 0.3, fontSize: 10, align: 'center', color: dark ? '9FB6CC' : C.accent3 });
});
const ex = [
  ['현재', '정부 공식 통계는 2014년 이후 공백. 2026년 점검으로 처음 97곳 확인, 민간 추정은 120~130곳·재학생 약 2.6만 명'],
  ['전환점', '교육부(2026.4.29) “영어로 학교 형태 운영 시 학원 등록 여부와 무관하게 폐쇄 대상”. 다만 강제력이 약해 오후반·온라인으로 우회 확산'],
  ['미래', '기본 시나리오(판단) — 학교형 축소 + 수요의 합법 채널 이동. 인가학교 ↔ 합법 보완교육으로 양극화'],
];
ex.forEach((e, i) => {
  const y = 4.35 + i * 0.75;
  R(s, { x: MX, y: y + 0.08, w: 1.3, h: 0.5, fill: { color: C.accent1 } }, pres.shapes.ROUNDED_RECTANGLE);
  T(s, e[0], { x: MX, y: y + 0.08, w: 1.3, h: 0.5, fontSize: 14, bold: true, color: C.background1, align: 'center' });
  T(s, e[1], { x: MX + 1.55, y, w: CW - 1.55, h: 0.66, fontSize: 15 });
});
s.addNotes('핵심: 단속의 대상은 학교 형태이지 수요가 아니다. 수치는 공개 보도 기준이며 출처는 마지막 장에 정리.');

// ================= 3. DEFINITION =================
pres.addSection({ title: 'Situation' });
s = content('Situation', '01 · SITUATION — 정의와 법적 지위', '‘국제학교’라는 간판은 같지만, 법적으로는 네 개의 서로 다른 세계다', '국내 학력이 인정되는 정식 경로는 좁고, 비인가 시설은 학원법·평생교육법 틈새에서 운영된다');
table(s, [
  hdr(['구분', '근거·설립 주체', '내국인 입학', '국내 학력', '규모(보도 기준)'], [0, 0, 0, 'center', 0]),
  [cell('외국교육기관', { bold: true }), '경제자유구역·강원특별법, 비영리법인', '정원 30% 원칙(+20%p, 강원 50%)', cell('인정', { align: 'center', bold: true, color: H.green }), cell('인가 국제학교 7~8곳\n(제주·송도·대구 등)', { rowspan: 2 })],
  [cell('제주 국제학교', { bold: true }), '제주특별법', '제한 없음', cell('인정', { align: 'center', bold: true, color: H.green })],
  [cell('외국인학교', { bold: true }), '초·중등교육법 §60의2', '해외거주 3년 이상 등 제한적', cell('조건부', { align: 'center', bold: true, color: H.amber }), '인가 외국인·국제학교 약 29곳(ISC 2023)'],
  [cell('비인가 국제학교', { bold: true, fill: { color: H.lblue } }), cell('학원법상 어학원 / 평생교육시설 / 무등록', { fill: { color: H.lblue } }), cell('제한 없음(사실상 내국인 중심)', { fill: { color: H.lblue } }), cell('불인정', { align: 'center', bold: true, color: H.red, fill: { color: H.lblue } }), cell('97곳(교육부) ~ 130곳(KBS·ISC)', { bold: true, fill: { color: H.lblue } })],
], { y: 2.1, w: CW, colW: [2.1, 3.0, 2.9, 1.2, 2.93], rowH: 0.5 });
const why = [['초·중등교육법', '인가 없이 학교 명칭·형태(전일제·학년제·졸업장) 운영 시 위반'], ['대안교육기관법 §5②', '외국대학 입학 목적·외국어 주수업 시설은 등록 자체 불가'], ['학력·보호 공백', '검정고시 필수, 학교폭력 공적 절차·교원자격 기준 미적용']];
const ww = (CW - 0.6) / 3;
why.forEach((w, i) => {
  const x = MX + i * (ww + 0.3);
  R(s, { x, y: 4.95, w: ww, h: 1.0, fill: { color: 'F4F6F8' } });
  R(s, { x: x + 0.2, y: 5.12, w: 0.36, h: 0.36, fill: { color: C.accent1 } }, pres.shapes.OVAL);
  T(s, String(i + 1), { x: x + 0.2, y: 5.12, w: 0.36, h: 0.36, fontSize: 13, bold: true, color: C.background1, align: 'center' });
  T(s, w[0], { x: x + 0.7, y: 5.08, w: ww - 0.85, h: 0.42, fontSize: 15, bold: true, color: C.text2 });
  T(s, w[1], { x: x + 0.2, y: 5.5, w: ww - 0.4, h: 0.42, fontSize: 13 });
});
insight(s, 6.1, [{ text: '시사점 — 본질은 ‘불법 운영’이 아니라 ' }, { text: '수요는 있는데 합법 공급 경로가 없는 제도 공백', options: { color: C.accent1 } }, { text: '이다' }], 0.5);
s.addNotes('출처: 교육부 관리 강화 방안(2026.4.29) 관련 보도, 뉴시스(2025.7.23, ISC Research 인용), 강원일보(2026.10.13).');

// ================= 4. MARKET SIZE (charts) =================
s = content('Situation', '02 · SITUATION — 시장 규모', '정부가 처음 센 숫자는 97곳, 그중 74%가 수도권에 몰려 있다', '2014년 특별점검 이후 12년간 공식 통계 공백 — 출처마다 집계 범위가 다르다');
const chartBase = {
  fontFace: F, catAxisLabelFontFace: '+mn-lt', valAxisLabelFontFace: '+mn-lt', dataLabelFontFace: '+mn-lt', titleFontFace: '+mn-lt',
  catAxisLabelColor: H.text, valAxisLabelColor: H.muted, catAxisLabelFontSize: 13, valAxisLabelFontSize: 11,
  valGridLine: { color: 'D9E2EB', size: 0.5 }, catGridLine: { style: 'none' }, showLegend: false,
  showValue: true, dataLabelFontSize: 14, dataLabelFontBold: true, dataLabelColor: H.text,
  showTitle: true, titleFontSize: 15, titleColor: H.deep, valAxisMinVal: 0,
};
R(s, { x: MX, y: 2.1, w: 5.9, h: 4.35, fill: { color: 'FFFFFF' }, line: { color: 'D9E2EB', width: 0.75 } });
s.addChart(pres.charts.BAR, [{ name: '시설 수', labels: ['ISC Research (2023)', 'KBS 전수조사 (2026.3)', '언론 사설 (2026.4)', '교육부 점검 (2026.3)'], values: [130, 130, 120, 97] }],
  Object.assign({}, chartBase, { x: MX + 0.1, y: 2.2, w: 5.7, h: 3.7, barDir: 'bar', title: '출처별 비인가 국제학교 수 (곳)', chartColors: ['83C9F2', '83C9F2', '83C9F2', '2447C9'], valAxisMaxVal: 150, dataLabelPosition: 'outEnd', objectName: name('chart') }));
T(s, '교육부 수치는 점검으로 ‘확인’된 시설, 나머지는 추정·자체조사', { x: MX + 0.2, y: 5.9, w: 5.5, h: 0.45, fontSize: 11, color: C.accent3 });
const x2 = MX + 6.2;
R(s, { x: x2, y: 2.1, w: 5.93, h: 4.35, fill: { color: 'FFFFFF' }, line: { color: 'D9E2EB', width: 0.75 } });
s.addChart(pres.charts.BAR, [{ name: '지역', labels: ['경기', '서울', '기타 지역'], values: [39, 33, 25] }],
  Object.assign({}, chartBase, { x: x2 + 0.1, y: 2.2, w: 5.7, h: 3.7, barDir: 'col', title: '미인가 국제학교 지역 분포 (곳, n=97)', chartColors: ['2447C9', '2447C9', '83C9F2'], valAxisMaxVal: 45, dataLabelPosition: 'outEnd', objectName: name('chart') }));
T(s, '수도권 72곳(74%). 기타 25곳 = 97곳 − 경기·서울', { x: x2 + 0.2, y: 5.9, w: 5.5, h: 0.45, fontSize: 11, color: C.accent3 });
src(s, '교육부 점검(2025.12–2026.3) 관련 보도, 트루스데일리(2026.6), KBS 추적60분(2026.3.27), 뉴시스(2025.7.23, ISC Research 2023 인용)');

// ================= 5. PRICE / STRUCTURE =================
s = content('Situation', '03 · SITUATION — 가격·고객·공급자', '2년 새 상단 학비가 크게 올라, 시장은 수도권 고소득층 중심으로 프리미엄화했다', '주요 30곳 고등부 연평균 4,150만원 — 4년제 대학 평균 등록금의 약 6배');
R(s, { x: MX, y: 2.1, w: 3.9, h: 4.35, fill: { color: C.text2 } });
T(s, '고등부 연평균 학비', { x: MX + 0.35, y: 2.35, w: 3.2, h: 0.4, fontSize: 15, color: 'D6E6F5' });
T(s, [{ text: '4,150', options: { fontSize: 54, bold: true } }, { text: ' 만원', options: { fontSize: 22, bold: true } }], { x: MX + 0.35, y: 2.8, w: 3.4, h: 1.0, color: C.background1 });
T(s, '주요 30곳, 사교육걱정없는세상\n‘2026 전국 미인가 대안교육시설\n학비 실태조사’(2026.2.18)', { x: MX + 0.35, y: 3.85, w: 3.3, h: 0.95, fontSize: 12, color: '9FB6CC', valign: 'top' });
T(s, [{ text: '2024년 보도 ', options: { color: 'D6E6F5' } }, { text: '통상 2~3천만원', options: { bold: true, color: C.accent2 } }], { x: MX + 0.3, y: 5.15, w: 3.55, h: 0.45, fontSize: 13 });
T(s, '재학생 약 2.6만 명(ISC 2023 추정)', { x: MX + 0.35, y: 5.65, w: 3.3, h: 0.4, fontSize: 14, color: 'D6E6F5' });
table(s, [
  hdr(['항목', '확인된 사실', '해석']),
  [cell('고객', { bold: true }), '강남·서초·수도권 신도시 고소득층, 영어유치원 졸업 가정, 해외대학 지망·국내 입시 회피층', cell('영어유치원 → 국제학교 ‘영어 몰입 파이프라인’', { bold: true, color: H.royal })],
  [cell('공급자', { bold: true }), '어학원 등록 후 전일제 운영, 미국 온라인스쿨·해외 인증 연계, 무자격 교사·급폐업 사례', cell('진입장벽이 낮아 품질 편차가 극단적', { bold: true, color: H.royal })],
  [cell('시장 크기', { bold: true }), '정부 공식 재학생·매출 통계 없음', cell('학비 편차가 커 연 수천억 원대 범위로만 추정(판단)', { bold: true, color: H.royal })],
], { x: MX + 4.2, y: 2.1, w: CW - 4.2, colW: [1.2, 4.1, 2.83], rowH: [0.5, 1.25, 1.25, 1.25] });
src(s, '연세춘추(사걱세 조사 인용), 한국경제(2024.6.5, 2026.10.20), 뉴시스(2025.7.23)');

// ================= 6. DEMAND =================
s = content('Situation', '04 · SITUATION — 수요 구조', '수요의 뿌리는 ‘영어 몰입 + 해외진학 + 공교육 이탈’이며, 단속으로 바뀌지 않는다', '학부모는 ‘불법성’보다 ‘대안의 부재’를 더 크게 느낀다');
const flow = [['유아기', '영어유치원\n(반일제 영어학원)', 'EAF4FC', H.text], ['초등 진입', '“영어 환경을\n끊을 수 없다”', 'D3E8F8', H.text], ['초·중·고', '비인가 국제학교\n(검정고시 병행)', H.royal, H.white], ['출구', '해외대학 /\n국내 검정고시 입시', H.deep, H.white]];
const fw = CW / 4;
flow.forEach((f, i) => {
  const x = MX + i * fw;
  R(s, { x, y: 2.1, w: fw - (i < 3 ? 0.02 : 0), h: 1.1, fill: { color: f[2] } }, i < 3 ? pres.shapes.PENTAGON : pres.shapes.RECTANGLE);
  T(s, [{ text: f[0] + '\n', options: { fontSize: 11, bold: true } }, { text: f[1], options: { fontSize: 14 } }], { x: x + 0.25, y: 2.1, w: fw - 0.6, h: 1.1, color: f[3] });
});
const cols = [['수요 동인 (Pull)', ['영어 조기교육의 연장 — 초등 이후에도 영어 환경 유지', '코로나 이후 유학 수요의 국내 유입', '인가 국제학교 7~8곳, 제주·송도 편중', '조기유학·인가학교보다 저렴하다는 인식']], ['이탈 동인 (Push)', ['팬데믹 이후 공교육 불신 확대', '수능·내신 중심 국내 입시 경쟁 회피', '토론·프로젝트형 수업, 소규모 학급 선호', '연예인 자녀 재학 보도 등 사회적 신호']]];
cols.forEach((c, i) => {
  const x = MX + i * (CW / 2 + 0.15), w = CW / 2 - 0.15;
  T(s, c[0], { x, y: 3.45, w, h: 0.4, fontSize: 16, bold: true, color: C.text2 });
  T(s, c[1].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < c[1].length - 1, paraSpaceAfter: 6 } })), { x, y: 3.9, w, h: 1.75, fontSize: 15, valign: 'top' });
});
insight(s, 5.85, [{ text: '시사점 — 사라지는 것은 ‘전일제 학교 형태’이지 ‘영어 몰입 교육비 지출’이 아니다. 지출은 ' }, { text: '합법 방과후·온라인·인가학교', options: { color: C.accent1 } }, { text: '로 재배치된다' }], 0.62);
src(s, '한국경제(2026.10.20), 사설(2026.4.30), 한국경제(2024.6.5)');

// ================= 7. RISKS =================
pres.addSection({ title: 'Complication' });
s = content('Complication', '05 · COMPLICATION — 문제 구조', '학부모가 떠안는 위험은 ‘학력 불인정’보다 ‘보호 장치 전무’와 ‘갑작스러운 폐업’이다', '고액 학비를 선납하지만 품질·안전·환불을 보증할 제도적 장치가 없다');
const sev = (t) => cell(t, { align: 'center', bold: true, color: t === '상' ? H.red : H.amber });
table(s, [
  hdr(['리스크', '확인된 사례·사실', '심각도'], [0, 0, 'center']),
  [cell('허위·과장 홍보', { bold: true }), '‘미국 졸업장’, ‘해외대학 MOU’, ‘아이비리그 합격’ 홍보 검증 보도(KBS 2026.3)', sev('상')],
  [cell('폐업·환불 분쟁', { bold: true }), '폐원 후 수업 강행·환불 거부, 인천 시설 퇴거 요구에도 수업 지속(2026.10)', sev('상')],
  [cell('학생 보호 공백', { bold: true }), '학교폭력 발생 시 공적 절차(학폭위 등) 보호가 어려움', sev('상')],
  [cell('학력 불인정', { bold: true }), '국내 학력 불인정 → 검정고시로 별도 입증', sev('중')],
  [cell('교원·시설', { bold: true }), '무자격 교사·부실 교육, 학비 3,000만원에도 체육관 미사용 사례', sev('중')],
], { y: 2.1, w: 7.4, colW: [1.9, 4.5, 1.0], rowH: 0.6, fontSize: 12.5 });
const px = MX + 7.7, pw = CW - 7.7;
R(s, { x: px, y: 2.1, w: pw, h: 0.62, fill: { color: C.text2 } });
T(s, '상위 문제: 고액·무보호 시장이 제도 밖에서 성장', { x: px, y: 2.1, w: pw, h: 0.62, fontSize: 14, bold: true, color: C.background1, align: 'center' });
const roots = [['① 공급 경로 부재', '대안교육기관법이 외국어·해외진학 목적 시설을 등록에서 배제'], ['② 집행 수단 부족', '폐쇄명령 후 강제 중단 권한·이행강제금 없음 → 명령 0건'], ['③ 정보 비대칭', '‘국제학교’ 명칭 제한 규정 부재, 해외 인증 판별 어려움']];
roots.forEach((r, i) => {
  const y = 2.9 + i * 0.96;
  R(s, { x: px, y, w: pw, h: 0.84, fill: { color: C.background2 } });
  T(s, [{ text: r[0] + '\n', options: { bold: true, color: C.text2, fontSize: 14 } }, { text: r[1], options: { fontSize: 12 } }], { x: px + 0.2, y, w: pw - 0.4, h: 0.84 });
});
insight(s, 6.05, [{ text: '시사점 — 정부가 단기에 손댈 수 있는 것은 ②집행·③명칭이며, ①합법 공급은 국회 입법에 달린 ' }, { text: '중기 과제', options: { color: C.accent1 } }], 0.55);
src(s, 'KBS 추적60분(2026.3.27), 한국경제(2026.10.20 / 10.30), 연세춘추, 교육부 관리 강화 방안(2026.4.29) 보도');

// ================= 8. TIMELINE =================
pres.addSection({ title: 'Evidence' });
s = content('Evidence', '06 · EVIDENCE — 정책 흐름', '2026년은 ‘방치’에서 ‘집중 단속’으로 정책 기조가 바뀐 원년이다', '점검 → 방안 발표 → 1·2차 고지 → 고발 순으로 진행 중이며, 11월 첫 고발이 시험대');
const tl = [
  ['2014', '교육부 특별점검\n이후 12년 공백'], ['2025.12–26.3', '213곳 점검\n국제학교형 97곳', 1], ['2026.4.29', '관리 강화 방안\n학교형 폐쇄 방침', 1], ['2026.6', '광주 원장 송치\n인천 1차 고지'],
  ['2026.8', '‘오후반·온라인’\n우회 등장'], ['2026.10.22', '강원 외국교육기관\n전역 허용'], ['2026.10.23', '국감: 11월 고발,\n이행강제금 검토', 1], ['2026.11~', '첫 고발 집행\n= 의지의 시험대'],
];
const ly = 3.75, step = CW / tl.length;
R(s, { x: MX, y: ly - 0.015, w: CW, h: 0.03, fill: { color: 'D9E2EB' } });
tl.forEach((t, i) => {
  const cx = MX + step * i + step / 2, key = t[2];
  const d = key ? 0.34 : 0.24;
  R(s, { x: cx - d / 2, y: ly - d / 2, w: d, h: d, fill: { color: key ? C.text2 : C.accent1 } }, pres.shapes.OVAL);
  const up = i % 2 === 0;
  T(s, t[0], { x: cx - step / 2 + 0.03, y: up ? 2.2 : 4.05, w: step - 0.06, h: 0.38, fontSize: 13, bold: true, color: C.accent1, align: 'center' });
  T(s, t[1], { x: cx - step / 2 + 0.03, y: up ? 2.58 : 4.43, w: step - 0.06, h: 0.9, fontSize: 12.5, align: 'center', valign: 'top', bold: !!key, color: key ? C.text2 : C.text1 });
});
insight(s, 5.65, [{ text: '관전 포인트 — 인천·서울의 첫 고발과 2027년 이행강제금·과징금 입법이 ' }, { text: '시장 재편 속도', options: { color: C.accent1 } }, { text: '를 결정한다' }], 0.62);
src(s, '뉴스핌(2026.4.29, 10.23), 한국경제(2026.4.29, 10.20, 10.30), 아시아경제(2026.6.15), 머니투데이(2026.8.26), 강원일보(2026.10.13)');

// ================= 9. ENFORCEMENT LIMITS =================
s = content('Evidence', '07 · EVIDENCE — 집행의 한계', '단속 의지는 강하지만 수단은 약하다 — 시장은 이미 ‘우회 모델’로 움직이고 있다', '“폐쇄명령을 내려도 수업을 계속하면 강제로 막을 권한이 없다” — 교육부 관계자');
const st = (t, col) => cell(t, { align: 'center', bold: true, color: col });
table(s, [
  hdr(['정부 수단', '현재', '입법 후(예상)'], [0, 'center', 'center']),
  ['시정 고지(1·2차)', st('시행', H.green), cell('유지', { align: 'center' })],
  ['고발·수사의뢰', st('11월 첫 예정', H.amber), cell('확대', { align: 'center' })],
  ['폐쇄명령', st('0건', H.red), cell('이행강제금 결합', { align: 'center' })],
  ['이행강제금·과징금', st('없음', H.red), cell('신설 추진', { align: 'center' })],
  ['위반사항 공표', st('없음', H.red), cell('신설 추진', { align: 'center' })],
  ['‘국제학교’ 명칭 제한', st('규정 없음', H.red), cell('논의 단계', { align: 'center' })],
], { y: 2.1, w: 5.6, colW: [2.2, 1.6, 1.8], rowH: 0.56 });
const by = [['① 오후반 전환', '정규 수업 후 영어 몰입. “학교 형태면 제재”', '회색지대', H.amber], ['② 온라인스쿨 + 학습관리', '해외 온라인스쿨이 졸업장 발급. 교육부 “추후 검토”', '규제 공백', H.red], ['③ 대안교육기관 등록', '담양 사례처럼 목적 변경 후 신청. 법 §5② 해석 관건', '특혜 논란', H.amber], ['④ 인가 국제학교 이동', '단속 이후 제주 등 인가학교 관심 증가', '합법 흡수', H.green]];
const bx = MX + 5.9, bw = (CW - 5.9 - 0.25) / 2;
by.forEach((b, i) => {
  const x = bx + (i % 2) * (bw + 0.25), y = 2.1 + Math.floor(i / 2) * 1.98;
  R(s, { x, y, w: bw, h: 1.78, fill: { color: 'F4F6F8' } });
  T(s, b[0], { x: x + 0.2, y: y + 0.15, w: bw - 0.4, h: 0.4, fontSize: 15, bold: true, color: C.text2 });
  T(s, b[1], { x: x + 0.2, y: y + 0.58, w: bw - 0.4, h: 0.7, fontSize: 13, valign: 'top' });
  R(s, { x: x + 0.2, y: y + 1.3, w: 1.3, h: 0.34, fill: { color: b[3] }, rectRadius: 0.17 }, pres.shapes.ROUNDED_RECTANGLE);
  T(s, b[2], { x: x + 0.2, y: y + 1.3, w: 1.3, h: 0.34, fontSize: 11, bold: true, color: C.background1, align: 'center' });
});
src(s, '한국경제(2026.10.20), 머니투데이(2026.8.26), 경향신문(2026.5.28), 2026.6 제주 인가 국제학교 관련 보도');

// ================= 10. SCENARIOS =================
pres.addSection({ title: 'Future' });
s = content('Future', '08 · FUTURE — 2027–2030 시나리오', '가장 가능성 높은 미래는 ‘학교형 축소 + 수요의 합법 채널 이동’이다', '두 변수 — 집행력 입법 통과 여부, 인가형 국제학교 공급 확대 속도 — 가 시나리오를 가른다');
const sc = [
  ['A. 지하화·변형', '25%', '입법 지연, 집행 미온', '‘학원’ 간판 아래 사실상 학교 운영 지속. 오후반·온라인 혼합으로 회피, 피해 사고 반복', '시설 소폭 감소, 리스크는 학부모에게 전가'],
  ['B. 축소·재배치', '55%', '이행강제금 등 입법(2027), 인가형 점진 확대', '전일제 학교형은 절반 이하로 축소. 수요는 인가학교·합법 방과후·온라인스쿨로 분산', '‘인가 프리미엄’ ↔ ‘합법 보완교육’ 양극화', true],
  ['C. 제도권 흡수', '20%', '대안교육기관법 개정 또는 국제학교 특례 확대', '상위 시설은 등록형 국제 대안학교·특구 국제학교로 양성화, 하위 시설 퇴출', '자본력 있는 대형 사업자 중심 재편'],
];
const sw = (CW - 0.6) / 3;
sc.forEach((c, i) => {
  const x = MX + i * (sw + 0.3), base = c[5];
  R(s, { x, y: 2.1, w: sw, h: 4.2, fill: { color: base ? C.background2 : 'FFFFFF' }, line: { color: base ? H.royal : H.div, width: base ? 2 : 0.75 } });
  T(s, c[0], { x: x + 0.25, y: 2.25, w: sw - 1.6, h: 0.55, fontSize: 17, bold: true, color: C.text2 });
  if (base) { R(s, { x: x + 0.25, y: 1.95, w: 1.3, h: 0.3, fill: { color: H.royal }, rectRadius: 0.15 }, pres.shapes.ROUNDED_RECTANGLE); T(s, 'BASE CASE', { x: x + 0.25, y: 1.95, w: 1.3, h: 0.3, fontSize: 11, bold: true, color: C.background1, align: 'center' }); }
  T(s, c[1], { x: x + sw - 1.45, y: 2.2, w: 1.2, h: 0.65, fontSize: 32, bold: true, color: C.accent1, align: 'right' });
  T(s, [{ text: '조건  ', options: { bold: true, color: C.accent1 } }, { text: c[2] }], { x: x + 0.25, y: 3.0, w: sw - 0.5, h: 0.7, fontSize: 13.5, valign: 'top' });
  T(s, c[3], { x: x + 0.25, y: 3.75, w: sw - 0.5, h: 1.4, fontSize: 14, valign: 'top' });
  T(s, [{ text: '시장  ', options: { bold: true, color: C.accent1 } }, { text: c[4] }], { x: x + 0.25, y: 5.25, w: sw - 0.5, h: 0.85, fontSize: 13.5, valign: 'top', bold: false });
});
T(s, '※ 확률은 공개 정보에 근거한 분석자의 판단이며 정량 모델 결과가 아니다.', { x: MX, y: 6.45, w: CW, h: 0.35, fontSize: 11, color: C.accent3 });

// ================= 11. WATCH LIST =================
s = content('Future', '09 · FUTURE — 선행 지표', '시나리오는 향후 6~12개월의 다섯 가지 신호로 판가름 난다', '경영기획이 월 1회 점검하고, 분기마다 A/B/C 확률을 갱신한다');
table(s, [
  hdr(['선행 지표', '시점', '신호 해석']),
  [cell('인천·서울 첫 고발 및 수사 결과', { bold: true }), '2026.11 ~ 2027 상반기', '기소·유죄 시 B 가속, 불기소·지연 시 A 강화'],
  [cell('이행강제금·과징금·공표제 법안', { bold: true }), '2027 정기국회 전후', '통과 시 B 확정적, 전일제 모델 수익성 급락'],
  [cell('온라인스쿨 연계 모델 유권해석', { bold: true }), '2027', '허용 시 ‘온라인 + 학습관리’가 주력 대체재로 부상'],
  [cell('강원특별법 국제학교 특례(4차 개정안)', { bold: true }), '국회 계류 중', '통과 시 C 확률 상승, 영리 국제학교 신규 공급'],
  [cell('‘국제학교’ 명칭 사용 제한 입법', { bold: true }), '논의 단계', '통과 시 ‘○○ 아카데미’ 등 브랜드 전환 대거 발생'],
], { y: 2.1, w: CW, colW: [4.3, 2.8, 5.03], rowH: 0.6, fontSize: 14 });
insight(s, 6.2, [{ text: '판단 — 2027년 상반기까지는 ' }, { text: '‘오후반·온라인’ 혼합형', options: { color: C.accent1 } }, { text: '으로 빠르게 재편되고, 입법 통과 시 전일제 모델은 급격히 위축된다' }], 0.55);

// ================= 12. STAKEHOLDERS / MAP =================
pres.addSection({ title: 'Implication' });
s = content('Implication', '10 · IMPLICATION — 수요 이동 지도', '단속의 최대 수혜자는 인가 국제학교와 ‘합법 형태의 영어 몰입 보완교육’이다', '합법성 × 영어 몰입도 — 기본 시나리오(B)에서 예상되는 수요 이동 방향(판단)');
const gx = MX + 0.5, gy = 2.15, gw = 7.0, gh = 4.2;
const q = [[0, 0, 'F4F6F8', '합법 · 부분 몰입', H.muted], [1, 0, H.lblue, '합법 · 고몰입 (성장 영역)', H.royal], [0, 1, 'FBECEC', '불법 · 부분 몰입', H.red], [1, 1, 'FDF3E3', '불법 · 고몰입 (단속 핵심)', H.red]];
q.forEach((qq) => {
  R(s, { x: gx + qq[0] * gw / 2, y: gy + qq[1] * gh / 2, w: gw / 2, h: gh / 2, fill: { color: qq[2] } });
  T(s, qq[3], { x: gx + qq[0] * gw / 2 + 0.15, y: gy + qq[1] * gh / 2 + (qq[1] ? gh / 2 - 0.45 : 0.1), w: gw / 2 - 0.3, h: 0.35, fontSize: 12, bold: true, color: qq[4] });
});
T(s, '영어 몰입도 →', { x: gx, y: gy + gh + 0.02, w: gw, h: 0.3, fontSize: 12, color: C.accent3, align: 'center' });
T(s, '합법성 →', { x: MX - 0.55, y: gy + gh / 2 - 0.15, w: 1.6, h: 0.3, fontSize: 12, color: C.accent3, align: 'center', rotate: 270 });
const bub = (cx, cy, r, fill, txt, col) => { R(s, { x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, fill: { color: fill } }, pres.shapes.OVAL); T(s, txt, { x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, fontSize: 11, bold: true, align: 'center', color: col || H.text }); };
bub(gx + 1.2, gy + 1.05, 0.5, H.sky, '일반\n어학원');
bub(gx + 2.7, gy + 1.15, 0.55, H.sky, '방과후\n영어 몰입');
bub(gx + 4.6, gy + 1.15, 0.62, H.royal, '인가\n국제학교', H.white);
bub(gx + 6.25, gy + 1.05, 0.55, H.sky, '온라인스쿨\n+학습관리');
bub(gx + 2.6, gy + 3.05, 0.5, 'E8A0A0', '오후반\n(학교형)');
bub(gx + 5.3, gy + 3.0, 0.72, H.red, '전일제\n비인가', H.white);
const arrow = (x1, y1, x2, y2, dash) => s.addShape(pres.shapes.LINE, { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1), flipH: x2 < x1, flipV: y2 < y1, line: { color: H.deep, width: 2, endArrowType: 'triangle', dashType: dash ? 'dash' : 'solid' }, objectName: name('arrow') });
arrow(gx + 5.05, gy + 2.3, gx + 4.8, gy + 1.8);
arrow(gx + 5.7, gy + 2.4, gx + 6.05, gy + 1.62);
arrow(gx + 4.6, gy + 2.85, gx + 3.25, gy + 1.45, true);
T(s, '원 크기는 개념적 상대 규모이며 실측치가 아니다', { x: gx, y: 6.62, w: gw, h: 0.3, fontSize: 10, color: C.accent3 });
const sx = MX + 8.0, sww = CW - 8.0;
const stake = [['인가 국제학교·외국교육기관', '긍정', H.green, '안정성 프리미엄, 대기 수요 증가'], ['합법 어학원·방과후 몰입', '긍정', H.green, '전일제 이탈 수요 흡수'], ['해외 온라인스쿨', '긍정', H.green, '단기 수혜, 유권해석 리스크'], ['유아 영어학원', '주의', H.amber, '상위 출구 축소, 규제 여론 동반'], ['재학생·학부모', '혼재', H.amber, '폐업·환불 리스크'], ['비인가 운영자', '부정', H.red, '형사·행정 리스크, 전환 강요']];
stake.forEach((k, i) => {
  const y = 2.15 + i * 0.72;
  R(s, { x: sx, y, w: sww, h: 0.62, fill: { color: 'F4F6F8' } });
  R(s, { x: sx + 0.12, y: y + 0.15, w: 0.7, h: 0.32, fill: { color: k[2] }, rectRadius: 0.16 }, pres.shapes.ROUNDED_RECTANGLE);
  T(s, k[1], { x: sx + 0.12, y: y + 0.15, w: 0.7, h: 0.32, fontSize: 11, bold: true, color: C.background1, align: 'center' });
  T(s, [{ text: k[0] + '\n', options: { bold: true, fontSize: 13, color: C.text2 } }, { text: k[3], options: { fontSize: 11.5 } }], { x: sx + 0.95, y, w: sww - 1.05, h: 0.62 });
});

// ================= 13. OPTIONS + RECO =================
pres.addSection({ title: 'Recommendation' });
s = content('Recommendation', '11 · STRATEGIC OPTIONS — 시사컴퍼니', '시사컴퍼니는 ‘학교’가 아닌 ‘합법 보완교육’에서 이탈 수요를 흡수할 수 있다', 'Language Design(럭스어학원·시사어학원·출판) 자산과 연결한 3개 대안 비교');
const rk = (t, col) => cell(t, { align: 'center', bold: true, color: col });
table(s, [
  hdr(['대안', '내용', '효과', '비용', '규제 위험', '평가'], [0, 0, 'center', 'center', 'center', 'center']),
  [cell('A. 관망·방어', { bold: true }), '럭스어학원 컴플라이언스 점검만 실시, 신규 사업 없음', cell('낮음', { align: 'center' }), cell('낮음', { align: 'center' }), rk('낮음', H.green), cell('보조', { align: 'center' })],
  [cell('B. 합법 보완교육 확장', { bold: true, fill: { color: H.lblue } }), cell('럭스어학원 졸업생·초등 대상 방과후 영어 몰입 + 시사어학원 일·중 제2외국어(IB·국제학교생). 1~2개 지점 파일럿', { fill: { color: H.lblue } }), cell('중~상', { align: 'center', fill: { color: H.lblue } }), cell('중', { align: 'center', fill: { color: H.lblue } }), rk('낮음', H.green), cell('추천', { align: 'center', bold: true, color: H.royal, fill: { color: H.lblue } })],
  [cell('C. 국제학교형 진출', { bold: true }), '비인가 시설 인수 또는 전일제 영어 과정 개설', cell('상(단기)', { align: 'center' }), cell('상', { align: 'center' }), rk('매우 높음', H.red), cell('배제', { align: 'center' })],
], { y: 2.1, w: CW, colW: [2.5, 5.73, 1.0, 0.9, 1.1, 0.9], rowH: [0.5, 0.62, 0.9, 0.62] });
R(s, { x: MX, y: 4.95, w: CW, h: 1.15, fill: { color: C.text2 } });
R(s, { x: MX + 0.3, y: 5.25, w: 2.4, h: 0.55, fill: { color: C.accent2 } });
T(s, 'CONDITIONAL GO', { x: MX + 0.3, y: 5.25, w: 2.4, h: 0.55, fontSize: 16, bold: true, color: C.text2, align: 'center' });
T(s, '대안 B — 학교 형태가 아닌 합법 보완교육으로 파일럿 추진. 조건: 법무 검토 통과 + 럭스어학원 컴플라이언스 점검 완료', { x: MX + 3.0, y: 4.95, w: CW - 3.3, h: 1.15, fontSize: 16, bold: true, color: C.background1 });
T(s, [{ text: '왜 B인가  ', options: { bold: true, color: C.accent1 } }, { text: '① 단속은 ‘형태’를 겨냥하므로 형태 리스크 없는 보완교육이 자연스러운 수용처  ② 기존 원생·콘텐츠·강사 활용으로 추가 CAPEX 작음  ③ C는 2027년 입법 시 존속 위협' }], { x: MX, y: 6.25, w: CW, h: 0.6, fontSize: 13.5, valign: 'top' });
s.addNotes('효과·비용은 정성 평가. 사업부 내부 실적(매출·원생수·유지율) 확보 시 정량 검토로 대체해야 함. 본 덱은 내부 수치를 가정하지 않음.');

// ================= 14. ROADMAP =================
s = content('Recommendation', '12 · ACTION ROADMAP', '180일 안에 점검·파일럿·확장 판단을 마치고 입법 전 포지션을 확정한다', '담당·기한·KPI는 제안안이며 대표이사 결재 후 확정');
const rm = [
  ['30일', '~2026.12', '컴플라이언스 점검', '럭스어학원 운영·홍보물 점검(‘학교·졸업·학년’ 표현, 교습시간·교습비 준수). 외부 법무 자문', 'Language Design 본부 + 법무', '체크리스트 100% 완료, 위반 소지 0건'],
  ['90일', '~2027.2', '파일럿 설계·모집', '방과후 영어 몰입·제2외국어 과정 설계, 1~2개 지점 모집. 첫 고발 결과 월간 모니터링', '럭스어학원 · 시사어학원', '모집 달성률, 상담→등록 전환율(목표는 내부 실적 기준)'],
  ['180일', '~2027.5', '평가·확장 결정', '파일럿 평가 후 확장/중단 결정. 입법 상황 반영, 인가 국제학교 제휴(제2외국어 위탁 등) 타진', '경영기획', '3개월 유지율, 공헌이익, 제휴 MOU 건수'],
];
const rw = (CW - 0.6) / 3;
rm.forEach((r, i) => {
  const x = MX + i * (rw + 0.3);
  R(s, { x, y: 2.1, w: rw, h: 0.95, fill: { color: i === 2 ? C.text2 : C.accent1 } });
  T(s, [{ text: r[0], options: { fontSize: 24, bold: true } }, { text: '   ' + r[1], options: { fontSize: 13 } }], { x: x + 0.25, y: 2.1, w: rw - 0.5, h: 0.55, color: C.background1 });
  T(s, r[2], { x: x + 0.25, y: 2.6, w: rw - 0.5, h: 0.38, fontSize: 15, bold: true, color: C.background1 });
  R(s, { x, y: 3.05, w: rw, h: 3.4, fill: { color: 'F4F6F8' } });
  T(s, r[3], { x: x + 0.25, y: 3.2, w: rw - 0.5, h: 1.45, fontSize: 14, valign: 'top' });
  T(s, [{ text: '담당(안)\n', options: { bold: true, color: C.accent1, fontSize: 12 } }, { text: r[4], options: { fontSize: 13.5 } }], { x: x + 0.25, y: 4.7, w: rw - 0.5, h: 0.7, valign: 'top' });
  T(s, [{ text: 'KPI · 게이트\n', options: { bold: true, color: C.accent1, fontSize: 12 } }, { text: r[5], options: { fontSize: 13.5 } }], { x: x + 0.25, y: 5.45, w: rw - 0.5, h: 0.9, valign: 'top' });
});
T(s, '거버넌스: 경영기획이 월 1회 규제 동향 브리핑, 분기마다 시나리오 확률 갱신. 이행강제금 입법·온라인스쿨 유권해석 발표 시 즉시 재검토', { x: MX, y: 6.55, w: CW, h: 0.4, fontSize: 12, color: C.accent3 });

// ================= 15. CEO DECISIONS =================
s = content('Recommendation', '13 · CEO DECISIONS', '대표이사께 세 가지 결정을 요청드립니다', '1·2는 즉시 결정, 3은 1·2 완료를 조건으로 한 승인');
const dc = [
  ['1', '‘국제학교형(전일제) 사업 미진출’ 원칙 확정', '인수·제휴·과정 개설 모두 배제. 내부 지침으로 명문화', '예산: 없음', '기한: 2026.11 내'],
  ['2', '럭스어학원 컴플라이언스 점검 착수 승인', '외부 법무 자문 포함, 결과를 30일 내 보고', '예산: 법무 자문비(견적 후 확정)', '기한: 2026.12'],
  ['3', '방과후 영어 몰입·제2외국어 파일럿 승인 (Conditional)', '결정 1·2 완료 전제, 1~2개 지점 시행. 180일 후 확장 여부 재결재', '예산: 파일럿 운영비(사업부 산정)', '기한: 2027.2 개시'],
];
dc.forEach((d, i) => {
  const y = 2.15 + i * 1.45;
  R(s, { x: MX, y, w: CW, h: 1.25, fill: { color: i === 2 ? C.background2 : 'F4F6F8' } });
  R(s, { x: MX + 0.3, y: y + 0.27, w: 0.72, h: 0.72, fill: { color: C.accent1 } }, pres.shapes.OVAL);
  T(s, d[0], { x: MX + 0.3, y: y + 0.27, w: 0.72, h: 0.72, fontSize: 24, bold: true, color: C.background1, align: 'center' });
  T(s, [{ text: d[1] + '\n', options: { fontSize: 18, bold: true, color: C.text2 } }, { text: d[2], options: { fontSize: 14 } }], { x: MX + 1.3, y, w: 7.4, h: 1.25 });
  R(s, { x: MX + 8.95, y: y + 0.15, w: 3.0, h: 0.95, fill: { color: 'FFFFFF' } });
  T(s, d[3] + '\n' + d[4], { x: MX + 9.1, y: y + 0.15, w: 2.75, h: 0.95, fontSize: 12.5, color: C.accent3 });
});

// ================= 16. SOURCES =================
pres.addSection({ title: 'Appendix' });
s = content('Appendix', 'APPENDIX — 출처 및 유의사항', '주요 출처와 유의사항', '2026.11.8 기준 웹 검색으로 확인한 공개 보도. 원문 페이지는 직접 대조하지 못했으므로 대외 인용 전 확인 권장');
const srcs = [
  '뉴스핌, 교육부 미인가·미등록 교육시설 집중 단속, 2026.4.29', '한국경제, 정부 미인가 국제학교에 칼 뺐다, 2026.4.29', '머니투데이, 교육부 “미인가 국제학교 전부 폐쇄 대상”, 2026.4.29',
  '머니투데이, 단속 피한 ‘오후반’ 등장, 2026.8.26', '뉴스핌, [2026 국감] 11월 고발·수사 의뢰 예정, 2026.10.23', '한국경제, 폐원 후 배짱 수업·환불 거부, 2026.10.20',
  '한국경제, 인천 미인가 국제학교 퇴거 요구에도 버티기, 2026.10.30', '트루스데일리, 미인가 국제학교 97곳 ‘등록 막막’, 2026', '연세춘추, 1년 학비 4천만 원 비인가 국제학교(사걱세 조사 인용), 2026',
  'KBS 추적60분 관련 보도, 2026.3.27', '뉴시스, “학원도 학교도 아니었다”(ISC Research 인용), 2025.7.23', '경향신문, 담양 미인가 국제학교 ‘편법 등록’ 논란, 2026.5.28',
  'Asia Business Daily, 6년 운영 원장 송치, 2026.6.15', '강원일보, 강원 전역 외국교육기관 설립 가능, 2026.10.13', '한국경제, 정식 국제학교보다 싼 수업료, 2024.6.5',
];
const half = Math.ceil(srcs.length / 2);
[srcs.slice(0, half), srcs.slice(half)].forEach((col, ci) => {
  T(s, col.map((t, j) => ({ text: t, options: { bullet: { type: 'number', startAt: ci * half + 1 }, breakLine: j < col.length - 1, paraSpaceAfter: 4 } })), { x: MX + ci * (CW / 2 + 0.1), y: 2.1, w: CW / 2 - 0.1, h: 3.1, fontSize: 12, valign: 'top' });
});
R(s, { x: MX, y: 5.35, w: CW, h: 1.5, fill: { color: 'F4F6F8' } });
T(s, ['시설 수는 집계 주체·범위에 따라 97~130곳으로 다르며, 정부 확인치(97곳)를 기준으로 했다', '재학생 2.6만 명은 ISC Research(2023) 추정을 언론이 인용한 값으로 정부 공식 통계가 아니다', '시나리오 확률·이해관계자 영향·대안 평가는 분석자 판단이며, 내부 실적 수치는 사용하지 않았다', '법령 해석은 보도 기준 요약이며, 사업 실행 전 법무 검토가 필요하다'].map((t, j, a) => ({ text: t, options: { bullet: true, breakLine: j < a.length - 1, paraSpaceAfter: 3 } })), { x: MX + 0.25, y: 5.42, w: CW - 0.5, h: 1.36, fontSize: 12.5, valign: 'middle' });

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log('ok');
})();
