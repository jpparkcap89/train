/**
 * '연수 관리' 시트 ↔ 연수 안내·신청 사이트 (누적 관리형)
 *
 * 구조: 연수마다 '연수 ID'를 붙여 모든 탭에 쌓습니다. 첫 화면에는 Settings의 currentCourse(대표 연수)가 보입니다.
 *   Settings  사이트 공통 설정 (관리자 비밀번호, 대표 연수 등)
 *   연수      연수 1개 = 1줄 (상태: 준비 / 모집중 / 마감 / 종료)
 *   회차      연수별 회차
 *   세부과정  연수별 주제 탭 내용
 *   특징      특징 카드 (연수 ID를 비우면 모든 연수 공통)
 *   FAQ       자주 묻는 질문 (연수 ID를 비우면 모든 연수 공통)
 *   후기      연수별 후기 (노출=예인 것만 사이트에 표시)
 *   신청      신청 1건 = 1줄 (연수 ID로 구분, 상태·이수 기록)
 *   참가자    학교 메일 기준 누적 집계 (자동으로 다시 계산되니 직접 고치지 마세요)
 *
 * 처음 한 번: 이 코드를 [확장 프로그램 > Apps Script]에 붙여넣고 setup을 실행한 뒤 [배포 > 새 배포 > 웹 앱].
 * 열 제목(1행)은 바꾸지 마세요. 열 순서는 바꿔도 됩니다.
 */

var T = {
  courses: { name: '연수', cols: [['id', '연수 ID'], ['status', '상태'], ['name', '연수명'], ['eyebrow', '라벨'], ['title1', '제목 첫 줄'], ['title2', '제목 둘째 줄'], ['lead', '소개'], ['period', '기간'], ['time', '시간'], ['credit', '이수'], ['place', '장소'], ['capacity', '정원'], ['openAt', '신청 시작일'], ['closeAt', '신청 마감일'], ['applyText', '신청 안내'], ['courseLead', '과정 설명'], ['voicesLead', '후기 설명'], ['completion', '이수 기준'], ['createdAt', '만든 날']],
             widths: [110, 80, 220, 220, 220, 200, 360, 200, 170, 150, 200, 60, 100, 100, 320, 280, 280, 220, 100] },
  sessions: { name: '회차', cols: [['course', '연수 ID'], ['week', '회차'], ['date', '날짜'], ['title', '제목'], ['desc', '설명'], ['time', '시간']], widths: [110, 70, 100, 260, 300, 80] },
  topics: { name: '세부과정', cols: [['course', '연수 ID'], ['topic', '주제 (탭 이름)'], ['title', '제목'], ['meta', '부가 정보'], ['time', '시간']], widths: [110, 120, 240, 160, 80] },
  features: { name: '특징', cols: [['course', '연수 ID (비우면 공통)'], ['icon', '아이콘'], ['title', '제목'], ['text', '설명 (**굵게** 가능)']], widths: [150, 70, 180, 440] },
  faq: { name: 'FAQ', cols: [['course', '연수 ID (비우면 공통)'], ['q', '질문'], ['a', '답변'], ['linkText', '링크 문구'], ['target', '이동 위치 (course / apply)']], widths: [150, 280, 420, 140, 160] },
  reviews: { name: '후기', cols: [['course', '연수 ID'], ['show', '노출'], ['quote', '후기 (*강조할 말*)'], ['name', '이름'], ['role', '소속']], widths: [110, 60, 420, 140, 160] },
  apps: { name: '신청', cols: [['appId', '신청 ID'], ['course', '연수 ID'], ['at', '신청일시'], ['name', '이름'], ['email', '학교 메일'], ['school', '소속 학교'], ['phone', '전화번호'], ['status', '상태'], ['done', '이수'], ['memo', '메모']], widths: [130, 110, 140, 90, 220, 150, 130, 70, 70, 240] },
  people: { name: '참가자', cols: [['email', '학교 메일'], ['name', '이름'], ['school', '소속 학교'], ['phone', '전화번호'], ['applied', '신청 횟수'], ['confirmed', '확정 횟수'], ['completed', '이수 횟수'], ['lastCourse', '최근 연수'], ['lastAt', '최근 신청일']], widths: [220, 90, 150, 130, 80, 80, 80, 110, 140] }
};
var SETTINGS_TAB = 'Settings';
var COURSE_STATUS = ['준비', '모집중', '마감', '종료'];
var APP_STATUS = ['접수', '확정', '대기', '취소'];
var DONE = ['', '이수', '미이수'];
var LISTS = ['sessions', 'topics', 'features', 'faq', 'reviews'];
var COMMON_OK = ['features', 'faq'];             // 연수 ID를 비우면 모든 연수 공통
var TZ = 'Asia/Seoul';

/* ============ 처음 설정 ============ */

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var id = '2027-W01';
  var pw = Math.random().toString(36).slice(2, 10);

  if (!ss.getSheetByName(SETTINGS_TAB)) {
    var st = ss.insertSheet(SETTINGS_TAB, 0);
    var rows = [['키', '값', '설명'],
      ['siteTitle', 'JP Park 연수실', '브라우저 탭 제목 (연수명이 앞에 붙습니다)'],
      ['brandName', 'JP Park 연수실', '왼쪽 위 로고 옆 이름'],
      ['brandInitials', 'JP', '로고 상자 안 글자 (2자 권장)'],
      ['adminPassword', pw, '관리자 로그인 비밀번호. 이 칸만 고치면 바뀝니다'],
      ['currentCourse', id, '첫 화면에 보일 대표 연수 ID (연수 탭의 연수 ID)'],
      ['notifyEmail', '', '신청이 들어오면 알림 메일을 받을 주소 (비워 두면 보내지 않음)'],
      ['contactPlace', '율포초등학교 수석교사실', '문의처'],
      ['contactPhone', '031-000-0000', '문의 전화'],
      ['footerText', '© 2026 박종필 · AI와 디지털로 더 나은 배움을 설계합니다.', '맨 아래 문구']];
    st.getRange(1, 1, rows.length, 3).setValues(rows);
    styleHeader_(st, 3, [160, 420, 380]);
  }

  makeTable_('courses', [{
    id: id, status: '모집중', name: '2027 겨울방학 AI 수업 설계 연수', eyebrow: '✨ 2027 겨울방학 교원 자율연수',
    title1: 'AI로 수업을 다시 설계하는', title2: '4주간의 실습 연수',
    lead: '생성형 AI를 **우리 반 교육과정에 자연스럽게 연결하는 방법**을, 매주 직접 만든 수업안으로 익힙니다. 이론보다 실습, 도구보다 **수업 설계**에 집중합니다.',
    period: '1월 6일 – 1월 27일 · 매주 수', time: '16:00 – 19:00 (주 3시간)', credit: '15시간 · 1학점 인정', place: '율포초 미래교실 + 온라인 병행',
    capacity: 30, openAt: '', closeAt: '2026-12-20',
    applyText: '12월 20일까지 선착순으로 받습니다. 신청 후 3일 안에 학교 메일로 확정 안내를 보내 드립니다.',
    courseLead: '주제별로 나눠 보면 어떤 회차가 내 수업에 맞는지 바로 고를 수 있어요. 오른쪽 탭을 눌러 보세요.',
    voicesLead: '지난 연수에 참여한 선생님들의 후기와 자주 받은 질문을 모았습니다.',
    completion: '4회 중 3회 이상 출석', createdAt: today_()
  }]);
  makeTable_('sessions', [
    [id, '1주', '2027-01-06', 'AI 리터러시와 수업 윤리', '학생과 함께 정하는 AI 사용 규칙', '3시간'],
    [id, '2주', '2027-01-13', '성취기준으로 시작하는 AI 수업', '백워드 설계에 AI 활동 끼워 넣기', '3시간'],
    [id, '3주', '2027-01-20', '피드백과 평가 도우미', '루브릭 기반 피드백 자동화', '3시간'],
    [id, '4주', '2027-01-27', '수업 나눔과 공개', '완성한 수업안 발표와 상호 피드백', '3시간']]);
  makeTable_('topics', [
    [id, 'AI 기초', '생성형 AI 작동 원리', '1주차 · 강의', '40분'], [id, 'AI 기초', '프롬프트 작성 실습', '1주차 · 실습', '80분'],
    [id, 'AI 기초', 'AI 사용 규칙 만들기', '1주차 · 토의', '60분'], [id, 'AI 기초', '저작권과 개인정보', '사전 영상', '30분'],
    [id, '수업 설계', '성취기준 분석하기', '2주차 · 실습', '50분'], [id, '수업 설계', 'AI 활동 끼워 넣기', '2주차 · 설계', '70분'],
    [id, '수업 설계', '수업안 템플릿 채우기', '2주차 · 자료', '60분'], [id, '수업 설계', '동료 검토', '4주차 · 나눔', '90분'],
    [id, '평가', '루브릭 만들기', '3주차 · 실습', '60분'], [id, '평가', 'AI 피드백 초안 받기', '3주차 · 실습', '60분'],
    [id, '평가', '피드백 다듬기', '3주차 · 토의', '60분'], [id, '평가', '수업안 발표', '4주차 · 발표', '90분']]);
  makeTable_('features', [
    ['', '🧭', '수업 설계 중심', '도구 사용법보다 **성취기준에서 출발하는 설계**를 다룹니다.'],
    ['', '🤖', '매주 실습', '매주 **바로 쓸 수 있는 수업안 1개**를 완성해 가져갑니다.'],
    ['', '🌱', '함께 성장', '연수 후에도 **교사 학습공동체**에서 사례를 나눕니다.']]);
  makeTable_('faq', [
    [id, '연수 시간이 학점으로 인정되나요?', '네, 15시간을 모두 이수하면 교원 자율연수 1학점으로 인정됩니다. 출석은 매 회차 확인합니다.', '일정 보기', 'course'],
    ['', 'AI를 처음 써 보는데 따라갈 수 있을까요?', '첫 회차에 기초부터 함께 실습합니다. 노트북만 준비하면 계정 만들기부터 안내해 드려요.', '과정 보기', 'course'],
    ['', '온라인으로만 참여해도 되나요?', '네, 모든 회차를 실시간 온라인으로 함께 진행합니다. 실습 자료는 하루 전에 공유합니다.', '신청하러 가기', 'apply'],
    ['', '다른 학교 선생님도 신청할 수 있나요?', '경기도 초등 교원이면 누구나 신청할 수 있어요. 정원이 차면 대기로 접수됩니다.', '신청하러 가기', 'apply']]);
  makeTable_('reviews', [
    [id, '예', 'AI가 겁났는데 *첫 주부터* 우리 반 수업안이 나왔어요.', '김○○ 선생님', '초등 4학년 담임'],
    [id, '예', '평가 피드백에 쓰던 시간이 *눈에 띄게* 줄었습니다.', '이○○ 선생님', '초등 6학년 담임'],
    [id, '예', '설계부터 배우니 *다른 과목에도* 바로 옮겨 쓰게 돼요.', '박○○ 선생님', '과학 전담'],
    [id, '예', '동료 수업안을 보며 *아이디어가* 두 배로 늘었어요.', '최○○ 선생님', '초등 2학년 담임']]);
  makeTable_('apps', []);
  makeTable_('people', []);

  validate_('courses', 'status', COURSE_STATUS);
  validate_('reviews', 'show', ['예', '아니오']);
  validate_('apps', 'status', APP_STATUS);
  validate_('apps', 'done', ['이수', '미이수']);

  var first = ss.getSheets().filter(function (s) { return /^(시트|Sheet)\s?1$/.test(s.getName()) && s.getLastRow() === 0; })[0];
  if (first && ss.getSheets().length > 1) ss.deleteSheet(first);

  var msg = '탭 구성이 끝났습니다.\n\n관리자 비밀번호: ' + setting_('adminPassword') +
    '\n(Settings 탭에서 언제든 바꿀 수 있어요)\n\n다음 단계: [배포 > 새 배포 > 웹 앱]';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

function makeTable_(key, rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), t = T[key];
  if (ss.getSheetByName(t.name)) return;                     // 이미 있으면 그대로 둡니다
  var sh = ss.insertSheet(t.name);
  var header = t.cols.map(function (c) { return c[1]; });
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  var data = rows.map(function (r) { return Array.isArray(r) ? r : t.cols.map(function (c) { return r[c[0]] == null ? '' : r[c[0]]; }); });
  if (data.length) sh.getRange(2, 1, data.length, header.length).setValues(data);
  styleHeader_(sh, header.length, t.widths);
  // 전화번호·ID가 숫자로 바뀌지 않게 글자로 둡니다
  ['phone', 'id', 'appId', 'course'].forEach(function (f) {
    var i = colIndex_(t, f);
    if (i >= 0) sh.getRange(2, i + 1, 999, 1).setNumberFormat('@');
  });
}

function styleHeader_(sh, n, widths) {
  sh.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#e8eefc').setFontColor('#24324a');
  sh.setFrozenRows(1);
  (widths || []).forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
}

function validate_(key, field, list) {
  var i = colIndex_(T[key], field);
  if (i < 0) return;
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(T[key].name).getRange(2, i + 1, 999, 1).setDataValidation(rule);
}

function colIndex_(t, field) {
  for (var i = 0; i < t.cols.length; i++) if (t.cols[i][0] === field) return i;
  return -1;
}

/* ============ 표 읽기·쓰기 ============ */

function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }

function cell_(v) {
  if (v instanceof Date) {
    var hasTime = v.getHours() || v.getMinutes();
    return Utilities.formatDate(v, TZ, hasTime ? 'yyyy-MM-dd HH:mm' : 'yyyy-MM-dd');
  }
  return v == null ? '' : String(v);
}

function sheetOf_(key) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(T[key].name);
  if (!sh) throw new Error('"' + T[key].name + '" 탭이 없습니다. setup을 먼저 실행하세요.');
  return sh;
}

// 1행의 열 제목으로 열을 찾으니, 열 순서를 바꿔도 동작합니다
function read_(key) {
  var t = T[key], sh = sheetOf_(key);
  var last = sh.getLastRow(), width = Math.max(sh.getLastColumn(), t.cols.length);
  var all = sh.getRange(1, 1, Math.max(last, 1), width).getValues();
  var head = all[0].map(function (h) { return String(h).trim(); });
  var pos = {};
  t.cols.forEach(function (c) { pos[c[0]] = head.indexOf(c[1]); });
  var rows = [];
  for (var r = 1; r < all.length; r++) {
    var o = { _row: r + 1 }, empty = true;
    t.cols.forEach(function (c) {
      var v = pos[c[0]] >= 0 ? cell_(all[r][pos[c[0]]]) : '';
      o[c[0]] = v; if (v !== '') empty = false;
    });
    if (!empty) rows.push(o);
  }
  return { sh: sh, pos: pos, rows: rows, width: width };
}

function safe_(v) { v = String(v == null ? '' : v); return /^[=+\-@]/.test(v) ? "'" + v : v; }
function clean_(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u0008\u000b-\u001f]/g, ' ').trim().slice(0, max || 500); }

function setCells_(tb, rowNo, obj) {
  Object.keys(obj).forEach(function (f) {
    var p = tb.pos[f];
    if (p >= 0) tb.sh.getRange(rowNo, p + 1).setValue(safe_(obj[f]));
  });
}

function appendObj_(key, obj) {
  var tb = read_(key);
  var row = [];
  for (var i = 0; i < tb.width; i++) row.push('');
  T[key].cols.forEach(function (c) { if (tb.pos[c[0]] >= 0) row[tb.pos[c[0]]] = safe_(obj[c[0]] == null ? '' : obj[c[0]]); });
  tb.sh.appendRow(row);
}

// 조건에 맞는 줄을 지우고 새 줄로 바꿉니다 (다른 연수의 줄은 그대로)
function replaceRows_(key, keep, newObjs) {
  var tb = read_(key), sh = tb.sh;
  var kept = tb.rows.filter(keep);
  var all = kept.concat(newObjs);
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, tb.width).clearContent();
  if (!all.length) return;
  var data = all.map(function (o) {
    var row = []; for (var i = 0; i < tb.width; i++) row.push('');
    T[key].cols.forEach(function (c) { if (tb.pos[c[0]] >= 0) row[tb.pos[c[0]]] = safe_(o[c[0]] == null ? '' : o[c[0]]); });
    return row;
  });
  sh.getRange(2, 1, data.length, tb.width).setValues(data);
}

function settings_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_TAB);
  if (!sh) throw new Error('Settings 탭이 없습니다. setup을 먼저 실행하세요.');
  var o = {};
  sh.getDataRange().getValues().slice(1).forEach(function (r) { if (String(r[0]).trim()) o[String(r[0]).trim()] = cell_(r[1]); });
  return o;
}
function setting_(k) { return settings_()[k]; }

function setSetting_(k, v) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_TAB);
  var vals = sh.getDataRange().getValues(), at = -1;
  for (var i = 1; i < vals.length; i++) if (String(vals[i][0]).trim() === k) at = i + 1;
  if (at > 0) sh.getRange(at, 2).setValue(safe_(v)); else sh.appendRow([k, safe_(v), '']);
}

function findRow_(tb, id) {
  for (var i = 0; i < tb.rows.length; i++) if (tb.rows[i].id === id) return tb.rows[i]._row;
  return -1;
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ============ 전화번호 ============ */

function phone_(v) {
  var d = String(v || '').replace(/\D/g, '');
  if (!d) return '';
  if (/^02/.test(d)) return d.length > 9 ? d.replace(/^(02)(\d{4})(\d{4})$/, '$1-$2-$3') : d.replace(/^(02)(\d{3})(\d{4})$/, '$1-$2-$3');
  if (d.length === 8) return d.replace(/^(\d{4})(\d{4})$/, '$1-$2');
  if (d.length === 11) return d.replace(/^(\d{3})(\d{4})(\d{4})$/, '$1-$2-$3');
  if (d.length === 10) return d.replace(/^(\d{3})(\d{3})(\d{4})$/, '$1-$2-$3');
  return d;
}

/* ============ 연수 데이터 ============ */

function findCourse_(id) {
  var list = read_('courses').rows;
  if (id) { for (var i = 0; i < list.length; i++) if (list[i].id === id) return list[i]; return null; }
  return null;
}

function applyState_(c) {
  var t = today_();
  if (c.status === '준비') return { open: false, label: '준비 중이에요' };
  if (c.status !== '모집중') return { open: false, label: '신청이 마감됐어요' };
  if (c.openAt && t < c.openAt) return { open: false, label: c.openAt + '부터 신청할 수 있어요' };
  if (c.closeAt && t > c.closeAt) return { open: false, label: '신청이 마감됐어요' };
  return { open: true, label: '' };
}

function countActive_(courseId) {
  return read_('apps').rows.filter(function (a) { return a.course === courseId && a.status !== '취소'; }).length;
}

function siteSettings_(s) {
  return { siteTitle: s.siteTitle || '', brandName: s.brandName || '', brandInitials: s.brandInitials || '',
           contactPlace: s.contactPlace || '', contactPhone: s.contactPhone || '', footerText: s.footerText || '' };
}

function courseData_(id, admin) {
  var s = settings_();
  var isMain = !id;
  id = id || s.currentCourse;
  var c = findCourse_(id);
  if (!c && isMain) return { ok: true, preparing: true, settings: siteSettings_(s), archiveCount: archive_().length };
  if (!c) throw new Error('연수를 찾을 수 없습니다: ' + id);
  if (!admin && c.status === '준비') {
    if (isMain) return { ok: true, preparing: true, settings: siteSettings_(s), archiveCount: archive_().length };
    throw new Error('아직 공개되지 않은 연수입니다.');
  }
  var mine = function (r) { return r.course === c.id; };
  var mineOrCommon = function (r) { return r.course === c.id || r.course === ''; };
  var cap = parseInt(c.capacity, 10) || 0, applied = countActive_(c.id), st = applyState_(c);
  var reviews = read_('reviews').rows.filter(function (r) { return r.show !== '아니오'; });
  var own = reviews.filter(mine);
  return {
    ok: true,
    course: { id: c.id, name: c.name, status: c.status, openAt: c.openAt, closeAt: c.closeAt, completion: c.completion, isCurrent: c.id === s.currentCourse },
    settings: (function () { var o = siteSettings_(s); o.applyOpen = st.open; o.applyLabel = st.label; o.capacity = cap; o.deadlineText = c.applyText; return o; })(),
    archiveCount: archive_().length,
    content: { eyebrow: c.eyebrow, title1: c.title1, title2: c.title2, lead: c.lead, factPeriod: c.period, factTime: c.time,
               factCredit: c.credit, factPlace: c.place, courseLead: c.courseLead, voicesLead: c.voicesLead },
    features: read_('features').rows.filter(mineOrCommon).map(strip_),
    weeks: read_('sessions').rows.filter(mine).map(strip_),
    topics: read_('topics').rows.filter(mine).map(strip_),
    faq: read_('faq').rows.filter(mineOrCommon).map(strip_),
    reviews: (own.length ? own : reviews).map(strip_),       // 이 연수 후기가 없으면 노출된 후기 전체
    stats: { applied: applied, remaining: cap ? Math.max(0, cap - applied) : null },
    // 관리자 편집용 원본 (숨긴 후기 포함)
    edit: admin ? { course: strip_(c), reviews: read_('reviews').rows.filter(mine).map(strip_) } : undefined
  };
}
// 지난 연수 보기: 종료된 연수의 개요와 후기 (최근 것부터)
function archive_() {
  var reviews = read_('reviews').rows.filter(function (r) { return r.show !== '아니오'; });
  var sessions = read_('sessions').rows;
  return read_('courses').rows.filter(function (c) { return c.status === '종료'; }).reverse().map(function (c) {
    return {
      id: c.id, name: c.name, eyebrow: c.eyebrow, title1: c.title1, title2: c.title2, lead: c.lead,
      period: c.period, time: c.time, credit: c.credit, place: c.place,
      sessions: sessions.filter(function (x) { return x.course === c.id; }).map(function (x) { return { week: x.week, title: x.title }; }),
      reviews: reviews.filter(function (r) { return r.course === c.id; }).map(function (r) { return { quote: r.quote, name: r.name, role: r.role }; })
    };
  });
}

function strip_(o) { var r = {}; for (var k in o) if (k !== '_row') r[k] = o[k]; return r; }

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.view === 'archive') return json_({ ok: true, archive: archive_() });
    return json_(courseData_(p.c, false));
  }
  catch (err) { return json_({ ok: false, error: String(err.message || err) }); }
}

function doPost(e) {
  try {
    var b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (b.action === 'apply') return json_(apply_(b));
    if (b.action === 'admin') return json_(admin_(b));
    return json_({ ok: false, error: '알 수 없는 요청입니다.' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

/* ============ 신청 ============ */

function apply_(b) {
  if (b.website) return { ok: true };                         // 봇이 채우는 숨은 칸
  var s = settings_();
  var c = findCourse_(b.course || s.currentCourse);
  if (!c) return { ok: false, error: '연수를 찾을 수 없습니다.' };
  var st = applyState_(c);
  if (!st.open) return { ok: false, error: st.label };
  var name = clean_(b.name, 40), email = clean_(b.email, 120).toLowerCase(), school = clean_(b.school, 80), phone = phone_(b.phone);
  if (!name || !school || !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return { ok: false, error: '입력값을 확인해 주세요.' };
  if (phone && !/^\d{2,4}-?\d{3,4}-\d{4}$/.test(phone)) return { ok: false, error: '전화번호를 확인해 주세요.' };

  var res = withLock_(function () {
    var apps = read_('apps').rows;
    var dup = apps.some(function (a) { return a.course === c.id && a.email.toLowerCase() === email && a.status !== '취소'; });
    if (dup) return { ok: false, error: '이미 이 메일로 신청하셨어요.' };
    var cap = parseInt(c.capacity, 10) || 0;
    var active = apps.filter(function (a) { return a.course === c.id && a.status !== '취소'; }).length;
    var status = cap && active >= cap ? '대기' : '접수';
    var appId = c.id + '-' + ('000' + (apps.filter(function (a) { return a.course === c.id; }).length + 1)).slice(-3);
    appendObj_('apps', { appId: appId, course: c.id, at: now_(), name: name, email: email, school: school, phone: phone, status: status, done: '', memo: '' });
    return { ok: true, status: status, appId: appId };
  });
  if (!res.ok) return res;
  refreshPeople_();
  if (s.notifyEmail) {
    try {
      MailApp.sendEmail(s.notifyEmail, '[' + c.name + '] 새 신청: ' + name + ' (' + school + ')',
        '연수: ' + c.name + ' (' + c.id + ')\n이름: ' + name + '\n학교 메일: ' + email + '\n소속: ' + school +
        '\n전화: ' + (phone || '-') + '\n상태: ' + res.status + '\n\n시트: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl());
    } catch (err) { /* 메일이 실패해도 신청은 저장됩니다 */ }
  }
  var cap2 = parseInt(c.capacity, 10) || 0, applied = countActive_(c.id);
  res.stats = { applied: applied, remaining: cap2 ? Math.max(0, cap2 - applied) : null };
  return res;
}

// 참가자 탭: 학교 메일 기준 누적 집계 (신청 탭에서 다시 계산)
function refreshPeople_() {
  var apps = read_('apps').rows, by = {};
  apps.forEach(function (a) {
    var k = (a.email || '').toLowerCase();
    if (!k) return;
    var p = by[k] || (by[k] = { email: k, name: '', school: '', phone: '', applied: 0, confirmed: 0, completed: 0, lastCourse: '', lastAt: '' });
    if (a.status !== '취소') p.applied++;
    if (a.status === '확정') p.confirmed++;
    if (a.done === '이수') p.completed++;
    if (a.at >= p.lastAt) { p.lastAt = a.at; p.lastCourse = a.course; p.name = a.name; p.school = a.school; p.phone = a.phone || p.phone; }
  });
  var list = Object.keys(by).map(function (k) { return by[k]; }).sort(function (x, y) { return x.lastAt < y.lastAt ? 1 : -1; });
  replaceRows_('people', function () { return false; }, list);
}

/* ============ 관리자 ============ */

var SETTINGS_EDITABLE = ['siteTitle', 'brandName', 'brandInitials', 'adminPassword', 'currentCourse', 'notifyEmail', 'contactPlace', 'contactPhone', 'footerText'];
var COURSE_EDITABLE = ['status', 'name', 'eyebrow', 'title1', 'title2', 'lead', 'period', 'time', 'credit', 'place', 'capacity', 'openAt', 'closeAt', 'applyText', 'courseLead', 'voicesLead', 'completion'];

function adminSettings_() {
  var s = settings_(), o = {};
  SETTINGS_EDITABLE.forEach(function (k) { if (k !== 'adminPassword') o[k] = s[k] == null ? '' : s[k]; });
  return o;
}

function courseList_() {
  var apps = read_('apps').rows, cur = setting_('currentCourse');
  return read_('courses').rows.map(function (c) {
    var mine = apps.filter(function (a) { return a.course === c.id; });
    return {
      id: c.id, name: c.name, status: c.status, period: c.period, capacity: parseInt(c.capacity, 10) || 0,
      openAt: c.openAt, closeAt: c.closeAt, createdAt: c.createdAt, isCurrent: c.id === cur,
      applied: mine.filter(function (a) { return a.status !== '취소'; }).length,
      confirmed: mine.filter(function (a) { return a.status === '확정'; }).length,
      completed: mine.filter(function (a) { return a.done === '이수'; }).length
    };
  }).reverse();
}

function apps_(courseId) {
  return read_('apps').rows.filter(function (a) { return !courseId || a.course === courseId; })
    .map(strip_).reverse();
}

function admin_(b) {
  var real = String(setting_('adminPassword') || '');
  if (!real || String(b.password || '') !== real) {
    Utilities.sleep(800);
    return { ok: false, error: '비밀번호가 맞지 않습니다.' };
  }
  var op = b.op;

  if (op === 'login') return { ok: true, settings: adminSettings_(), courses: courseList_() };
  if (op === 'view') return { ok: true, data: courseData_(b.course, true) };
  if (op === 'courses') return { ok: true, courses: courseList_(), settings: adminSettings_() };

  if (op === 'saveSettings') {
    withLock_(function () {
      Object.keys(b.values || {}).forEach(function (k) {
        if (SETTINGS_EDITABLE.indexOf(k) < 0) return;
        var v = clean_(b.values[k], 1000);
        if (k === 'adminPassword' && v.length < 4) return;
        if (k === 'currentCourse' && !findCourse_(v)) throw new Error('없는 연수 ID입니다: ' + v);
        setSetting_(k, v);
      });
    });
    return { ok: true, settings: adminSettings_(), courses: courseList_(), data: courseData_(b.course, true) };
  }

  if (op === 'saveCourse') {
    withLock_(function () {
      var tb = read_('courses'), c = null;
      tb.rows.forEach(function (r) { if (r.id === b.course) c = r; });
      if (!c) throw new Error('연수를 찾을 수 없습니다.');
      var upd = {};
      Object.keys(b.values || {}).forEach(function (k) {
        if (COURSE_EDITABLE.indexOf(k) < 0) return;
        var v = clean_(b.values[k], 2000);
        if (k === 'status' && COURSE_STATUS.indexOf(v) < 0) return;
        if (k === 'capacity') v = String(parseInt(v, 10) || 0);
        if ((k === 'openAt' || k === 'closeAt') && v && !/^\d{4}-\d{2}-\d{2}$/.test(v)) throw new Error('날짜는 2026-12-20 형식으로 입력하세요.');
        upd[k] = v;
      });
      setCells_(tb, c._row, upd);
    });
    return { ok: true, data: courseData_(b.course, true), courses: courseList_() };
  }

  if (op === 'saveList') {
    if (LISTS.indexOf(b.tab) < 0) return { ok: false, error: '고칠 수 없는 목록입니다.' };
    if (!findCourse_(b.course)) return { ok: false, error: '연수를 찾을 수 없습니다.' };
    var cols = T[b.tab].cols.map(function (c) { return c[0]; });
    var common = COMMON_OK.indexOf(b.tab) >= 0;
    var rows = (b.rows || []).slice(0, 200).map(function (r) {
      var o = {};
      cols.forEach(function (k) { o[k] = clean_(r[k], 1000); });
      o.course = common && r.course === '' ? '' : b.course;
      if (b.tab === 'reviews' && o.show !== '아니오') o.show = '예';
      return o;
    }).filter(function (o) { return cols.some(function (k) { return k !== 'course' && k !== 'show' && o[k]; }); });
    withLock_(function () {
      replaceRows_(b.tab, function (r) { return !(r.course === b.course || (common && r.course === '')); }, rows);
    });
    return { ok: true, data: courseData_(b.course, true) };
  }

  // 새 연수 시작: 개요를 받아 새 연수를 만들고, 지금의 대표 연수는 '종료'로 아카이브합니다
  if (op === 'startCourse') {
    var v = b.values || {};
    var id = clean_(v.id, 30);
    if (!/^[A-Za-z0-9_-]{2,30}$/.test(id)) return { ok: false, error: '연수 ID는 영문·숫자·- 로 2~30자로 정해 주세요 (예: 2027-S01)' };
    if (findCourse_(id)) return { ok: false, error: '이미 있는 연수 ID입니다.' };
    if (!clean_(v.name, 120)) return { ok: false, error: '연수명을 입력해 주세요.' };
    var prevId = setting_('currentCourse'), prev = findCourse_(prevId);
    withLock_(function () {
      var o = { id: id, createdAt: today_() };
      COURSE_EDITABLE.forEach(function (k) { o[k] = clean_(v[k], 2000); });
      if (COURSE_STATUS.indexOf(o.status) < 0 || o.status === '종료') o.status = '준비';
      o.capacity = String(parseInt(o.capacity, 10) || 0);
      ['openAt', 'closeAt'].forEach(function (k) { if (o[k] && !/^\d{4}-\d{2}-\d{2}$/.test(o[k])) o[k] = ''; });
      appendObj_('courses', o);
      if (prev && b.copySessions) {
        ['sessions', 'topics'].forEach(function (key) {
          read_(key).rows.filter(function (r) { return r.course === prev.id; }).forEach(function (r) {
            var x = strip_(r); x.course = id; appendObj_(key, x);
          });
        });
      }
      if (prev && prev.status !== '종료') { var tb = read_('courses'); setCells_(tb, findRow_(tb, prev.id), { status: '종료' }); }
      setSetting_('currentCourse', id);
    });
    return { ok: true, id: id, archived: prev ? prev.id : '', courses: courseList_(), settings: adminSettings_() };
  }

  if (op === 'apps') return { ok: true, applications: apps_(b.course), statuses: APP_STATUS, courses: courseList_(), sheetUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl() };

  if (op === 'updateApp') {
    withLock_(function () {
      var tb = read_('apps'), a = null;
      tb.rows.forEach(function (r) { if (r.appId === b.appId) a = r; });
      if (!a) throw new Error('신청을 찾을 수 없습니다.');
      var upd = {};
      if (b.status != null) { if (APP_STATUS.indexOf(b.status) < 0) throw new Error('잘못된 상태입니다.'); upd.status = b.status; }
      if (b.done != null) { if (DONE.indexOf(b.done) < 0) throw new Error('잘못된 이수 값입니다.'); upd.done = b.done; }
      if (b.memo != null) upd.memo = clean_(b.memo, 300);
      if (b.phone != null) upd.phone = phone_(b.phone);
      setCells_(tb, a._row, upd);
    });
    refreshPeople_();
    return { ok: true, applications: apps_(b.course), statuses: APP_STATUS, courses: courseList_() };
  }

  return { ok: false, error: '알 수 없는 관리자 요청입니다.' };
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
