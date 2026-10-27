/**
 * '연수 관리' 시트 ↔ AI 수업 설계 연수 사이트 연동 스크립트
 *
 * 처음 한 번: 메뉴 [확장 프로그램 > Apps Script]에 이 코드를 붙여넣고 setup 함수를 실행하세요.
 *   → Settings, 내용, 특징, 주차, 세부과정, FAQ, 후기, 신청 탭이 만들어지고 예시 내용이 채워집니다.
 *   → 이미 있는 탭은 건드리지 않습니다.
 * 그다음: [배포 > 새 배포 > 웹 앱]으로 배포하고, 나온 웹 앱 URL을 사이트에 넣습니다.
 *
 * 관리자 비밀번호는 Settings 탭의 adminPassword 값입니다. 비밀번호는 이 스크립트 안에서만 비교하고
 * 사이트로 내보내지 않습니다.
 */

var TABS = {
  settings: 'Settings',
  content: '내용',
  features: '특징',
  weeks: '주차',
  topics: '세부과정',
  faq: 'FAQ',
  reviews: '후기',
  apply: '신청'
};

var APPLY_HEADER = ['신청일시', '이름', '학교 메일', '소속 학교', '전화번호', '상태', '메모'];
var STATUSES = ['접수', '확정', '대기', '취소'];

/* ---------- 처음 설정 ---------- */

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pw = Math.random().toString(36).slice(2, 10);

  makeTab_(ss, TABS.settings, ['키', '값', '설명'], [
    ['siteTitle', 'AI 수업 설계 연수', '브라우저 탭에 보이는 사이트 제목'],
    ['brandName', 'JP Park 연수실', '왼쪽 위 로고 옆 이름'],
    ['brandInitials', 'JP', '로고 상자 안 글자 (2자 권장)'],
    ['adminPassword', pw, '사이트 관리자 로그인 비밀번호. 바꾸려면 이 칸만 고치면 됩니다'],
    ['applyOpen', '예', '신청 받기: 예 / 아니오'],
    ['capacity', 30, '정원 (숫자). 남은 자리는 신청 탭의 접수·확정 건수로 자동 계산'],
    ['deadlineText', '12월 20일까지 선착순으로 받습니다. 신청 후 3일 안에 학교 메일로 확정 안내를 보내 드립니다.', '신청 섹션 안내 문구'],
    ['notifyEmail', '', '신청이 들어오면 알림 메일을 받을 주소 (비워 두면 보내지 않음)'],
    ['contactPlace', '율포초등학교 수석교사실', '문의처'],
    ['contactPhone', '031-000-0000', '문의 전화'],
    ['footerText', '© 2026 박종필 · AI와 디지털로 더 나은 배움을 설계합니다.', '맨 아래 문구']
  ], [180, 420, 420]);

  makeTab_(ss, TABS.content, ['키', '값', '설명'], [
    ['eyebrow', '✨ 2027 겨울방학 교원 자율연수', '제목 위 알약 라벨'],
    ['title1', 'AI로 수업을 다시 설계하는', '큰 제목 첫 줄'],
    ['title2', '4주간의 실습 연수', '큰 제목 둘째 줄 (그라디언트)'],
    ['lead', '생성형 AI를 **우리 반 교육과정에 자연스럽게 연결하는 방법**을, 매주 직접 만든 수업안으로 익힙니다. 이론보다 실습, 도구보다 **수업 설계**에 집중합니다.', '소개 문장. **굵게** 표시 가능'],
    ['factPeriod', '1월 6일 – 1월 27일 · 매주 수', '연수 한눈에: 기간'],
    ['factTime', '16:00 – 19:00 (주 3시간)', '연수 한눈에: 시간'],
    ['factCredit', '15시간 · 1학점 인정', '연수 한눈에: 이수'],
    ['factPlace', '율포초 미래교실 + 온라인 병행', '연수 한눈에: 장소'],
    ['courseLead', '주제별로 나눠 보면 어떤 회차가 내 수업에 맞는지 바로 고를 수 있어요. 오른쪽 탭을 눌러 보세요.', '과정 섹션 설명'],
    ['voicesLead', '지난 여름 연수에 참여한 선생님들의 후기와 자주 받은 질문을 모았습니다.', '후기·질문 섹션 설명']
  ], [180, 420, 300]);

  makeTab_(ss, TABS.features, ['아이콘', '제목', '설명 (**굵게** 가능)'], [
    ['🧭', '수업 설계 중심', '도구 사용법보다 **성취기준에서 출발하는 설계**를 다룹니다.'],
    ['🤖', 'AI 실습 4회', '매주 **바로 쓸 수 있는 수업안 1개**를 완성해 가져갑니다.'],
    ['🌱', '함께 성장', '연수 후에도 **교사 학습공동체**에서 사례를 나눕니다.']
  ], [80, 180, 460]);

  makeTab_(ss, TABS.weeks, ['주차', '제목', '설명', '시간'], [
    ['1주', 'AI 리터러시와 수업 윤리', '학생과 함께 정하는 AI 사용 규칙', '3시간'],
    ['2주', '성취기준으로 시작하는 AI 수업', '백워드 설계에 AI 활동 끼워 넣기', '3시간'],
    ['3주', '피드백과 평가 도우미', '루브릭 기반 피드백 자동화', '3시간'],
    ['4주', '수업 나눔과 공개', '완성한 수업안 발표와 상호 피드백', '3시간']
  ], [70, 260, 300, 80]);

  makeTab_(ss, TABS.topics, ['주제 (탭 이름)', '제목', '부가 정보', '시간'], [
    ['AI 기초', '생성형 AI 작동 원리', '1주차 · 강의', '40분'],
    ['AI 기초', '프롬프트 작성 실습', '1주차 · 실습', '80분'],
    ['AI 기초', 'AI 사용 규칙 만들기', '1주차 · 토의', '60분'],
    ['AI 기초', '저작권과 개인정보', '사전 영상', '30분'],
    ['수업 설계', '성취기준 분석하기', '2주차 · 실습', '50분'],
    ['수업 설계', 'AI 활동 끼워 넣기', '2주차 · 설계', '70분'],
    ['수업 설계', '수업안 템플릿 채우기', '2주차 · 자료', '60분'],
    ['수업 설계', '동료 검토', '4주차 · 나눔', '90분'],
    ['평가', '루브릭 만들기', '3주차 · 실습', '60분'],
    ['평가', 'AI 피드백 초안 받기', '3주차 · 실습', '60분'],
    ['평가', '피드백 다듬기', '3주차 · 토의', '60분'],
    ['평가', '수업안 발표', '4주차 · 발표', '90분']
  ], [120, 240, 160, 80]);

  makeTab_(ss, TABS.faq, ['질문', '답변', '링크 문구', '이동 위치 (course / apply)'], [
    ['연수 시간이 학점으로 인정되나요?', '네, 15시간을 모두 이수하면 교원 자율연수 1학점으로 인정됩니다. 출석은 매 회차 확인합니다.', '일정 보기', 'course'],
    ['AI를 처음 써 보는데 따라갈 수 있을까요?', '1주차에 기초부터 함께 실습합니다. 노트북만 준비하면 계정 만들기부터 안내해 드려요.', '1주차 내용 보기', 'course'],
    ['온라인으로만 참여해도 되나요?', '네, 모든 회차를 실시간 온라인으로 함께 진행합니다. 실습 자료는 하루 전에 공유합니다.', '신청하러 가기', 'apply'],
    ['다른 학교 선생님도 신청할 수 있나요?', '경기도 초등 교원이면 누구나 신청할 수 있어요. 정원 30명, 선착순입니다.', '신청하러 가기', 'apply']
  ], [280, 420, 140, 160]);

  makeTab_(ss, TABS.reviews, ['후기 (*강조할 말*은 별표로 감싸기)', '이름', '소속'], [
    ['AI가 겁났는데 *첫 주부터* 우리 반 수업안이 나왔어요.', '김○○ 선생님', '초등 4학년 담임'],
    ['평가 피드백에 쓰던 시간이 *눈에 띄게* 줄었습니다.', '이○○ 선생님', '초등 6학년 담임'],
    ['설계부터 배우니 *다른 과목에도* 바로 옮겨 쓰게 돼요.', '박○○ 선생님', '과학 전담'],
    ['동료 수업안을 보며 *아이디어가* 두 배로 늘었어요.', '최○○ 선생님', '초등 2학년 담임']
  ], [420, 140, 160]);

  var ap = makeTab_(ss, TABS.apply, APPLY_HEADER, [], [150, 100, 220, 160, 130, 80, 240]);
  if (ap) {
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build();
    ap.getRange('F2:F1000').setDataValidation(rule);
  }

  var msg = '탭 구성이 끝났습니다.\n\n관리자 비밀번호: ' + getSetting_('adminPassword') +
    '\n(Settings 탭에서 언제든 바꿀 수 있어요)\n\n다음 단계: [배포 > 새 배포 > 웹 앱]';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

function makeTab_(ss, name, header, rows, widths) {
  if (ss.getSheetByName(name)) return null;               // 이미 있으면 그대로 둡니다
  var sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground('#e8eefc').setFontColor('#24324a');
  if (rows.length) sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  sh.setFrozenRows(1);
  (widths || []).forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.getRange(1, 1, Math.max(rows.length + 1, 2), header.length).setWrap(true).setVerticalAlignment('top');
  // 처음 시트에 있던 빈 '시트1'은 정리합니다
  var first = ss.getSheets()[0];
  if (first.getName() !== name && /^(시트|Sheet)\s?1$/.test(first.getName()) && first.getLastRow() === 0) {
    ss.deleteSheet(first);
  }
  return sh;
}

/* ---------- 읽기 도우미 ---------- */

function sheet_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('"' + name + '" 탭이 없습니다. setup을 먼저 실행하세요.');
  return sh;
}

function rows_(name) {
  var v = sheet_(name).getDataRange().getDisplayValues();
  return v.slice(1).filter(function (r) { return r.join('').trim() !== ''; });
}

function keyValues_(name) {
  var o = {};
  rows_(name).forEach(function (r) { if (r[0]) o[String(r[0]).trim()] = r[1]; });
  return o;
}

function getSetting_(key) { return keyValues_(TABS.settings)[key]; }

function counts_() {
  var sh = sheet_(TABS.apply), n = sh.getLastRow() - 1, active = 0;
  if (n > 0) {
    sh.getRange(2, 6, n, 1).getDisplayValues().forEach(function (r) {
      if (r[0] !== '취소') active++;
    });
  }
  return active;
}

/* ---------- 사이트가 읽는 공개 내용 ---------- */

function publicData_() {
  var s = keyValues_(TABS.settings);
  var capacity = parseInt(s.capacity, 10) || 0;
  var applied = counts_();
  return {
    ok: true,
    settings: {
      siteTitle: s.siteTitle || '', brandName: s.brandName || '', brandInitials: s.brandInitials || '',
      applyOpen: String(s.applyOpen || '예').trim() !== '아니오',
      capacity: capacity, deadlineText: s.deadlineText || '',
      contactPlace: s.contactPlace || '', contactPhone: s.contactPhone || '', footerText: s.footerText || ''
    },
    content: keyValues_(TABS.content),
    features: rows_(TABS.features).map(function (r) { return { icon: r[0], title: r[1], text: r[2] }; }),
    weeks: rows_(TABS.weeks).map(function (r) { return { week: r[0], title: r[1], desc: r[2], time: r[3] }; }),
    topics: rows_(TABS.topics).map(function (r) { return { topic: r[0], title: r[1], meta: r[2], time: r[3] }; }),
    faq: rows_(TABS.faq).map(function (r) { return { q: r[0], a: r[1], linkText: r[2], target: r[3] }; }),
    reviews: rows_(TABS.reviews).map(function (r) { return { quote: r[0], name: r[1], role: r[2] }; }),
    stats: { applied: applied, remaining: capacity ? Math.max(0, capacity - applied) : null }
  };
}

function doGet() {
  try { return json_(publicData_()); }
  catch (err) { return json_({ ok: false, error: String(err.message || err) }); }
}

/* ---------- 신청 접수와 관리자 기능 ---------- */

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'apply') return json_(apply_(body));
    if (body.action === 'admin') return json_(admin_(body));
    return json_({ ok: false, error: '알 수 없는 요청입니다.' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function clean_(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max || 200); }

function apply_(b) {
  var s = keyValues_(TABS.settings);
  if (String(s.applyOpen || '예').trim() === '아니오') return { ok: false, error: '지금은 신청을 받지 않습니다.' };
  if (b.website) return { ok: true };                       // 자동 입력 봇이 채우는 숨은 칸
  var name = clean_(b.name, 40), email = clean_(b.email, 120), school = clean_(b.school, 80), phone = clean_(b.phone, 30);
  if (!name || !school || !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return { ok: false, error: '입력값을 확인해 주세요.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = sheet_(TABS.apply);
    var n = sh.getLastRow() - 1;
    if (n > 0) {
      var dup = sh.getRange(2, 3, n, 4).getDisplayValues().some(function (r) {
        return r[0].toLowerCase() === email.toLowerCase() && r[3] !== '취소';
      });
      if (dup) return { ok: false, error: '이미 이 메일로 신청하셨어요.' };
    }
    var capacity = parseInt(s.capacity, 10) || 0;
    var status = capacity && counts_() >= capacity ? '대기' : '접수';
    // 수식으로 해석되지 않도록 앞에 ' 를 붙입니다
    var safe = function (v) { return /^[=+\-@]/.test(v) ? "'" + v : v; };
    var at = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
    sh.appendRow([at, safe(name), safe(email), safe(school), safe(phone), status, '']);
  } finally {
    lock.releaseLock();
  }

  if (s.notifyEmail) {
    try {
      MailApp.sendEmail(s.notifyEmail, '[연수 신청] ' + name + ' (' + school + ')',
        '새 신청이 들어왔습니다.\n\n이름: ' + name + '\n학교 메일: ' + email + '\n소속: ' + school +
        '\n전화: ' + (phone || '-') + '\n상태: ' + status + '\n\n시트: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl());
    } catch (err) { /* 메일이 실패해도 신청은 저장됩니다 */ }
  }
  return { ok: true, status: status, stats: publicData_().stats };
}

function admin_(b) {
  var real = String(getSetting_('adminPassword') || '');
  if (!real || String(b.password || '') !== real) {
    Utilities.sleep(800);                                   // 비밀번호 대입을 느리게 합니다
    return { ok: false, error: '비밀번호가 맞지 않습니다.' };
  }
  if (b.op === 'login') return { ok: true, settings: adminSettings_() };
  if (b.op === 'saveKV') { saveKV_(b.values || {}); return { ok: true, data: publicData_(), settings: adminSettings_() }; }
  if (b.op === 'saveList') { saveList_(b.tab, b.rows || []); return { ok: true, data: publicData_() }; }

  var sh = sheet_(TABS.apply);
  if (b.op === 'status') {
    var row = parseInt(b.row, 10);
    if (!(row >= 2 && row <= sh.getLastRow()) || STATUSES.indexOf(b.status) < 0) return { ok: false, error: '잘못된 요청입니다.' };
    sh.getRange(row, 6).setValue(b.status);
  }
  if (b.op === 'memo') {
    var r2 = parseInt(b.row, 10);
    if (!(r2 >= 2 && r2 <= sh.getLastRow())) return { ok: false, error: '잘못된 요청입니다.' };
    sh.getRange(r2, 7).setValue(clean_(b.memo, 300));
  }
  var n = sh.getLastRow() - 1;
  var list = n > 0 ? sh.getRange(2, 1, n, 7).getDisplayValues().map(function (r, i) {
    return { row: i + 2, at: r[0], name: r[1], email: r[2], school: r[3], phone: r[4], status: r[5], memo: r[6] };
  }) : [];
  return {
    ok: true,
    applications: list.reverse(),
    statuses: STATUSES,
    sheetUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl(),
    stats: publicData_().stats
  };
}

/* ---------- 사이트의 ✏️ 편집 저장 ---------- */

// 사이트에서 고칠 수 있는 키 (그 밖의 키는 시트에서만 고칩니다)
var EDITABLE = {
  settings: ['siteTitle', 'brandName', 'brandInitials', 'adminPassword', 'applyOpen', 'capacity', 'deadlineText',
             'notifyEmail', 'contactPlace', 'contactPhone', 'footerText'],
  content: ['eyebrow', 'title1', 'title2', 'lead', 'factPeriod', 'factTime', 'factCredit', 'factPlace', 'courseLead', 'voicesLead']
};
var LIST_COLS = { features: 3, weeks: 4, topics: 4, faq: 4, reviews: 3 };

function adminSettings_() {
  var s = keyValues_(TABS.settings), o = {};
  EDITABLE.settings.forEach(function (k) { if (k !== 'adminPassword') o[k] = s[k] == null ? '' : s[k]; });
  return o;
}

// values 예: { 'content.title1': '…', 'settings.capacity': '30' }
function saveKV_(values) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    Object.keys(values).forEach(function (full) {
      var p = full.split('.'), tab = p[0], key = p[1];
      if (!EDITABLE[tab] || EDITABLE[tab].indexOf(key) < 0) return;
      var val = clean_(values[full], 2000);
      if (key === 'adminPassword' && val.length < 4) return;      // 비워 두면 기존 비밀번호 유지
      if (key === 'applyOpen') val = (val === '아니오' || val === 'false') ? '아니오' : '예';
      var sh = sheet_(TABS[tab]), last = sh.getLastRow();
      var keys = last > 1 ? sh.getRange(2, 1, last - 1, 1).getDisplayValues() : [];
      var at = -1;
      keys.forEach(function (r, i) { if (String(r[0]).trim() === key) at = i + 2; });
      if (/^[=+\-@]/.test(val)) val = "'" + val;
      if (at > 0) sh.getRange(at, 2).setValue(val);
      else sh.appendRow([key, val, '']);
    });
  } finally { lock.releaseLock(); }
}

function saveList_(tab, rows) {
  var cols = LIST_COLS[tab];
  if (!cols) throw new Error('고칠 수 없는 목록입니다.');
  var data = rows.slice(0, 100).map(function (r) {
    var out = [];
    for (var i = 0; i < cols; i++) {
      var v = clean_(r[i], 1000);
      out.push(/^[=+\-@]/.test(v) ? "'" + v : v);
    }
    return out;
  }).filter(function (r) { return r.join('').trim() !== ''; });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = sheet_(TABS[tab]), last = sh.getLastRow();
    if (last > 1) sh.getRange(2, 1, last - 1, cols).clearContent();
    if (data.length) sh.getRange(2, 1, data.length, cols).setValues(data);
  } finally { lock.releaseLock(); }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
