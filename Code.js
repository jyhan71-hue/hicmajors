/**
 * 한양YK인터칼리지 융합전공 소개행사 시스템 (V8)
 *
 * ── 시트 구성 ──────────────────────────────────────────────────────────────
 * [Settings]        A열=항목명, B열=값, C열=시작시간(세션전용), D열=종료시간(세션전용)
 *   예약오픈여부   | OPEN / CLOSED
 *   행사명         | (행사 부제목)
 *   행사일시       | (날짜·시간 문자열)
 *   행사장소       | (장소 문자열)
 *   currentSessionIdx | 0
 *   설명회세션     | 프로그램명 | 시작HH:MM | 종료HH:MM  (순서대로 반복)
 *   부스프로그램   | 프로그램명                             (순서대로 반복)
 *
 * [AdminUsers]      A=이름, B=호칭, C=담당프로그램, D=비밀번호
 *                   담당프로그램='전체관리' → 행정팀 계정
 *
 * [BoothReservations]  부스 상담 예약
 *                   이름|학번|학과|이메일|연락처|프로그램|시간|문의내용|서명|상태|코멘트|예약일시
 *
 * [SessionPreReg]   설명회 사전예약 (선착순 30명/프로그램)
 *                   이름|학번|학과|연락처|이메일|설명회명|등록일시
 *
 * [CheckIns]        설명회 현장 체크인: 체크인시각|학번|이름|학과|연락처|이메일|참석유형|참석설명회
 *                   참석설명회당 1행 (설명회 없으면 1행, 설명회 3개면 3행)
 *
 * [GiftReceipts]    기념품 수령 서명 (예산 증빙)
 *                   수령시각|학번|이름|학과|스탬프수|가산점과목|서명
 *
 * [BlockedSlots]    부스 슬롯 수동 차단
 *                   프로그램|시간|차단여부(TRUE/FALSE)
 *
 * ──────────────────────────────────────────────────────────────────────────
 */

// ─────────────────────────────────────────────
// 상수
// ─────────────────────────────────────────────
var SESSION_CAPACITY    = 30;  // 설명회 사전예약 프로그램별 최대 인원
var WALK_IN_CAPACITY    = 20;  // 설명회 당일방문 최대 인원 (전체 합산)
var BOOTH_STAMP_REQUIRED = 3;  // 기념품 수령 최소 스탬프 수
var INTERCOLLEGE_DEPT   = '한양인터칼리지학부';
var BONUS_SUBJECTS = [
  'Life Project',
  'Unified Systems Science',
  'Systems Thinking and Design Thinking',
  'Algorithmic, Computational, and Data Thinking',
  'Mathematical Thinking'
];

// 참여확인서 위조 방지 설정
// SEAL_IMAGE_URL: 학장 직인 이미지 Google Drive 공유 URL (썸네일 형식 권장)
// ex) https://drive.google.com/thumbnail?id=YOUR_FILE_ID&sz=w200
var SEAL_IMAGE_URL = ''; // ← 직인 이미지 URL을 여기에 입력하세요
var VERIFY_PAGE_URL = 'https://tinyurl.com/hicmajors'; // 검증 페이지 기본 URL (배포 후 교체)

// 설명회 진행은 Settings 시트 'currentSessionIdx' 키로 관리
// 스태프가 "다음 세션 →" 버튼을 눌러 인덱스를 직접 진행시킴

// ─────────────────────────────────────────────
// 유틸
// ─────────────────────────────────────────────
function toTimeStr(val) {
  if (!val && val !== 0) return '';
  if (typeof val === 'string') {
    var s = val.trim();
    if (/^\d{2}:\d{2}\s*~/.test(s)) return s.split('~')[0].trim();
    if (/^\d{2}:\d{2}$/.test(s)) return s;
  }
  if (val instanceof Date) {
    return String(val.getHours()).padStart(2,'0') + ':' + String(val.getMinutes()).padStart(2,'0');
  }
  if (typeof val === 'number') {
    var totalMin = Math.round(val * 24 * 60);
    return String(Math.floor(totalMin/60)).padStart(2,'0') + ':' + String(totalMin%60).padStart(2,'0');
  }
  return val.toString().trim();
}

function addMinutes(timeStr, mins) {
  var parts = timeStr.split(':');
  var total = parseInt(parts[0],10)*60 + parseInt(parts[1],10) + mins;
  return String(Math.floor(total/60)).padStart(2,'0') + ':' + String(total%60).padStart(2,'0');
}

function slotDuration(programName) { return 15; }

function toMin(t) {
  var p = t.split(':');
  return parseInt(p[0],10)*60 + parseInt(p[1],10);
}

function timesOverlap(aStart, aDur, bStart, bDur) {
  var aS = toMin(aStart), aE = aS + aDur;
  var bS = toMin(bStart), bE = bS + bDur;
  return aS < bE && bS < aE;
}


// ─────────────────────────────────────────────
// 시트 초기화 (최초 1회 에디터에서 직접 실행)
// 필요한 모든 시트와 헤더를 자동 생성합니다
// ─────────────────────────────────────────────
// ─────────────────────────────────────────────
// Settings 시트 전체 설정 읽기
// ─────────────────────────────────────────────
function _getSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Settings');
  var result = { isOpen:true, eventName:'', eventDate:'', eventPlace:'', sessions:[], boothPrograms:[] };
  if (!sh) return result;
  var rows = sh.getDataRange().getValues();
  for (var i = 0; i < rows.length; i++) {
    var key = rows[i][0] ? rows[i][0].toString().trim() : '';
    var val = rows[i][1] ? rows[i][1].toString().trim() : '';
    if      (key === '예약오픈여부')  result.isOpen = val.toUpperCase() === 'OPEN';
    else if (key === '행사명')        result.eventName = val;
    else if (key === '행사일시')      result.eventDate = val;
    else if (key === '행사장소')      result.eventPlace = val;
    else if (key === '설명회세션' && val) {
      result.sessions.push({
        name:  val,
        start: rows[i][2] ? toTimeStr(rows[i][2]) : '',
        end:   rows[i][3] ? toTimeStr(rows[i][3]) : ''
      });
    }
    else if (key === '부스프로그램' && val) result.boothPrograms.push(val);
  }
  return result;
}

// HTML 템플릿에 주입할 설정 데이터 반환 (JSON 문자열)
function getSettingsData() {
  var s = _getSettings();
  var sessionTimes = {};
  s.sessions.forEach(function(sess) { sessionTimes[sess.name] = sess.start + '~' + sess.end; });
  s.sessionTimes = sessionTimes;
  return JSON.stringify(s);
}

// ─────────────────────────────────────────────
// 현재 설명회 인덱스 조회 (Settings.currentSessionIdx)
// ─────────────────────────────────────────────
function _getCurrentSessionIdx() {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
    if (!sh) return 0;
    var rows = sh.getDataRange().getValues();
    for (var i = 0; i < rows.length; i++) {
      if (rows[i][0] && rows[i][0].toString().trim() === 'currentSessionIdx') {
        return parseInt(rows[i][1]) || 0;
      }
    }
    return 0;
  } catch(e) { return 0; }
}

// ─────────────────────────────────────────────
// 설명회 세션 변경 (staff.html에서 호출)
// newIdx: 0~_getSessionSchedule().length (length = 행사 종료)
// ─────────────────────────────────────────────
function setSessionIdx(password, newIdx) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var admins = ss.getSheetByName('AdminUsers').getDataRange().getValues();
  var authorized = false;
  for (var i = 1; i < admins.length; i++) {
    if (admins[i][2].toString().trim() === '전체관리' &&
        admins[i][3].toString().trim() === password.toString().trim()) {
      authorized = true; break;
    }
  }
  if (!authorized) throw new Error('비밀번호가 올바르지 않습니다.');
  var idx = Math.max(0, Math.min(parseInt(newIdx) || 0, _getSessionSchedule().length));
  var sh = ss.getSheetByName('Settings');
  if (!sh) throw new Error('Settings 시트가 없습니다.');
  var rows = sh.getDataRange().getValues();
  for (var j = 0; j < rows.length; j++) {
    if (rows[j][0] && rows[j][0].toString().trim() === 'currentSessionIdx') {
      sh.getRange(j+1, 2).setValue(idx);
      return idx;
    }
  }
  sh.appendRow(['currentSessionIdx', idx]);
  return idx;
}

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var SHEETS = [
    { name:'Settings', headers:null,
      init:function(sh){ if(sh.getLastRow()===0){
        var initData = [
          ['예약오픈여부',    'OPEN',                                    '', ''],
          ['행사명',          '융합 Universe : 6개 융합의 별을 연결하다','', ''],
          ['행사일시',        '2026. 5. 8.(금) 09:30 ~ 17:00',          '', ''],
          ['행사장소',        '한양종합기술원(HIT) 1층 양민용 커리어라운지','',''],
          ['currentSessionIdx', 0,                                       '', ''],
          ['설명회세션', '미래사회디자인',        '09:30', '09:55'],
          ['설명회세션', '융합의과학/융합의공학', '10:00', '10:25'],
          ['설명회세션', '인지융합과학',          '10:30', '10:55'],
          ['설명회세션', '혁신공학경영',          '11:00', '11:25'],
          ['설명회세션', '미래반도체공학',        '11:30', '11:55'],
          ['부스프로그램', '미래반도체공학',        '', ''],
          ['부스프로그램', '혁신공학경영',          '', ''],
          ['부스프로그램', '융합의과학/융합의공학', '', ''],
          ['부스프로그램', '미래사회디자인',        '', ''],
          ['부스프로그램', '인지융합과학',          '', ''],
          ['부스프로그램', '라이프디자인센터',      '', ''],
        ];
        sh.getRange(1, 1, initData.length, 4).setValues(initData);
        sh.getRange(1, 1, initData.length, 1).setFontWeight('bold');
      }}
    },
    { name:'AdminUsers',        headers:['이름','호칭','담당프로그램','비밀번호'] },
    { name:'BoothReservations', headers:['이름','학번','학과','이메일','연락처','프로그램','시간','문의내용','서명','상태','코멘트','예약일시'] },
    { name:'SessionPreReg',     headers:['이름','학번','학과','연락처','이메일','설명회명','등록일시'] },
    { name:'CheckIns',          headers:['체크인시각','학번','이름','학과','연락처','이메일','참석유형','참석설명회'] },
    { name:'GiftReceipts',      headers:['수령시각','학번','이름','학과','방문부스','스탬프수','가산점과목','서명','검증코드'] },
    { name:'BlockedSlots',      headers:['프로그램','시간','차단여부'] },
  ];
  SHEETS.forEach(function(def) {
    var sh = ss.getSheetByName(def.name);
    if (!sh) { sh = ss.insertSheet(def.name); Logger.log('생성: '+def.name); }
    else { Logger.log('확인: '+def.name); }
    if (def.headers && sh.getLastRow()===0) {
      sh.getRange(1,1,1,def.headers.length).setValues([def.headers]);
      sh.getRange(1,1,1,def.headers.length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    if (def.init) def.init(sh);
  });
  Logger.log('setupSheets 완료 — '+SHEETS.length+'개 시트 확인/생성');
}

// ─────────────────────────────────────────────
// 라우팅
// ─────────────────────────────────────────────
function doGet(e) {
  var page = e.parameter.page || 'index';

  // index.html이 GAS_CONFIG_URL(?page=config)로 설정값 요청
  if (page === 'config') {
    var cfg = _getSettings();
    var sessionTimes = {};
    cfg.sessions.forEach(function(s) { sessionTimes[s.name] = s.start + '~' + s.end; });

    var sessionCounts = {};
    cfg.sessions.forEach(function(s) { sessionCounts[s.name] = 0; });
    try {
      var ss2 = SpreadsheetApp.getActiveSpreadsheet();
      var sessSheet = ss2.getSheetByName('SessionPreReg');
      if (sessSheet && sessSheet.getLastRow() > 1) {
        sessSheet.getDataRange().getValues().slice(1).forEach(function(r) {
          var prog = r[5] ? r[5].toString().trim() : '';
          if (sessionCounts[prog] !== undefined) sessionCounts[prog]++;
        });
      }
    } catch(e2) {}

    var payload = {
      isOpen:          cfg.isOpen,
      eventName:       cfg.eventName,
      eventDate:       cfg.eventDate,
      eventPlace:      cfg.eventPlace,
      sessions:        cfg.sessions,
      boothPrograms:   cfg.boothPrograms,
      sessionCapacity: SESSION_CAPACITY,
      sessionTimes:    sessionTimes,
      sessionCounts:   sessionCounts
    };
    return ContentService.createTextOutput(JSON.stringify(payload))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (['index','admin','staff'].indexOf(page) === -1) {
    return HtmlService.createHtmlOutput('<h3>잘못된 접근입니다.</h3>');
  }
  return HtmlService.createTemplateFromFile(page)
    .evaluate()
    .setTitle('한양YK인터칼리지 융합전공 소개행사')
    .addMetaTag('viewport','width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}


// ─────────────────────────────────────────────
// 테스트: GAS 편집기에서 직접 실행 → 메일 발송 + doPost 시뮬레이션 확인
// ─────────────────────────────────────────────
function testMailAndPost() {
  var me = Session.getEffectiveUser().getEmail();
  // 1) 메일 발송 테스트
  try {
    GmailApp.sendEmail(me, '[테스트] GAS 메일 발송 확인', '이 메일이 오면 GmailApp 정상 작동 중입니다.');
    Logger.log('메일 발송 성공: ' + me);
  } catch(e) {
    Logger.log('메일 발송 실패: ' + e.message);
  }
  // 2) doPost 시뮬레이션 (더미 페이로드)
  try {
    var dummy = {
      action: 'change',
      name: '테스트', studentId: '0000000000', dept: '한양인터칼리지학부',
      email: me,
      phone: '010-0000-0000', sessions: [], booths: []
    };
    var fakeE = { postData: { contents: JSON.stringify(dummy) } };
    var res = doPost(fakeE);
    Logger.log('doPost 시뮬레이션 결과: ' + res.getContent());
  } catch(e) {
    Logger.log('doPost 시뮬레이션 실패: ' + e.message);
  }
}

// ─────────────────────────────────────────────
// doPost: index.html → GAS 예약 데이터 수신
// Sheets 기록 + 메일 발송
// ─────────────────────────────────────────────
function doPost(e) {
  var result = { ok: false };
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('postData 없음 — GET 요청이거나 body 누락');
    }
    var payload = JSON.parse(e.postData.contents);
    var action  = payload.action; // 'reserve' | 'cancel' | 'change'
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var now     = new Date();

    if (action === 'reserve') {
      var name    = payload.name    || '';
      var sid     = payload.studentId || '';
      var dept    = payload.dept    || '';
      var email   = payload.email   || ''; // 이미 @hanyang.ac.kr 포함
      var phone   = payload.phone   || '';
      var sessions  = payload.sessions  || []; // ['설명회명', ...]
      var booths    = payload.booths    || []; // [{program, time, memo}]

      // ── SessionPreReg 시트에 기록 ──
      var sessSheet = _getOrCreateSheet(ss, 'SessionPreReg', ['이름','학번','학과','연락처','이메일','설명회명','등록일시']);
      var existSessKeys = {};
      if (sessSheet.getLastRow() > 1) {
        sessSheet.getDataRange().getValues().slice(1).forEach(function(r) {
          existSessKeys[r[1].toString().trim() + '|' + r[5].toString().trim()] = true;
        });
      }
      sessions.forEach(function(sessName) {
        var key = sid + '|' + sessName;
        if (!existSessKeys[key]) {
          sessSheet.appendRow([name, sid, dept, phone, email, sessName, now]);
        }
      });

      // ── BoothReservations 시트에 기록 ──
      var boothSheet = _getOrCreateSheet(ss, 'BoothReservations', ['이름','학번','학과','이메일','연락처','프로그램','시간','문의내용','서명','상태','코멘트','예약일시']);
      var existBoothKeys = {};
      if (boothSheet.getLastRow() > 1) {
        boothSheet.getDataRange().getValues().slice(1).forEach(function(r) {
          var st = r[9] ? r[9].toString().trim() : '';
          if (st !== '취소') existBoothKeys[r[1].toString().trim() + '|' + r[5].toString().trim()] = true;
        });
      }
      booths.forEach(function(b) {
        var key = sid + '|' + b.program;
        if (!existBoothKeys[key]) {
          boothSheet.appendRow([name, sid, dept, email, phone, b.program, b.time, b.memo||'', '', '예약완료', '', now]);
        }
      });

      // ── 예약확인 메일 발송 ──
      if (email) {
        var hasLdc = booths.some(function(b){ return b.program === '라이프디자인센터'; });
        var boothsForMail = booths.map(function(b){ return {program: b.program, time: b.time}; });
        try { sendConfirmMail(email, name, sessions, boothsForMail, hasLdc); }
        catch(mailErr) { Logger.log('메일 오류(reserve): ' + mailErr.message); }
      }

    } else if (action === 'change') {
      var name    = payload.name    || '';
      var sid     = payload.studentId || '';
      var dept    = payload.dept    || '';
      var email   = payload.email   || '';
      var phone   = payload.phone   || '';
      var sessions  = payload.sessions  || [];
      var booths    = payload.booths    || [];

      var sessSheet = _getOrCreateSheet(ss, 'SessionPreReg', ['이름','학번','학과','연락처','이메일','설명회명','등록일시']);
      var boothSheet = _getOrCreateSheet(ss, 'BoothReservations', ['이름','학번','학과','이메일','연락처','프로그램','시간','문의내용','서명','상태','코멘트','예약일시']);

      // ── 기존 설명회 행 삭제 (해당 학번 전체) ──
      if (sessSheet.getLastRow() > 1) {
        var sessRows = sessSheet.getDataRange().getValues();
        for (var si = sessRows.length - 1; si >= 1; si--) {
          if (sessRows[si][1].toString().trim() === sid) sessSheet.deleteRow(si + 1);
        }
      }
      // ── 기존 부스 예약 → '취소' 처리 (해당 학번 전체) ──
      if (boothSheet.getLastRow() > 1) {
        var boothRows = boothSheet.getDataRange().getValues();
        for (var bi = 1; bi < boothRows.length; bi++) {
          if (boothRows[bi][1].toString().trim() !== sid) continue;
          var bSt = boothRows[bi][9] ? boothRows[bi][9].toString().trim() : '';
          if (bSt !== '취소' && bSt !== '상담취소' && bSt !== '상담완료') {
            boothSheet.getRange(bi + 1, 10).setValue('취소');
          }
        }
      }
      // ── 새 설명회 추가 ──
      sessions.forEach(function(sessName) {
        sessSheet.appendRow([name, sid, dept, phone, email, sessName, now]);
      });
      // ── 새 부스 추가 ──
      booths.forEach(function(b) {
        boothSheet.appendRow([name, sid, dept, email, phone, b.program, b.time, b.memo||'', '', '예약완료', '', now]);
      });

      // ── 변경 확인 메일 발송 ──
      if (email) {
        var hasLdc = booths.some(function(b){ return b.program === '라이프디자인센터'; });
        var boothsForMail = booths.map(function(b){ return {program: b.program, time: b.time}; });
        try { sendChangeMail(email, name, sessions, boothsForMail, hasLdc); }
        catch(mailErr) { Logger.log('메일 오류(change): ' + mailErr.message); }
      }

    } else if (action === 'cancel') {
      var sid2     = payload.studentId || '';
      var email2   = payload.email     || '';
      var name2    = payload.name      || '';
      var cBooths  = payload.booths    || [];
      var cSessions= payload.sessions  || [];

      // BoothReservations → 상태 '취소' 처리
      var bSheet = ss.getSheetByName('BoothReservations');
      if (bSheet && bSheet.getLastRow() > 1) {
        var bRows = bSheet.getDataRange().getValues();
        for (var i = 1; i < bRows.length; i++) {
          if (bRows[i][1].toString().trim() !== sid2) continue;
          var bProg = bRows[i][5].toString().trim();
          if (cBooths.some(function(b){ return b.program === bProg; })) {
            bSheet.getRange(i+1, 10).setValue('취소');
          }
        }
      }
      // SessionPreReg → 행 삭제
      var sSheet = ss.getSheetByName('SessionPreReg');
      if (sSheet && sSheet.getLastRow() > 1) {
        var sRows = sSheet.getDataRange().getValues();
        for (var j = sRows.length - 1; j >= 1; j--) {
          if (sRows[j][1].toString().trim() !== sid2) continue;
          var sName = sRows[j][5].toString().trim();
          if (cSessions.indexOf(sName) !== -1) sSheet.deleteRow(j+1);
        }
      }
      // 취소 메일
      if (email2) {
        try { sendCancelMail(email2, name2, cBooths, cSessions); } catch(e2) {}
      }
    }

    result.ok = true;
  } catch(err) {
    result.error = err.message;
    Logger.log('doPost 오류: ' + err.message);
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────
// 초기 데이터 (설명회 선착순 현황 포함)
// ─────────────────────────────────────────────
function getInitialData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = _getSettings();

  // 설명회별 예약 수 집계
  var sessionCounts = {};
  cfg.sessions.forEach(function(s){ sessionCounts[s.name] = 0; });

  try {
    var sessSheet = ss.getSheetByName('SessionPreReg');
    if (sessSheet && sessSheet.getLastRow() > 1) {
      var rows = sessSheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        var prog = rows[i][5] ? rows[i][5].toString().trim() : '';
        if (sessionCounts[prog] !== undefined) sessionCounts[prog]++;
      }
    }
  } catch(e) {}

  return {
    isOpen:          cfg.isOpen,
    programs:        cfg.boothPrograms,
    sessions:        cfg.sessions,
    sessionCounts:   sessionCounts,
    sessionCapacity: SESSION_CAPACITY,
    eventName:       cfg.eventName,
    eventDate:       cfg.eventDate,
    eventPlace:      cfg.eventPlace
  };
}

// ─────────────────────────────────────────────
// 신청 과목 목록 반환
// ─────────────────────────────────────────────
function getBonusSubjects() {
  return BONUS_SUBJECTS;
}

// ─────────────────────────────────────────────
// 설명회 예약 수 조회 (프로그램별)
// ─────────────────────────────────────────────
function getSessionReservationCount(program) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('SessionPreReg');
  if (!sheet || sheet.getLastRow() <= 1) return 0;
  var rows = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][5].toString().trim() === program) count++;
  }
  return count;
}

// ─────────────────────────────────────────────
// 설명회 현장 체크인
// CheckIns: 체크인시각|학번|이름|학과|연락처|이메일|참석유형|참석설명회
// ─────────────────────────────────────────────
// 당일방문 현황 조회 (checkin.html 시작 화면용)
// 반환: { walkCount: n, walkCapacity: n, remaining: n }
// ─────────────────────────────────────────────
function getCheckinPageData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var walkCount = 0, preCount = 0;
  var sessionCounts = {}; // { 세션명: 고유학번 수 }
  var ciSheet = ss.getSheetByName('CheckIns');
  if (ciSheet && ciSheet.getLastRow() > 1) {
    var ciRows = ciSheet.getDataRange().getValues();
    var seenWalk = {}, seenPre = {}, seenSess = {};
    for (var i = 1; i < ciRows.length; i++) {
      var sid   = ciRows[i][1] ? ciRows[i][1].toString().trim() : '';
      var type  = ciRows[i][6] ? ciRows[i][6].toString().trim() : '';
      var sess  = ciRows[i][7] ? ciRows[i][7].toString().trim() : '';
      if (!sid) continue;
      if (type === '당일방문' && !seenWalk[sid]) { seenWalk[sid] = true; walkCount++; }
      if (type === '사전예약' && !seenPre[sid])  { seenPre[sid]  = true; preCount++;  }
      if (sess) {
        if (!seenSess[sess]) seenSess[sess] = {};
        seenSess[sess][sid] = true;
      }
    }
    for (var sn in seenSess) sessionCounts[sn] = Object.keys(seenSess[sn]).length;
  }
  var isBlocked = false, currentSessionIdx = 0;
  var stSheet = ss.getSheetByName('Settings');
  if (stSheet) {
    var sRows = stSheet.getDataRange().getValues();
    for (var s = 0; s < sRows.length; s++) {
      var key = sRows[s][0] ? sRows[s][0].toString().trim() : '';
      if (key === 'walkInBlocked') isBlocked = sRows[s][1] && sRows[s][1].toString().trim() === 'TRUE';
      if (key === 'currentSessionIdx') currentSessionIdx = parseInt(sRows[s][1]) || 0;
    }
  }
  return {
    walkCount:         walkCount,
    preCount:          preCount,
    totalCount:        walkCount + preCount,
    sessionCounts:     sessionCounts,
    walkCapacity:      WALK_IN_CAPACITY,
    isFull:            walkCount >= WALK_IN_CAPACITY || isBlocked,
    isBlocked:         isBlocked,
    currentSessionIdx: currentSessionIdx
  };
}

// ─────────────────────────────────────────────
// 당일방문 수동 마감/해제 (admin용)
// ─────────────────────────────────────────────
// 모드 변경 비밀번호 검증 (AdminUsers 시트의 어떤 계정이든 일치하면 허용)
function verifyAdminPassword(password) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var admins = ss.getSheetByName('AdminUsers').getDataRange().getValues();
    for (var i = 1; i < admins.length; i++) {
      if (admins[i][3] && admins[i][3].toString().trim() === password.toString().trim()) {
        return true;
      }
    }
    return false;
  } catch(e) { return false; }
}

function setWalkInBlock(password, block) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var admins = ss.getSheetByName('AdminUsers').getDataRange().getValues();
  var authorized = false;
  for (var i = 1; i < admins.length; i++) {
    if (admins[i][2].toString().trim() === '전체관리' &&
        admins[i][3].toString().trim() === password.toString().trim()) {
      authorized = true; break;
    }
  }
  if (!authorized) throw new Error('권한이 없습니다.');
  var sheet = ss.getSheetByName('Settings');
  if (!sheet) throw new Error('Settings 시트가 없습니다.');
  var rows = sheet.getDataRange().getValues();
  for (var j = 0; j < rows.length; j++) {
    if (rows[j][0] && rows[j][0].toString().trim() === 'walkInBlocked') {
      sheet.getRange(j+1, 2).setValue(block ? 'TRUE' : 'FALSE');
      return block ? '당일방문이 수동 마감되었습니다.' : '당일방문이 재개되었습니다.';
    }
  }
  sheet.appendRow(['walkInBlocked', block ? 'TRUE' : 'FALSE']);
  return block ? '당일방문이 수동 마감되었습니다.' : '당일방문이 재개되었습니다.';
}

// ─────────────────────────────────────────────
// 설명회 세션 정의 (Settings 시트에서 동적 로드)
// ─────────────────────────────────────────────
function _getSessionSchedule() {
  return _getSettings().sessions;
}

// 현재 설명회 인덱스 기준으로 체크인 가능한 세션 반환
// - 스태프가 직접 세션을 진행시킴 (시간 기반 X)
// - currentSessionIdx >= _getSessionSchedule().length 이면 null (모든 세션 종료)
function _getCurrentCheckinableSession() {
  var schedule = _getSessionSchedule();
  var idx = _getCurrentSessionIdx();
  if (idx >= schedule.length) return null;
  return { session: schedule[idx], idx: idx, phase: 'open' };
}

// ─────────────────────────────────────────────
// 설명회 현장 체크인
// CheckIns: 체크인시각|학번|이름|학과|연락처|이메일|참석유형|참석설명회 (참석설명회당 1행)
// ─────────────────────────────────────────────
function checkInStudent(data) {
  // data: { studentId, name, dept, phone, type('pre'|'walk'), programs[] }
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var sid = data.studentId.toString().trim();
  if (!/^\d{10}$/.test(sid)) throw new Error('학번은 숫자 10자리여야 합니다.');

  // ── 시간 제한 체크: 모든 설명회가 종료된 경우만 차단 ──
  var checkinableSession = _getCurrentCheckinableSession();
  if (!checkinableSession) {
    throw new Error('모든 설명회가 종료되었습니다. 체크인이 불가합니다.');
  }

  // CheckIns 시트 준비
  var sheet = ss.getSheetByName('CheckIns');
  if (!sheet) {
    sheet = ss.insertSheet('CheckIns');
    sheet.getRange(1,1,1,8).setValues([['체크인시각','학번','이름','학과','연락처','이메일','참석유형','참석설명회']]);
    sheet.getRange(1,1,1,8).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var rows = sheet.getLastRow() > 1 ? sheet.getDataRange().getValues() : [];

    // 이미 체크인 여부 확인
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][1].toString().trim() === sid) {
        return { alreadyCheckedIn: true, checkedAt: rows[i][0].toString() };
      }
    }

    // 수동 마감 체크 (당일방문 + 사전예약 모두 차단)
    var stSheet = ss.getSheetByName('Settings');
    if (stSheet) {
      var stRows = stSheet.getDataRange().getValues();
      for (var st = 0; st < stRows.length; st++) {
        if (stRows[st][0] && stRows[st][0].toString().trim() === 'walkInBlocked' &&
            stRows[st][1] && stRows[st][1].toString().trim() === 'TRUE') {
          throw new Error('체크인이 마감되었습니다.');
        }
      }
    }

    // 당일방문 선착순 체크
    if (data.type === 'walk') {
      var walkCount = 0, seen2 = {};
      for (var j = 1; j < rows.length; j++) {
        var jSid = rows[j][1] ? rows[j][1].toString().trim() : '';
        if (rows[j][6] && rows[j][6].toString().trim() === '당일방문' && jSid && !seen2[jSid]) {
          seen2[jSid] = true; walkCount++;
        }
      }
      if (walkCount >= WALK_IN_CAPACITY) {
        throw new Error('당일방문 체크인이 마감되었습니다. (최대 ' + WALK_IN_CAPACITY + '명)');
      }
    }

    var now   = new Date();
    var name  = data.name.toString().trim();
    var dept  = data.dept.toString().trim();
    var phone = data.phone ? data.phone.toString().trim() : '';
    var email = data.email ? data.email.toString().trim() : '';
    var type  = data.type === 'pre' ? '사전예약' : '당일방문';

    // CheckIns에 참석설명회별 1행씩 저장
    var sessionsToRecord = [];
    if (data.type === 'pre') {
      // 사전예약자도 체크인 시점 세션부터 나머지 전체 기록 (당일방문과 동일)
      var curIdx3 = checkinableSession.idx;
      var _sched = _getSessionSchedule();
      for (var si = curIdx3; si < _sched.length; si++) {
        sessionsToRecord.push(_sched[si].name);
      }
    } else {
      var _sched2 = _getSessionSchedule();
      var curIdxW = checkinableSession.idx;
      for (var si = curIdxW; si < _sched2.length; si++) {
        sessionsToRecord.push(_sched2[si].name);
      }
    }

    if (sessionsToRecord.length === 0) {
      sheet.appendRow([now, sid, name, dept, phone, email, type, '']);
    } else {
      sessionsToRecord.forEach(function(sessName) {
        sheet.appendRow([now, sid, name, dept, phone, email, type, sessName]);
      });
    }

    return { alreadyCheckedIn: false, checkedAt: now.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────
// 사전예약자 체크인 화면용 복합 조회
// 체크인 여부 + 사전예약 내역을 1회 왕복으로 반환
// ─────────────────────────────────────────────
function getPreCheckinData(studentId) {
  var sid = studentId.toString().trim();
  if (!/^\d{10}$/.test(sid)) throw new Error('학번은 숫자 10자리여야 합니다.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var checkIn = null;
  var ciSheet = ss.getSheetByName('CheckIns');
  if (ciSheet && ciSheet.getLastRow() > 1) {
    var ciRows = ciSheet.getDataRange().getValues();
    for (var i = 1; i < ciRows.length; i++) {
      if (ciRows[i][1].toString().trim() === sid) {
        checkIn = {
          checkedAt: ciRows[i][0] ? ciRows[i][0].toString() : '',
          name:  ciRows[i][2] ? ciRows[i][2].toString().trim() : '',
          dept:  ciRows[i][3] ? ciRows[i][3].toString().trim() : '',
          phone: ciRows[i][4] ? ciRows[i][4].toString().trim() : '',
          email: ciRows[i][5] ? ciRows[i][5].toString().trim() : '',
          type:  ciRows[i][6] ? ciRows[i][6].toString().trim() : ''
        };
        break;
      }
    }
  }

  var preData = null;
  var prSheet = ss.getSheetByName('SessionPreReg');
  if (prSheet && prSheet.getLastRow() > 1) {
    var prRows = prSheet.getDataRange().getValues();
    var sessions = [], name = '', dept = '', phone = '', email = '';
    for (var j = 1; j < prRows.length; j++) {
      if (prRows[j][1].toString().trim() === sid) {
        if (!name)  name  = prRows[j][0].toString().trim();
        if (!dept)  dept  = prRows[j][2].toString().trim();
        if (!phone) phone = prRows[j][3].toString().trim();
        if (!email) email = prRows[j][4] ? prRows[j][4].toString().trim() : '';
        sessions.push(prRows[j][5].toString().trim());
      }
    }
    if (sessions.length) preData = { name: name, dept: dept, phone: phone, email: email, sessions: sessions };
  }

  return { checkIn: checkIn, preData: preData };
}

// ─────────────────────────────────────────────
// 체크인 여부 조회
// ─────────────────────────────────────────────
function getCheckInStatus(studentId) {
  var sid = studentId.toString().trim();
  if (!/^\d{10}$/.test(sid)) throw new Error('학번은 숫자 10자리여야 합니다.');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CheckIns');
  if (!sheet || sheet.getLastRow() <= 1) return null;
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][1].toString().trim() === sid) {
      return {
        checkedAt: rows[i][0] ? rows[i][0].toString() : '',
        name:      rows[i][2] ? rows[i][2].toString().trim() : '',
        dept:      rows[i][3] ? rows[i][3].toString().trim() : '',
        phone:     rows[i][4] ? rows[i][4].toString().trim() : '',
        email:     rows[i][5] ? rows[i][5].toString().trim() : '',
        type:      rows[i][6] ? rows[i][6].toString().trim() : ''
      };
    }
  }
  return null;
}



// ─────────────────────────────────────────────
// 기념품 수령 명단 조회 (admin용)
// ─────────────────────────────────────────────
function getGiftList(password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var admins = ss.getSheetByName('AdminUsers').getDataRange().getValues();
  var authorized = false;
  for (var i = 1; i < admins.length; i++) {
    if (admins[i][2].toString().trim() === '전체관리' &&
        admins[i][3].toString().trim() === password.toString().trim()) {
      authorized = true; break;
    }
  }
  if (!authorized) throw new Error('권한이 없습니다.');

  var sheet = ss.getSheetByName('GiftReceipts');
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var rows = sheet.getDataRange().getValues();
  // 체크인 학번 맵
  var ciMap = {};
  var ciSheet = ss.getSheetByName('CheckIns');
  if (ciSheet && ciSheet.getLastRow() > 1) {
    var ciRows = ciSheet.getDataRange().getValues();
    for (var ci = 1; ci < ciRows.length; ci++) {
      var ciSid = ciRows[ci][1] ? ciRows[ci][1].toString().trim() : '';
      if (ciSid) ciMap[ciSid] = true;
    }
  }

  // 1패스: 학번 기준으로 전체 집계 (부스당 1행이므로)
  var byStudent = {};
  for (var j = 1; j < rows.length; j++) {
    var sid2 = rows[j][1] ? rows[j][1].toString().trim() : '';
    if (!sid2) continue;
    if (!byStudent[sid2]) {
      byStudent[sid2] = {
        receivedAt:   rows[j][0] ? rows[j][0].toString() : '',
        name:         rows[j][2] ? rows[j][2].toString().trim() : '',
        dept:         rows[j][3] ? rows[j][3].toString().trim() : '',
        stampCount:   rows[j][5] ? Number(rows[j][5]) : 0,
        bonusSubject: rows[j][6] ? rows[j][6].toString().trim() : '',
        signature:    rows[j][7] ? rows[j][7].toString() : '',
        verifyCode:   rows[j][8] ? rows[j][8].toString().trim() : '',
        boothList:    []
      };
    }
    var booth2 = rows[j][4] ? rows[j][4].toString().trim() : '';
    if (booth2) byStudent[sid2].boothList.push(booth2);
  }

  // 2패스: 자격 검증 후 결과 생성
  var result = [];
  Object.keys(byStudent).forEach(function(sid2) {
    var s = byStudent[sid2];
    if (!ciMap[sid2] || s.stampCount < BOOTH_STAMP_REQUIRED) return; // 자격 미달 제외
    result.push({
      no:             0,
      receivedAt:     s.receivedAt,
      studentId:      sid2,
      name:           s.name,
      dept:           s.dept,
      isIntercollege: s.dept === INTERCOLLEGE_DEPT,
      boothList:      s.boothList,
      booths:         s.boothList.join(', '),
      stampCount:     s.stampCount,
      bonusSubject:   s.bonusSubject,
      signature:      s.signature,
      verifyCode:     s.verifyCode
    });
  });
  result.forEach(function(r, i) { r.no = i + 1; });
  return result;
}

// ─────────────────────────────────────────────
// 학번으로 사전예약 설명회 조회 (체크인 화면용)
// ─────────────────────────────────────────────
function getPreRegisteredSessions(studentId) {
  var sid = studentId.toString().trim();
  if (!/^\d{10}$/.test(sid)) throw new Error('학번은 숫자 10자리여야 합니다.');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SessionPreReg');
  if (!sheet || sheet.getLastRow() <= 1) return null;
  var rows = sheet.getDataRange().getValues();
  var sessions = [], name = '', dept = '', phone = '', email = '';
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][1].toString().trim() === sid) {
      if (!name)  name  = rows[i][0].toString().trim();
      if (!dept)  dept  = rows[i][2].toString().trim();
      if (!phone) phone = rows[i][3].toString().trim();
      if (!email) email = rows[i][4] ? rows[i][4].toString().trim() : '';
      sessions.push(rows[i][5].toString().trim());
    }
  }
  if (!sessions.length) return null;
  return { name: name, dept: dept, phone: phone, email: email, sessions: sessions };
}

// ─────────────────────────────────────────────
// 체크인 전체 현황 조회 (admin용)
// ─────────────────────────────────────────────
function getCheckInList(password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var admins = ss.getSheetByName('AdminUsers').getDataRange().getValues();
  var authorized = false;
  for (var i = 1; i < admins.length; i++) {
    if (admins[i][2].toString().trim() === '전체관리' &&
        admins[i][3].toString().trim() === password.toString().trim()) {
      authorized = true; break;
    }
  }
  if (!authorized) throw new Error('권한이 없습니다.');
  var sheet = ss.getSheetByName('CheckIns');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var rows = sheet.getDataRange().getValues();
  // 학번 기준 첫 번째 행만 (대표행), 참석설명회는 콤마 합산
  var seen = {}, result = [];
  for (var j = 1; j < rows.length; j++) {
    var sid = rows[j][1] ? rows[j][1].toString().trim() : '';
    var sess = rows[j][7] ? rows[j][7].toString().trim() : '';
    if (!seen[sid]) {
      seen[sid] = result.length;
      result.push({
        no:        0,
        checkedAt: rows[j][0] ? rows[j][0].toString() : '',
        studentId: sid,
        name:      rows[j][2] ? rows[j][2].toString().trim() : '',
        dept:      rows[j][3] ? rows[j][3].toString().trim() : '',
        type:      rows[j][6] ? rows[j][6].toString().trim() : '',
        sessions:  sess ? [sess] : []
      });
    } else if (sess) {
      result[seen[sid]].sessions.push(sess);
    }
  }
  result.forEach(function(r, i) {
    r.no = i + 1;
    r.programs = r.sessions.join(', ');
  });
  return result;
}

// ─────────────────────────────────────────────
// 아래는 기존 V7 함수 전체 유지
// ─────────────────────────────────────────────

function getAvailableSlots(programName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bookedRanges = [];
  try {
    var resSheet = ss.getSheetByName('BoothReservations');
    if (resSheet && resSheet.getLastRow() > 1) {
      var allRows = resSheet.getDataRange().getValues();
      for (var i = 1; i < allRows.length; i++) {
        var r = allRows[i];
        var st = r[9] ? r[9].toString().trim() : '';
        if (st !== '' && st !== '취소' && st !== '상담취소' && st !== '상담완료') {
          var bProg = r[5] ? r[5].toString().trim() : '';
          var bT    = toTimeStr(r[6]);
          if (bProg && bT) {
            var bDur = slotDuration(bProg);
            var bS   = toMin(bT);
            bookedRanges.push({ prog: bProg, startMin: bS, endMin: bS + bDur });
          }
        }
      }
    }
  } catch(e) {}

  var blockedManual = new Set();
  try {
    var blkSheet = ss.getSheetByName('BlockedSlots');
    if (blkSheet && blkSheet.getLastRow() > 1) {
      var blkRows = blkSheet.getDataRange().getValues();
      for (var b = 1; b < blkRows.length; b++) {
        var bProg2  = blkRows[b][0] ? blkRows[b][0].toString().trim() : '';
        var bTime   = toTimeStr(blkRows[b][1]) || blkRows[b][1].toString().trim();
        var bActive = blkRows[b][2];
        var isBlocked = (bActive === true || bActive.toString().trim().toUpperCase() === 'TRUE'
                      || bActive.toString().trim() === 'Y' || bActive.toString().trim() === '차단');
        if (bProg2 === programName && bTime && isBlocked) blockedManual.add(bTime);
      }
    }
  } catch(e) {}

  var slotInterval = 15;
  var slots = [];
  var ranges = [{ start:'14:00', end:'17:00' }];
  ranges.forEach(function(range) {
    var startParts = range.start.split(':'), endParts = range.end.split(':');
    var startMin = parseInt(startParts[0])*60 + parseInt(startParts[1]);
    var endMin   = parseInt(endParts[0])*60   + parseInt(endParts[1]);
    for (var m = startMin; m < endMin; m += slotInterval) {
      var hh = String(Math.floor(m/60)).padStart(2,'0');
      var mm = String(m%60).padStart(2,'0');
      slots.push(hh+':'+mm);
    }
  });

  var SESSION_BLOCK_MAP = {
    '미래반도체공학':        ['09:30','09:45'],
    '혁신공학경영':          ['10:00','10:15'],
    '융합의과학/융합의공학': ['10:30','10:45'],
    '미래사회디자인':        ['11:00','11:15'],
    '인지융합과학':          ['11:30','11:45'],
    '라이프디자인센터':      []
  };
  var sessBlock = SESSION_BLOCK_MAP[programName] || [];
  var myDur = slotInterval;
  return slots
    .filter(function(t) {
      if (sessBlock.indexOf(t) !== -1) return false;
      if (blockedManual.has(t)) return false;
      var tS = toMin(t), tE = tS + myDur;
      for (var bi = 0; bi < bookedRanges.length; bi++) {
        var br = bookedRanges[bi];
        if (br.prog === programName && tS < br.endMin && br.startMin < tE) return false;
      }
      return true;
    })
    .map(function(t) {
      return { value: t, label: t + ' ~ ' + addMinutes(t, slotInterval) };
    });
}

function getAvailableSlotsForEdit(programName, excludeRowIdx) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resSheet = ss.getSheetByName('BoothReservations');
  var allRows  = resSheet.getDataRange().getValues();
  var bookedRanges = [];
  for (var i = 1; i < allRows.length; i++) {
    var r      = allRows[i];
    var status = r[9] ? r[9].toString().trim() : '';
    if (status === '취소' || status === '상담취소') continue;
    if (excludeRowIdx && (i + 1) === excludeRowIdx) continue;
    var bProg = r[5] ? r[5].toString().trim() : '';
    var bT    = toTimeStr(r[6]);
    if (!bT || !bProg) continue;
    bookedRanges.push({ prog: bProg, startMin: toMin(bT), endMin: toMin(bT) + 15 });
  }
  var blockedManual = new Set();
  try {
    var blkSheet2 = ss.getSheetByName('BlockedSlots');
    if (blkSheet2 && blkSheet2.getLastRow() > 1) {
      var blkRows = blkSheet2.getDataRange().getValues();
      for (var b = 1; b < blkRows.length; b++) {
        if (blkRows[b][0].toString().trim() === programName && blkRows[b][2].toString().trim().toUpperCase() === 'TRUE') {
          blockedManual.add(toTimeStr(blkRows[b][1]) || blkRows[b][1].toString().trim());
        }
      }
    }
  } catch(e) {}
  var SESSION_BLOCK_MAP = {
    '미래반도체공학':        ['09:30','09:45'],
    '혁신공학경영':          ['10:00','10:15'],
    '융합의과학/융합의공학': ['10:30','10:45'],
    '미래사회디자인':        ['11:00','11:15'],
    '인지융합과학':          ['11:30','11:45'],
    '라이프디자인센터':      []
  };
  var sessBlock = SESSION_BLOCK_MAP[programName] || [];
  var allSlots = [];
  var sm = 14*60, em = 17*60;
  for (var m = sm; m < em; m += 15) {
    allSlots.push(String(Math.floor(m/60)).padStart(2,'0') + ':' + String(m%60).padStart(2,'0'));
  }
  return allSlots
    .filter(function(t) {
      if (sessBlock.indexOf(t) !== -1) return false;
      if (blockedManual.has(t)) return false;
      var tS = toMin(t), tE = tS + 15;
      for (var bi = 0; bi < bookedRanges.length; bi++) {
        var br = bookedRanges[bi];
        if (br.prog === programName && tS < br.endMin && br.startMin < tE) return false;
      }
      return true;
    })
    .map(function(t) {
      return { value: t, label: t + ' ~ ' + addMinutes(t, 15) };
    });
}

function submitReservation(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var isOpen = ss.getSheetByName('Settings').getRange('B1').getValue().toString().trim().toUpperCase() === 'OPEN';
  if (!isOpen) throw new Error('현재 예약이 마감되었습니다.');
  var sid = data.studentId.toString().trim();
  if (!/^\d{10}$/.test(sid)) throw new Error('학번은 숫자 10자리여야 합니다.');

  if (data.reservations) {
    for (var vi = 0; vi < data.reservations.length; vi++) {
      if (data.reservations[vi].program && data.reservations[vi].program.toString().trim() === '라이프디자인센터') {
        var deptVal = data.dept ? data.dept.toString().trim() : '';
        if (deptVal !== '한양인터칼리지학부') {
          throw new Error('라이프디자인센터 부스는 한양인터칼리지학부 재학생만 예약 가능합니다.');
        }
      }
    }
  }

  var sheet = ss.getSheetByName('BoothReservations');
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    if (sheet.getLastRow() > 1) {
      var preCheck = sheet.getDataRange().getValues();
      for (var pc = 1; pc < preCheck.length; pc++) {
        var pcSid  = preCheck[pc][1] ? preCheck[pc][1].toString().trim() : '';
        var pcSt   = preCheck[pc][9] ? preCheck[pc][9].toString().trim() : '';
        var pcProg = preCheck[pc][5] ? preCheck[pc][5].toString().trim() : '';
        if (pcSid === sid && pcSt !== '' && pcSt !== '취소' && pcSt !== '상담취소') {
          for (var pci = 0; pci < data.reservations.length; pci++) {
            if (data.reservations[pci].program.toString().trim() === pcProg) {
              throw new Error('[' + pcProg + '] 이미 예약하셨습니다. 예약 변경은 예약 조회·취소 탭에서 이용해 주세요.');
            }
          }
        }
      }
    }
    var allRows = sheet.getDataRange().getValues();
    var active = [];
    for (var i = 1; i < allRows.length; i++) {
      var r = allRows[i];
      var st = r[9] ? r[9].toString().trim() : '';
      if (st !== '' && st !== '취소' && st !== '상담취소') {
        active.push({ studentId: r[1]?r[1].toString().trim():'', program:r[5]?r[5].toString().trim():'', time:toTimeStr(r[6]), duration:15, status:st });
      }
    }
    var reservations = data.reservations;
    var toAdd = [];
    for (var ri = 0; ri < reservations.length; ri++) {
      var item = reservations[ri];
      var prog = item.program.toString().trim();
      var timeVal = item.time.toString().trim();
      if (timeVal.indexOf('~') !== -1) timeVal = timeVal.split('~')[0].trim();
      for (var j = 0; j < active.length; j++) {
        if (active[j].studentId === sid && active[j].program === prog) throw new Error('[' + prog + '] 이미 해당 프로그램에 예약하셨습니다.');
      }
      for (var jj = 0; jj < toAdd.length; jj++) {
        if (toAdd[jj].program === prog) throw new Error('[' + prog + '] 동일 프로그램을 중복 선택했습니다.');
      }
      var newDur = slotDuration(prog);
      for (var k = 0; k < active.length; k++) {
        if (active[k].studentId === sid && timesOverlap(timeVal, newDur, active[k].time, active[k].duration)) {
          throw new Error('[' + prog + '] 시간이 기존 예약과 겹칩니다.');
        }
      }
      for (var kk = 0; kk < toAdd.length; kk++) {
        if (timesOverlap(timeVal, newDur, toAdd[kk].time, slotDuration(toAdd[kk].program))) {
          throw new Error('선택한 시간대가 서로 겹칩니다.');
        }
      }
      for (var l = 0; l < active.length; l++) {
        if (active[l].program === prog && active[l].time === timeVal && active[l].status !== '상담취소' && active[l].status !== '상담완료') {
          throw new Error('[' + prog + '] ' + timeVal + ' 시간은 방금 마감되었습니다.');
        }
      }
      toAdd.push({ program: prog, time: timeVal, duration: newDur, memo: item.memo ? item.memo.toString().trim() : '' });
    }
    var now      = new Date();
    var phoneVal = data.phone ? data.phone.toString().trim() : '';
    var baseMemo = data.memo ? data.memo.toString().trim() : '';
    var ldcTestStr = '\n[진로적성검사: 진로적성검사 응시 희망]';
    var memoWithoutLdc = baseMemo.replace(ldcTestStr, '').trim();
    for (var ai = 0; ai < toAdd.length; ai++) {
      var isLDC = (toAdd[ai].program === '라이프디자인센터');
      var rowMemo = isLDC ? baseMemo : memoWithoutLdc;
      sheet.appendRow([data.name.toString().trim(), sid, data.dept.toString().trim(), data.email.toString().trim()+'@hanyang.ac.kr', phoneVal, toAdd[ai].program, toAdd[ai].time, rowMemo, data.signature.toString(), '예약완료', '', now]);
    }
    if (data.sessions && data.sessions.length > 0) {
      var sessSheet = ss.getSheetByName('SessionPreReg');
      if (!sessSheet) {
        sessSheet = ss.insertSheet('SessionPreReg');
        sessSheet.getRange(1,1,1,7).setValues([['이름','학번','학과','연락처','이메일','설명회명','등록일시']]);
      }
      var alreadySet2 = new Set();
      if (sessSheet.getLastRow() > 1) {
        var sessRows2 = sessSheet.getDataRange().getValues();
        for (var si2 = 1; si2 < sessRows2.length; si2++) {
          if (sessRows2[si2][1].toString().trim() === sid) alreadySet2.add(sessRows2[si2][5].toString().trim());
        }
      }
      var toAddSess = data.sessions.filter(function(s){ return !alreadySet2.has(s.toString().trim()); });
      if (toAddSess.length > 0) _appendSessionRows(sessSheet, data.name.toString().trim(), sid, data.dept.toString().trim(), phoneVal, data.email.toString().trim()+'@hanyang.ac.kr', toAddSess, now);
    }
    try {
      var toEmail = data.email.toString().trim() + '@hanyang.ac.kr';
      sendConfirmMail(toEmail, data.name.toString().trim(), data.sessions||[], toAdd.map(function(b){return{program:b.program,time:b.time};}), data.memo&&data.memo.toString().indexOf('진로적성검사')!==-1);
    } catch(mailErr) { Logger.log('확인 메일 발송 실패: ' + mailErr.message); }
    var sessionCount=(data.sessions&&data.sessions.length)?data.sessions.length:0; var boothCount=toAdd.length; var msgParts=[]; if(sessionCount>0) msgParts.push('설명회 '+sessionCount+'개'); if(boothCount>0) msgParts.push('부스 상담 '+boothCount+'개'); return msgParts.join(', ')+' 예약이 완료되었습니다.';
  } finally {
    lock.releaseLock();
  }
}

function getMyReservations(studentId) {
  if (!/^\d{10}$/.test(studentId.toString())) throw new Error('학번은 숫자 10자리여야 합니다.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = [];
  try {
    var resSheet = ss.getSheetByName('BoothReservations');
    if (resSheet && resSheet.getLastRow() > 1) {
      var resRows = resSheet.getDataRange().getValues();
      for (var i = 1; i < resRows.length; i++) {
        var r = resRows[i];
        var status = r[9] ? r[9].toString().trim() : '';
        if (r[1].toString().trim() === studentId.toString() && status !== '취소') {
          var prog = r[5] ? r[5].toString().trim() : '';
          var st   = toTimeStr(r[6]);
          result.push({ type:'booth', rowIdx:i+1, program:prog, time:st?(st+' ~ '+addMinutes(st,15)):'', rawTime:st, memo:r[7]?r[7].toString().trim():'', status:status, name:r[0]?r[0].toString().trim():'', dept:r[2]?r[2].toString().trim():'', email:r[3]?r[3].toString().replace('@hanyang.ac.kr','').trim():'', phone:r[4]?r[4].toString().trim():'' });
        }
      }
    }
  } catch(e) {}
  try {
    var sessSheet = ss.getSheetByName('SessionPreReg');
    if (sessSheet && sessSheet.getLastRow() > 1) {
      var sessRows = sessSheet.getDataRange().getValues();
      for (var j = 1; j < sessRows.length; j++) {
        var sr = sessRows[j];
        if (sr[1].toString().trim() === studentId.toString()) {
          result.push({ type:'session', rowIdx:j+1, program:'설명회 참여', time:sr[5]?sr[5].toString().trim():'', status:'등록완료', name:sr[0]?sr[0].toString().trim():'', dept:sr[2]?sr[2].toString().trim():'', phone:sr[3]?sr[3].toString().trim():'', email:sr[4]?sr[4].toString().replace('@hanyang.ac.kr','').trim():'' });
        }
      }
    }
  } catch(e) {}
  return result;
}

function cancelReservation(studentId, rowIdx) {
  if (!/^\d{10}$/.test(studentId.toString())) throw new Error('학번은 숫자 10자리여야 합니다.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('BoothReservations');
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    // rowIdx는 힌트로만 사용. 실제로는 학번+프로그램+시간으로 행을 재탐색하여
    // 앞선 삭제로 인한 행 번호 밀림 버그를 방지한다.
    var hintRow = sheet.getRange(rowIdx,1,1,12).getValues()[0];
    var hintProg = hintRow[5] ? hintRow[5].toString().trim() : '';
    var hintTime = toTimeStr(hintRow[6]);
    var allRows = sheet.getDataRange().getValues();
    var actualIdx = -1;
    for (var i = 1; i < allRows.length; i++) {
      var r = allRows[i];
      if (r[1].toString().trim() !== studentId.toString()) continue;
      if (r[9] && (r[9].toString().trim() === '취소' || r[9].toString().trim() === '상담완료')) continue;
      var rProg = r[5] ? r[5].toString().trim() : '';
      var rTime = toTimeStr(r[6]);
      if (rProg === hintProg && rTime === hintTime) { actualIdx = i + 1; break; }
    }
    if (actualIdx === -1) throw new Error('취소할 예약을 찾을 수 없습니다. 이미 취소되었을 수 있습니다.');
    var row = allRows[actualIdx - 1];
    if (row[9] && (row[9].toString().trim() === '취소' || row[9].toString().trim() === '상담완료')) {
      throw new Error('취소할 수 없는 예약입니다. (이미 취소되었거나 상담 완료)');
    }
    // 취소 메일 먼저 발송 (행 삭제 전)
    try {
      var cEmail = row[3]?row[3].toString().trim():'';
      var cName  = row[0]?row[0].toString().trim():'';
      var cProg  = row[5]?row[5].toString().trim():'';
      var cTime  = toTimeStr(row[6]);
      var cTimeLabel = cTime?(cTime+' ~ '+addMinutes(cTime,15)):'';
      if (cEmail) sendCancelMail(cEmail, cName, [{program:cProg, time:cTimeLabel}], []);
    } catch(me) {}
    // 행 삭제 (슬롯 즉시 해제)
    sheet.deleteRow(actualIdx);
    return '예약이 취소되었습니다.';
  } finally {
    lock.releaseLock();
  }
}

function cancelSessionReservation(studentId, rowIdx) {
  if (!/^\d{10}$/.test(studentId.toString())) throw new Error('학번은 숫자 10자리여야 합니다.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('SessionPreReg');
  if (!sheet) throw new Error('설명회 예약 내역을 찾을 수 없습니다.');
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    // rowIdx는 힌트로만 사용. 학번+설명회명으로 행을 재탐색하여 행 번호 밀림 버그를 방지한다.
    var hintRow = sheet.getRange(rowIdx,1,1,7).getValues()[0];
    var hintSess = hintRow[5] ? hintRow[5].toString().trim() : '';
    var allRows = sheet.getDataRange().getValues();
    var actualIdx = -1;
    for (var i = 1; i < allRows.length; i++) {
      var r = allRows[i];
      if (r[1].toString().trim() !== studentId.toString()) continue;
      if ((r[5] ? r[5].toString().trim() : '') === hintSess) { actualIdx = i + 1; break; }
    }
    if (actualIdx === -1) throw new Error('취소할 설명회 예약을 찾을 수 없습니다. 이미 취소되었을 수 있습니다.');
    var row = allRows[actualIdx - 1];
    var scEmail = row[4]?row[4].toString().trim():'';
    var scName  = row[0]?row[0].toString().trim():'';
    var scSess  = row[5]?row[5].toString().trim():'';
    sheet.deleteRow(actualIdx);
    try { if (scEmail) sendCancelMail(scEmail, scName, [], [scSess]); } catch(me) {}
    return '설명회 참여 등록이 취소되었습니다.';
  } finally {
    lock.releaseLock();
  }
}

function getAdminData(selectedProg, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var admins = ss.getSheetByName('AdminUsers').getDataRange().getValues();
  var profName='', title='교수', authorized=false;
  for (var i=1;i<admins.length;i++) {
    var rowProg=admins[i][2]?admins[i][2].toString().trim():'';
    var rowPw=admins[i][3]?admins[i][3].toString().trim():'';
    if (rowProg===selectedProg.trim()&&rowPw===password.toString().trim()) {
      profName=admins[i][0]?admins[i][0].toString().trim():'';
      title=admins[i][1]?admins[i][1].toString().trim():'교수';
      authorized=true; break;
    }
  }
  if (!authorized) throw new Error('선택한 프로그램의 비밀번호가 일치하지 않습니다.');
  var allRows=ss.getSheetByName('BoothReservations').getDataRange().getValues();
  var reservations=[];
  for (var j=1;j<allRows.length;j++) {
    var r=allRows[j];
    var prog=r[5]?r[5].toString().trim():'';
    var status=r[9]?r[9].toString().trim():'';
    if (prog!==selectedProg||status==='취소') continue;
    var st=toTimeStr(r[6]);
    reservations.push({ rowIdx:j+1, name:r[0]?r[0].toString().trim():'', id:r[1]?r[1].toString().trim():'', dept:r[2]?r[2].toString().trim():'', email:r[3]?r[3].toString().trim():'', time:st?(st+' ~ '+addMinutes(st,15)):'', memo:r[7]?r[7].toString().trim():'', status:status, comment:r[10]?r[10].toString().trim():'' });
  }
  reservations.sort(function(a,b){ var ta=a.time?a.time.substring(0,5):'', tb=b.time?b.time.substring(0,5):''; return ta<tb?-1:ta>tb?1:0; });
  return { profName:profName, title:title, programName:selectedProg, reservations:reservations };
}

function batchUpdateStatus(selectedProg, password, updates) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var admins=ss.getSheetByName('AdminUsers').getDataRange().getValues();
  var authorized=false;
  for (var i=1;i<admins.length;i++) {
    if (admins[i][2].toString().trim()===selectedProg.trim()&&admins[i][3].toString().trim()===password.toString().trim()) { authorized=true; break; }
  }
  if (!authorized) throw new Error('권한이 없습니다.');
  var sheet=ss.getSheetByName('BoothReservations');
  for (var u=0;u<updates.length;u++) {
    var upd=updates[u];
    var rowData=sheet.getRange(upd.rowIdx,1,1,11).getValues()[0];
    if (rowData[5].toString().trim()!==selectedProg) continue;
    if (upd.status) sheet.getRange(upd.rowIdx,10).setValue(upd.status);
    if (upd.comment!==null&&upd.comment!==undefined) sheet.getRange(upd.rowIdx,11).setValue(upd.comment);
  }
  return updates.length+'건 저장되었습니다.';
}

function getSlotBlockStatus(selectedProg, password) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var admins=ss.getSheetByName('AdminUsers').getDataRange().getValues();
  var authorized=false;
  for (var i=1;i<admins.length;i++) {
    if (admins[i][2].toString().trim()===selectedProg.trim()&&admins[i][3].toString().trim()===password.toString().trim()) { authorized=true; break; }
  }
  if (!authorized) throw new Error('권한이 없습니다.');
  var allSlots=[];
  for (var m=14*60;m<17*60;m+=15) allSlots.push(String(Math.floor(m/60)).padStart(2,'0')+':'+String(m%60).padStart(2,'0'));
  var SESS_BLOCK_ADMIN={'미래반도체공학':['09:30','09:45'],'혁신공학경영':['10:00','10:15'],'융합의과학/융합의공학':['10:30','10:45'],'미래사회디자인':['11:00','11:15'],'인지융합과학':['11:30','11:45'],'라이프디자인센터':[]};
  var sessBlockSet={};
  var sbList=SESS_BLOCK_ADMIN[selectedProg]||[];
  sbList.forEach(function(t){ sessBlockSet[t]=true; });
  var blockedSet=new Set();
  var blkSheet=ss.getSheetByName('BlockedSlots');
  if (blkSheet&&blkSheet.getLastRow()>1) {
    var blkRows=blkSheet.getDataRange().getValues();
    for (var b=1;b<blkRows.length;b++) {
      var bProg=blkRows[b][0]?blkRows[b][0].toString().trim():'';
      var bTime=toTimeStr(blkRows[b][1])||blkRows[b][1].toString().trim();
      var bAct=blkRows[b][2];
      var isBlocked=(bAct===true||bAct.toString().trim().toUpperCase()==='TRUE'||bAct.toString().trim()==='Y'||bAct.toString().trim()==='차단');
      if (bProg===selectedProg&&bTime&&isBlocked) blockedSet.add(bTime);
    }
  }
  var bookedSet=new Set();
  try {
    var resSheet=ss.getSheetByName('BoothReservations');
    if (resSheet&&resSheet.getLastRow()>1) {
      var resRows=resSheet.getDataRange().getValues();
      for (var rr=1;rr<resRows.length;rr++) {
        var rr_prog=resRows[rr][5]?resRows[rr][5].toString().trim():'';
        var rr_st=resRows[rr][9]?resRows[rr][9].toString().trim():'';
        if (rr_prog===selectedProg&&rr_st!==''&&rr_st!=='취소'&&rr_st!=='상담취소'&&rr_st!=='상담완료') {
          var rr_t=toTimeStr(resRows[rr][6]);
          if (rr_t) bookedSet.add(rr_t);
        }
      }
    }
  } catch(e) {}
  return allSlots.map(function(t){ return { time:t, label:t+' ~ '+addMinutes(t,15), booked:bookedSet.has(t), blocked:blockedSet.has(t), sessionBlock:!!sessBlockSet[t] }; });
}

function toggleSlotBlock(selectedProg, password, timeVal, block) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var admins=ss.getSheetByName('AdminUsers').getDataRange().getValues();
  var authorized=false;
  for (var i=1;i<admins.length;i++) {
    if (admins[i][2].toString().trim()===selectedProg.trim()&&admins[i][3].toString().trim()===password.toString().trim()) { authorized=true; break; }
  }
  if (!authorized) throw new Error('권한이 없습니다.');
  var blkSheet=ss.getSheetByName('BlockedSlots');
  if (!blkSheet) {
    blkSheet=ss.insertSheet('BlockedSlots');
    blkSheet.getRange(1,1,1,3).setValues([['프로그램','시간','차단여부']]);
  }
  var found=false;
  if (blkSheet.getLastRow()>1) {
    var rows=blkSheet.getDataRange().getValues();
    for (var r=1;r<rows.length;r++) {
      if (rows[r][0].toString().trim()===selectedProg&&(toTimeStr(rows[r][1])||rows[r][1].toString().trim())===timeVal) {
        blkSheet.getRange(r+1,3).setValue(block?'TRUE':'FALSE');
        found=true; break;
      }
    }
  }
  if (!found) blkSheet.appendRow([selectedProg, timeVal, block?'TRUE':'FALSE']);
  return block?(timeVal+' 차단되었습니다.'):(timeVal+' 차단이 해제되었습니다.');
}

function submitSessionOnly(data) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var isOpen=ss.getSheetByName('Settings').getRange('B1').getValue().toString().trim().toUpperCase()==='OPEN';
  if (!isOpen) throw new Error('현재 예약이 마감되었습니다.');
  if (!/^\d{10}$/.test(data.studentId.toString())) throw new Error('학번은 숫자 10자리여야 합니다.');
  if (!data.sessions||!data.sessions.length) throw new Error('참여할 설명회를 선택해주세요.');

  // 선착순 체크
  var sessSheet = ss.getSheetByName('SessionPreReg');
  if (!sessSheet) {
    sessSheet = ss.insertSheet('SessionPreReg');
    sessSheet.getRange(1,1,1,7).setValues([['이름','학번','학과','연락처','이메일','설명회명','등록일시']]);
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    // 선착순 초과 검사
    var counts = {};
    if (sessSheet.getLastRow() > 1) {
      var allRows = sessSheet.getDataRange().getValues();
      for (var i = 1; i < allRows.length; i++) {
        var p = allRows[i][5] ? allRows[i][5].toString().trim() : '';
        counts[p] = (counts[p]||0) + 1;
      }
    }
    for (var si = 0; si < data.sessions.length; si++) {
      var prog = data.sessions[si].toString().trim();
      if ((counts[prog]||0) >= SESSION_CAPACITY) {
        throw new Error('[' + prog + '] 설명회 예약이 마감되었습니다. (선착순 ' + SESSION_CAPACITY + '명)');
      }
    }

    var alreadySet = new Set();
    if (sessSheet.getLastRow() > 1) {
      var rows2 = sessSheet.getDataRange().getValues();
      for (var j = 1; j < rows2.length; j++) {
        if (rows2[j][1].toString().trim() === data.studentId.toString()) alreadySet.add(rows2[j][5].toString().trim());
      }
    }
    var toAdd = data.sessions.filter(function(s){ return !alreadySet.has(s.toString().trim()); });
    if (!toAdd.length) throw new Error('선택하신 설명회는 이미 모두 등록되어 있습니다.');

    var now=new Date(), name=data.name.toString().trim(), sid=data.studentId.toString().trim();
    var dept=data.dept?data.dept.toString().trim():'', phone=data.phone?data.phone.toString().trim():'';
    var email=data.email?data.email.toString().trim()+'@hanyang.ac.kr':'';
    _appendSessionRows(sessSheet, name, sid, dept, phone, email, toAdd, now);

    try { if (email) sendConfirmMail(email, name, toAdd, [], false); } catch(mailErr) {}
    return toAdd.length+'개 설명회 참여 등록이 완료되었습니다.';
  } finally {
    lock.releaseLock();
  }
}

function _appendSessionRows(sheet, name, sid, dept, phone, email, sessions, now) {
  for (var si=0;si<sessions.length;si++) {
    sheet.appendRow([name, sid, dept, phone, email, sessions[si].toString().trim(), now]);
  }
}

function adminLogin(password) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var admins=ss.getSheetByName('AdminUsers').getDataRange().getValues();
  for (var i=1;i<admins.length;i++) {
    var prog=admins[i][2]?admins[i][2].toString().trim():'';
    var pw=admins[i][3]?admins[i][3].toString().trim():'';
    if (prog==='전체관리'&&pw===password.toString().trim()) return { ok:true, name:admins[i][0].toString().trim() };
  }
  throw new Error('비밀번호가 일치하지 않습니다.');
}

function getSessionAttendees(password, program) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var admins=ss.getSheetByName('AdminUsers').getDataRange().getValues();
  var authorized=false;
  for (var i=1;i<admins.length;i++) {
    if (admins[i][2].toString().trim()==='전체관리'&&admins[i][3].toString().trim()===password.toString().trim()) { authorized=true; break; }
  }
  if (!authorized) throw new Error('권한이 없습니다.');
  var sheet=ss.getSheetByName('SessionPreReg');
  if (!sheet||sheet.getLastRow()<=1) return [];
  var rows=sheet.getDataRange().getValues();
  var result=[];
  for (var j=1;j<rows.length;j++) {
    if (rows[j][5].toString().trim()===program) result.push({ no:result.length+1, name:rows[j][0]?rows[j][0].toString().trim():'', id:rows[j][1]?rows[j][1].toString().trim():'', dept:rows[j][2]?rows[j][2].toString().trim():'' });
  }
  result.sort(function(a,b){ return a.name<b.name?-1:a.name>b.name?1:0; });
  return result;
}



function getSessionStatus(password, program) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var admins=ss.getSheetByName('AdminUsers').getDataRange().getValues();
  var authorized=false;
  for (var i=1;i<admins.length;i++) {
    if (admins[i][2].toString().trim()==='전체관리'&&admins[i][3].toString().trim()===password.toString().trim()) { authorized=true; break; }
  }
  if (!authorized) throw new Error('권한이 없습니다.');

  // CheckIns에서 이 설명회를 체크인한 학생 목록만 구성
  var result=[];
  var seenIds={};
  var ciSheet=ss.getSheetByName('CheckIns');
  if (ciSheet&&ciSheet.getLastRow()>1) {
    var ciRows=ciSheet.getDataRange().getValues();
    for (var c=1;c<ciRows.length;c++) {
      var cSid=ciRows[c][1]?ciRows[c][1].toString().trim():'';
      var cSess=ciRows[c][7]?ciRows[c][7].toString().trim():'';
      if (!cSid||cSess!==program||seenIds[cSid]) continue;
      seenIds[cSid]=true;
      result.push({
        no:        0,
        name:      ciRows[c][2]?ciRows[c][2].toString().trim():'',
        id:        cSid,
        dept:      ciRows[c][3]?ciRows[c][3].toString().trim():'',
        attendType: ciRows[c][6]?ciRows[c][6].toString().trim():'사전예약',
        checkedIn:  true,
        checkedAt:  ciRows[c][0]?ciRows[c][0].toString():''
      });
    }
  }

  result.sort(function(a,b){ return a.name<b.name?-1:a.name>b.name?1:0; });
  result.forEach(function(r,i){ r.no=i+1; });
  return {
    program:   program,
    total:     result.length,
    signed:    result.length,
    attendees: result
  };
}

function getOverallStats(password) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var admins=ss.getSheetByName('AdminUsers').getDataRange().getValues();
  var authorized=false;
  for (var i=1;i<admins.length;i++) {
    if (admins[i][2].toString().trim()==='전체관리'&&admins[i][3].toString().trim()===password.toString().trim()) { authorized=true; break; }
  }
  if (!authorized) throw new Error('권한이 없습니다.');
  var SESSION_PROGS=_getSettings().sessions.map(function(s){ return s.name; });

  // ── 설명회 사전예약 수 (프로그램별) ──
  var preMap={};
  SESSION_PROGS.forEach(function(p){ preMap[p]=0; });
  var sessSheet=ss.getSheetByName('SessionPreReg');
  if (sessSheet&&sessSheet.getLastRow()>1) {
    var sessRows=sessSheet.getDataRange().getValues();
    for (var j=1;j<sessRows.length;j++) {
      var prog=sessRows[j][5]?sessRows[j][5].toString().trim():'';
      if (preMap[prog]!==undefined) preMap[prog]++;
    }
  }

  // ── 설명회 실제 참석 (CheckIns col[7]=참석설명회, col[6]=참석유형) ──
  var attMap={};
  SESSION_PROGS.forEach(function(p){ attMap[p]={pre:0,walk:0}; });
  var sessUniqueIds={}, sessPreUniqueIds={}, sessWalkUniqueIds={}; // 학번 기준 중복 제거용
  var ciSheet=ss.getSheetByName('CheckIns');
  if (ciSheet&&ciSheet.getLastRow()>1) {
    var ciRows=ciSheet.getDataRange().getValues();
    for (var c=1;c<ciRows.length;c++) {
      var aProg=ciRows[c][7]?ciRows[c][7].toString().trim():'';
      var aSid=ciRows[c][1]?ciRows[c][1].toString().trim():'';
      var aType=ciRows[c][6]?ciRows[c][6].toString().trim():'';
      if (aSid) {
        sessUniqueIds[aSid]=true;
        if (aType==='사전예약') sessPreUniqueIds[aSid]=true;
        else sessWalkUniqueIds[aSid]=true;
      }
      if (attMap[aProg]&&aSid) {
        if (aType==='사전예약') attMap[aProg].pre++;
        else attMap[aProg].walk++;
      }
    }
  }
  // 설명회 총계: 건수(중복포함) vs 인원(학번 중복제거)
  var sessTotalCount=0, sessTotalUnique=Object.keys(sessUniqueIds).length;
  var sessPreUnique=Object.keys(sessPreUniqueIds).length, sessWalkUnique=Object.keys(sessWalkUniqueIds).length;
  SESSION_PROGS.forEach(function(p){ sessTotalCount+=(attMap[p].pre||0)+(attMap[p].walk||0); });

  var byProg=SESSION_PROGS.map(function(p){
    return { program:p, preTotal:preMap[p]||0, preSigned:attMap[p].pre||0, walkSigned:attMap[p].walk||0, total:(attMap[p].pre||0)+(attMap[p].walk||0), capacity:SESSION_CAPACITY };
  });

  // ── 기념품 수령 통계 (GiftReceipts) ──
  var giftTotalCount=0, giftUniqueIds={};
  var giftSheet=ss.getSheetByName('GiftReceipts');
  if (giftSheet&&giftSheet.getLastRow()>1) {
    var gRows=giftSheet.getDataRange().getValues();
    for (var g=1;g<gRows.length;g++) {
      var gSid=gRows[g][1]?gRows[g][1].toString().trim():'';
      if (gSid) { giftTotalCount++; giftUniqueIds[gSid]=true; }
    }
  }
  var giftTotalUnique=Object.keys(giftUniqueIds).length;

  return {
    byProg:          byProg,
    // 설명회 통계
    sessTotalCount:  sessTotalCount,   // 참석 건수 (1명이 3개 설명회면 3)
    sessTotalUnique: sessTotalUnique,  // 참석 인원 (학번 중복 제거)
    sessPreUnique:   sessPreUnique,    // 사전예약 참석 인원 (학번 중복 제거)
    sessWalkUnique:  sessWalkUnique,   // 당일 참석 인원 (학번 중복 제거)
    // 기존 호환 필드
    totalPre:        byProg.reduce(function(s,r){ return s+r.preTotal; },0),
    totalPreSign:    byProg.reduce(function(s,r){ return s+r.preSigned; },0),
    totalWalk:       byProg.reduce(function(s,r){ return s+r.walkSigned; },0),
    totalAll:        sessTotalCount,
    // 기념품 통계
    giftTotalCount:  giftTotalCount,   // 수령 건수
    giftTotalUnique: giftTotalUnique   // 수령 인원 (학번 중복 제거)
  };
}

function getAdminDataForOffice(password, selectedProg) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var admins=ss.getSheetByName('AdminUsers').getDataRange().getValues();
  var authorized=false;
  for (var i=1;i<admins.length;i++) {
    if (admins[i][2].toString().trim()==='전체관리'&&admins[i][3].toString().trim()===password.toString().trim()) { authorized=true; break; }
  }
  if (!authorized) throw new Error('권한이 없습니다.');
  var allRows=ss.getSheetByName('BoothReservations').getDataRange().getValues();
  var reservations=[];
  for (var j=1;j<allRows.length;j++) {
    var r=allRows[j];
    var prog=r[5]?r[5].toString().trim():'';
    var status=r[9]?r[9].toString().trim():'';
    if (prog!==selectedProg||status==='취소') continue;
    var st=toTimeStr(r[6]);
    reservations.push({ rowIdx:j+1, name:r[0]?r[0].toString().trim():'', id:r[1]?r[1].toString().trim():'', dept:r[2]?r[2].toString().trim():'', time:st?(st+' ~ '+addMinutes(st,15)):'', status:status });
  }
  reservations.sort(function(a,b){ var ta=a.time?a.time.substring(0,5):'', tb=b.time?b.time.substring(0,5):''; return ta<tb?-1:ta>tb?1:0; });
  return { programName:selectedProg, reservations:reservations };
}

function changeReservation(data) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sid=data.studentId.toString().trim();
  var dept=data.dept?data.dept.toString().trim():'';

  // 설명회 필수 검증
  if (!data.newSessions || data.newSessions.length === 0) {
    throw new Error('설명회를 1개 이상 선택해야 합니다.');
  }
  if (data.newBooths&&data.newBooths.length) {
    // 버그3: LDC는 인터칼리지학부만
    data.newBooths.forEach(function(b){
      if (b.program.toString().trim()==='라이프디자인센터' && dept!==INTERCOLLEGE_DEPT) {
        throw new Error('라이프디자인센터 부스는 한양인터칼리지학부 재학생만 예약 가능합니다.');
      }
    });
    // 버그2: 동일 부스 중복 선택
    var progSet={};
    data.newBooths.forEach(function(b){
      var p=b.program.toString().trim();
      if (progSet[p]) throw new Error('['+p+'] 동일 프로그램을 중복 선택했습니다.');
      progSet[p]=true;
    });
    // 버그1,2: 기존 예약과 부스중복/시간겹침 (취소된 행 제외, oldBoothRowIdxList 제외)
    var resvSheet2=ss.getSheetByName('BoothReservations');
    if (resvSheet2&&resvSheet2.getLastRow()>1) {
      var allRows2=resvSheet2.getDataRange().getValues();
      var oldIdxSet={};
      (data.oldBoothRowIdxList||[]).forEach(function(item){ if(item&&item.rowIdx) oldIdxSet[item.rowIdx]=true; });
      var active2=[];
      for (var ai=1;ai<allRows2.length;ai++) {
        if (oldIdxSet[ai+1]) continue; // 이번에 취소될 행은 제외
        var r2=allRows2[ai];
        var st2=r2[9]?r2[9].toString().trim():'';
        if (st2!==''&&st2!=='취소'&&st2!=='상담취소') {
          active2.push({studentId:r2[1]?r2[1].toString().trim():'',program:r2[5]?r2[5].toString().trim():'',time:toTimeStr(r2[6])});
        }
      }
      data.newBooths.forEach(function(b){
        var prog=b.program.toString().trim();
        var timeVal=b.time.toString().trim();
        if (timeVal.indexOf('~')!==-1) timeVal=timeVal.split('~')[0].trim();
        var dur=slotDuration(prog);
        for (var ai2=0;ai2<active2.length;ai2++) {
          // 버그2: 같은 학번 같은 부스 중복
          if (active2[ai2].studentId===sid&&active2[ai2].program===prog) {
            throw new Error('['+prog+'] 이미 예약하신 부스입니다. 같은 부스는 두 번 예약할 수 없습니다.');
          }
          // 버그1: 같은 학번 시간 겹침
          if (active2[ai2].studentId===sid&&timesOverlap(timeVal,dur,active2[ai2].time,15)) {
            throw new Error('['+prog+'] 선택한 시간이 기존 예약과 겹칩니다.');
          }
        }
      });
      // 새로 선택한 부스들끼리 시간겹침
      for (var ni=0;ni<data.newBooths.length;ni++) {
        for (var nj=ni+1;nj<data.newBooths.length;nj++) {
          var tA=data.newBooths[ni].time.toString().trim(); if(tA.indexOf('~')!==-1)tA=tA.split('~')[0].trim();
          var tB=data.newBooths[nj].time.toString().trim(); if(tB.indexOf('~')!==-1)tB=tB.split('~')[0].trim();
          if (timesOverlap(tA,slotDuration(data.newBooths[ni].program),tB,slotDuration(data.newBooths[nj].program))) {
            throw new Error('['+data.newBooths[ni].program+'/'+data.newBooths[nj].program+'] 선택한 시간대가 서로 겹칩니다.');
          }
        }
      }
    }
  }

  var lock=LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var resvSheet=ss.getSheetByName('BoothReservations');
    var sessSheet=ss.getSheetByName('SessionPreReg');
    // 기존 부스 예약 삭제 — 프로그램+시간으로 재탐색하여 행 번호 밀림 방지
    if (data.oldBoothRowIdxList&&data.oldBoothRowIdxList.length) {
      var boothHints=data.oldBoothRowIdxList; // [{program, time}] 또는 [rowIdx(레거시)]
      var bAllRows=resvSheet&&resvSheet.getLastRow()>1?resvSheet.getDataRange().getValues():[];
      var bToDelete=[];
      boothHints.forEach(function(hint){
        var hProg=hint.program?hint.program.toString().trim():null;
        var hTime=hint.time?hint.time.toString().substring(0,5).trim():null;
        for (var bi=1;bi<bAllRows.length;bi++) {
          var br=bAllRows[bi];
          if (!br) continue;
          if (br[1].toString().trim()!==sid) continue;
          if (br[9]&&(br[9].toString().trim()==='취소'||br[9].toString().trim()==='상담완료')) continue;
          var bProg=br[5]?br[5].toString().trim():'';
          var bTime=br[6]?toTimeStr(br[6]):'';
          if (hProg&&hTime&&bProg===hProg&&bTime===hTime) { bToDelete.push(bi+1); bAllRows[bi]=null; break; }
        }
      });
      bToDelete.sort(function(a,b){return b-a;}).forEach(function(idx){ resvSheet.deleteRow(idx); });
    }
    // 기존 설명회 예약 삭제 — 설명회명으로 재탐색하여 행 번호 밀림 방지
    if (data.oldSessionRowIdxList&&data.oldSessionRowIdxList.length&&sessSheet) {
      var sessHints=data.oldSessionRowIdxList; // [{sessName}] 또는 [rowIdx(레거시)]
      var sAllRows=sessSheet.getLastRow()>1?sessSheet.getDataRange().getValues():[];
      var sToDelete=[];
      sessHints.forEach(function(hint){
        var hSess=hint.sessName?hint.sessName.toString().trim():null;
        for (var si2=1;si2<sAllRows.length;si2++) {
          var sr=sAllRows[si2];
          if (!sr||sr[1].toString().trim()!==sid) continue;
          var sName=sr[5]?sr[5].toString().trim():'';
          if (hSess&&sName===hSess) { sToDelete.push(si2+1); sAllRows[si2]=null; break; }
        }
      });
      sToDelete.sort(function(a,b){return b-a;}).forEach(function(idx){ sessSheet.deleteRow(idx); });
    }
    var now=new Date(), email=data.email?data.email.toString().trim()+'@hanyang.ac.kr':'', phone=data.phone?data.phone.toString().trim():'', name=data.name?data.name.toString().trim():'';
    if (data.newBooths&&data.newBooths.length) {
      data.newBooths.forEach(function(b){ resvSheet.appendRow([name,sid,dept,email,phone,b.program,b.time,b.memo||'','','예약완료','',now]); });
    }
    if (data.newSessions&&data.newSessions.length) {
      if (!sessSheet) { sessSheet=ss.insertSheet('SessionPreReg'); sessSheet.getRange(1,1,1,7).setValues([['이름','학번','학과','연락처','이메일','설명회명','등록일시']]); }
      _appendSessionRows(sessSheet, name, sid, dept, phone, email, data.newSessions, now);
    }
    try {
      if (email) {
        var mailBooths=(data.newBooths||[]).map(function(b){ return {program:b.program,time:b.time}; });
        var hasLdc=(data.newBooths||[]).some(function(b){ return b.memo&&b.memo.indexOf('진로적성검사')!==-1; });
        sendChangeMail(email, name, data.newSessions||[], mailBooths, hasLdc);
      }
    } catch(me) {}
    return '변경이 완료되었습니다.';
  } finally {
    lock.releaseLock();
  }
}

// ─── 메일 함수들 (기존 유지) ───────────────────

function sendConfirmMail(toEmail, name, sessions, reservations, hasLdcTest) {
  var SENDER='intercollege@hanyang.ac.kr';
  var subject='[한양YK인터칼리지] 2026 융합전공 소개행사 예약이 완료되었습니다';
  var SESSION_TIMES={'미래사회디자인':'09:30~09:55','융합의과학/융합의공학':'10:00~10:25','인지융합과학':'10:30~10:55','혁신공학경영':'11:00~11:25','미래반도체공학':'11:30~11:55'};
  var sessRows='';
  if (sessions&&sessions.length) sessions.forEach(function(s){ sessRows+='<tr><td style="padding:8px 14px;border-bottom:1px solid #EAECEF;">'+s+'</td><td style="padding:8px 14px;border-bottom:1px solid #EAECEF;color:#5A6778;">'+(SESSION_TIMES[s]||'')+'</td></tr>'; });
  var boothRows='';
  if (reservations&&reservations.length) reservations.forEach(function(r){ boothRows+='<tr><td style="padding:8px 14px;border-bottom:1px solid #EAECEF;">'+r.program+'</td><td style="padding:8px 14px;border-bottom:1px solid #EAECEF;color:#5A6778;">'+r.time+'</td></tr>'; });
  var ldcNotice=hasLdcTest?'<div style="margin:16px 0;padding:12px 16px;background:#EBF4FF;border-left:4px solid #2B6CB0;border-radius:6px;font-size:14px;color:#2B6CB0;">[진로적성검사] <strong>응시 희망</strong>으로 등록되었습니다.<br>예약하신 상담 시간보다 <strong>15분 일찍</strong> 라이프디자인센터 부스에 도착해 주세요.</div>':'';
  var html='<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="margin:0;padding:0;background:#F4F7F9;font-family:\'Apple SD Gothic Neo\',\'Malgun Gothic\',sans-serif;"><div style="max-width:600px;margin:32px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,0.08);"><div style="background:#1B263B;padding:28px 32px;text-align:center;"><div style="color:#fff;font-size:11px;letter-spacing:0.1em;margin-bottom:6px;opacity:0.7;">HANYANG YK INTERCOLLEGE</div><div style="color:#fff;font-size:20px;font-weight:800;">2026 융합전공 소개행사</div><div style="color:rgba(255,255,255,0.7);font-size:13px;margin-top:4px;">예약 확인 안내</div></div><div style="padding:28px 32px 0;"><p style="font-size:15px;color:#1B263B;margin:0 0 6px;"><strong>'+name+'</strong>님, 안녕하세요!</p><p style="font-size:14px;color:#5A6778;margin:0 0 20px;line-height:1.7;">2026 융합전공 소개행사 예약이 완료되었습니다.</p><div style="background:#F8FAFC;border-radius:8px;padding:14px 18px;margin-bottom:20px;font-size:13px;color:#5A6778;line-height:1.8;"><strong style="color:#1B263B;">일시</strong> &nbsp; 2026. 5. 8.(금) 09:30 ~ 17:00<br><strong style="color:#1B263B;">장소</strong> &nbsp; 한양종합기술원(HIT) 1층 양민용 커리어라운지</div>'+(sessRows?'<p style="font-size:13px;font-weight:700;color:#1B263B;margin:0 0 8px;">참여 설명회</p><table style="width:100%;border-collapse:collapse;font-size:13px;border:1px solid #EAECEF;border-radius:8px;overflow:hidden;margin-bottom:20px;"><thead><tr style="background:#F0F4F8;"><th style="padding:8px 14px;text-align:left;font-weight:600;color:#1B263B;">프로그램</th><th style="padding:8px 14px;text-align:left;font-weight:600;color:#1B263B;">시간</th></tr></thead><tbody>'+sessRows+'</tbody></table>':'')+(boothRows?'<p style="font-size:13px;font-weight:700;color:#1B263B;margin:0 0 8px;">부스 상담 예약</p><table style="width:100%;border-collapse:collapse;font-size:13px;border:1px solid #EAECEF;border-radius:8px;overflow:hidden;margin-bottom:20px;"><thead><tr style="background:#F0F4F8;"><th style="padding:8px 14px;text-align:left;font-weight:600;color:#1B263B;">프로그램</th><th style="padding:8px 14px;text-align:left;font-weight:600;color:#1B263B;">시간</th></tr></thead><tbody>'+boothRows+'</tbody></table>':'')+ldcNotice+'<div style="margin:16px 0;text-align:center;"><a href="https://tinyurl.com/hicmajors" style="display:inline-block;background:#1B263B;color:#fff;font-size:14px;font-weight:700;padding:12px 32px;border-radius:8px;text-decoration:none;">예약 신청 및 확인 / 변경</a><p style="font-size:12px;color:#C53030;font-weight:600;margin:10px 0 0;">※ 예약 취소 및 변경은 위 버튼을 통해 웹페이지에서만 가능합니다.</p></div></div><div style="padding:20px 32px;border-top:1px solid #EAECEF;margin-top:20px;text-align:center;font-size:11px;color:#9AA5B4;">본 메일은 발신 전용입니다.</div></div></body></html>';
  GmailApp.sendEmail(toEmail, subject, '', { htmlBody:html, replyTo:SENDER, name:'한양YK인터칼리지' });
}

function sendCancelMail(toEmail, name, cancelledBooths, cancelledSessions) {
  var SENDER='intercollege@hanyang.ac.kr';
  var subject='[한양YK인터칼리지] 2026 융합전공 소개행사 예약이 취소되었습니다';
  var SESSION_TIMES={'미래사회디자인':'09:30~09:55','융합의과학/융합의공학':'10:00~10:25','인지융합과학':'10:30~10:55','혁신공학경영':'11:00~11:25','미래반도체공학':'11:30~11:55'};
  var rows='';
  (cancelledSessions||[]).forEach(function(s){ rows+='<tr><td style="padding:8px 14px;border-bottom:1px solid #EAECEF;">설명회 | '+s+'</td><td style="padding:8px 14px;border-bottom:1px solid #EAECEF;color:#5A6778;">'+(SESSION_TIMES[s]||'')+'</td></tr>'; });
  (cancelledBooths||[]).forEach(function(b){ rows+='<tr><td style="padding:8px 14px;border-bottom:1px solid #EAECEF;">부스 | '+b.program+'</td><td style="padding:8px 14px;border-bottom:1px solid #EAECEF;color:#5A6778;">'+(b.time||'')+'</td></tr>'; });
  var html='<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="margin:0;padding:0;background:#F4F7F9;font-family:\'Apple SD Gothic Neo\',sans-serif;"><div style="max-width:600px;margin:32px auto;background:#fff;border-radius:12px;overflow:hidden;"><div style="background:#1B263B;padding:28px 32px;text-align:center;"><div style="color:#fff;font-size:20px;font-weight:800;">2026 융합전공 소개행사</div><div style="color:rgba(255,255,255,0.7);font-size:13px;margin-top:4px;">예약 취소 확인 안내</div></div><div style="padding:28px 32px;"><p style="font-size:15px;color:#1B263B;margin:0 0 6px;"><strong>'+name+'</strong>님, 안녕하세요!</p><p style="font-size:14px;color:#5A6778;margin:0 0 20px;">아래 예약이 취소 처리되었습니다.</p>'+(rows?'<table style="width:100%;border-collapse:collapse;font-size:13px;border:1px solid #EAECEF;border-radius:8px;overflow:hidden;margin-bottom:20px;"><thead><tr style="background:#F0F4F8;"><th style="padding:8px 14px;text-align:left;">취소된 예약</th><th style="padding:8px 14px;text-align:left;">시간</th></tr></thead><tbody>'+rows+'</tbody></table>':'')+'</div><div style="padding:20px 32px;border-top:1px solid #EAECEF;text-align:center;font-size:11px;color:#9AA5B4;">본 메일은 발신 전용입니다.</div></div></body></html>';
  GmailApp.sendEmail(toEmail, subject, '', { htmlBody:html, replyTo:SENDER, name:'한양YK인터칼리지' });
}

function sendChangeMail(toEmail, name, sessions, reservations, hasLdcTest) {
  var SENDER='intercollege@hanyang.ac.kr';
  var subject='[한양YK인터칼리지] 2026 융합전공 소개행사 예약이 변경되었습니다';
  var SESSION_TIMES={'미래사회디자인':'09:30~09:55','융합의과학/융합의공학':'10:00~10:25','인지융합과학':'10:30~10:55','혁신공학경영':'11:00~11:25','미래반도체공학':'11:30~11:55'};
  var sessRows='';
  if (sessions&&sessions.length) sessions.forEach(function(s){ sessRows+='<tr><td style="padding:8px 14px;border-bottom:1px solid #EAECEF;">'+s+'</td><td style="padding:8px 14px;border-bottom:1px solid #EAECEF;color:#5A6778;">'+(SESSION_TIMES[s]||'')+'</td></tr>'; });
  var boothRows='';
  if (reservations&&reservations.length) reservations.forEach(function(r){ boothRows+='<tr><td style="padding:8px 14px;border-bottom:1px solid #EAECEF;">'+r.program+'</td><td style="padding:8px 14px;border-bottom:1px solid #EAECEF;color:#5A6778;">'+r.time+'</td></tr>'; });
  var ldcNotice=hasLdcTest?'<div style="margin:16px 0;padding:12px 16px;background:#EBF4FF;border-left:4px solid #2B6CB0;border-radius:6px;font-size:14px;color:#2B6CB0;">[진로적성검사] <strong>응시 희망</strong>으로 등록되었습니다.<br>예약하신 상담 시간보다 <strong>15분 일찍</strong> 라이프디자인센터 부스에 도착해 주세요.</div>':'';
  var html='<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="margin:0;padding:0;background:#F4F7F9;font-family:\'Apple SD Gothic Neo\',\'Malgun Gothic\',sans-serif;"><div style="max-width:600px;margin:32px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,0.08);"><div style="background:#1B263B;padding:28px 32px;text-align:center;"><div style="color:#fff;font-size:11px;letter-spacing:0.1em;margin-bottom:6px;opacity:0.7;">HANYANG YK INTERCOLLEGE</div><div style="color:#fff;font-size:20px;font-weight:800;">2026 융합전공 소개행사</div><div style="color:rgba(255,255,255,0.7);font-size:13px;margin-top:4px;">예약 변경 완료 안내</div></div><div style="padding:28px 32px 0;"><p style="font-size:15px;color:#1B263B;margin:0 0 6px;"><strong>'+name+'</strong>님, 안녕하세요!</p><p style="font-size:14px;color:#5A6778;margin:0 0 20px;line-height:1.7;">2026 융합전공 소개행사 예약이 <strong style="color:#1B263B;">변경</strong>되었습니다. 아래 내용을 확인해 주세요.</p><div style="background:#F8FAFC;border-radius:8px;padding:14px 18px;margin-bottom:20px;font-size:13px;color:#5A6778;line-height:1.8;"><strong style="color:#1B263B;">일시</strong> &nbsp; 2026. 5. 8.(금) 09:30 ~ 17:00<br><strong style="color:#1B263B;">장소</strong> &nbsp; 한양종합기술원(HIT) 1층 양민용 커리어라운지</div>'+(sessRows?'<p style="font-size:13px;font-weight:700;color:#1B263B;margin:0 0 8px;">참여 설명회</p><table style="width:100%;border-collapse:collapse;font-size:13px;border:1px solid #EAECEF;border-radius:8px;overflow:hidden;margin-bottom:20px;"><thead><tr style="background:#F0F4F8;"><th style="padding:8px 14px;text-align:left;font-weight:600;color:#1B263B;">프로그램</th><th style="padding:8px 14px;text-align:left;font-weight:600;color:#1B263B;">시간</th></tr></thead><tbody>'+sessRows+'</tbody></table>':'')+(boothRows?'<p style="font-size:13px;font-weight:700;color:#1B263B;margin:0 0 8px;">부스 상담 예약</p><table style="width:100%;border-collapse:collapse;font-size:13px;border:1px solid #EAECEF;border-radius:8px;overflow:hidden;margin-bottom:20px;"><thead><tr style="background:#F0F4F8;"><th style="padding:8px 14px;text-align:left;font-weight:600;color:#1B263B;">프로그램</th><th style="padding:8px 14px;text-align:left;font-weight:600;color:#1B263B;">시간</th></tr></thead><tbody>'+boothRows+'</tbody></table>':'')+ldcNotice+'<div style="margin:16px 0;text-align:center;"><a href="https://tinyurl.com/hicmajors" style="display:inline-block;background:#1B263B;color:#fff;font-size:14px;font-weight:700;padding:12px 32px;border-radius:8px;text-decoration:none;">예약 신청 및 확인 / 변경</a><p style="font-size:12px;color:#C53030;font-weight:600;margin:10px 0 0;">※ 추가 변경 또는 취소는 위 버튼을 통해 웹페이지에서만 가능합니다.</p></div></div><div style="padding:20px 32px;border-top:1px solid #EAECEF;margin-top:20px;text-align:center;font-size:11px;color:#9AA5B4;">본 메일은 발신 전용입니다.</div></div></body></html>';
  GmailApp.sendEmail(toEmail, subject, '', { htmlBody:html, replyTo:SENDER, name:'한양YK인터칼리지' });
}

// ─── 트리거 전체 초기화 (불필요한 것 제거 + 필요한 것 복구) ──────────
// GAS 편집기에서 fixAllTriggers() 한 번만 실행하면 됨
function fixAllTriggers() {
  // 1. 모든 트리거 제거
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });

  // 2. warmup 트리거 등록 (30분마다 콜드스타트 방지, 행사 후 자동 삭제)
  ScriptApp.newTrigger('warmup').timeBased().everyMinutes(30).create();

  // 3. 리마인드 메일 트리거 등록 (2026-05-07 09:00 KST 1회 발송)
  ScriptApp.newTrigger('sendRemindMails').timeBased().at(new Date('2026-05-07T09:00:00+09:00')).create();

  Logger.log('트리거 설정 완료: warmup(30분), sendRemindMails(2026-05-07 09:00)');
  Logger.log('제거된 불필요 트리거: autoSyncFirebaseToSheets, syncFirestoreToSheets');
}

// ─── Firestore → Sheets 1분 트리거 동기화 (폐기됨 — Firebase IAM 403 오류) ──
// 현재 아키텍처: doPost 방식 사용, 이 트리거는 사용하지 않음
function removeSyncTrigger() {
  var removed = [];
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'syncFirestoreToSheets') {
      ScriptApp.deleteTrigger(t);
      removed.push('syncFirestoreToSheets');
    }
  });
  Logger.log('삭제된 트리거: ' + (removed.length ? removed.join(', ') : '없음'));
}
function setupSyncTrigger() {
  Logger.log('이 함수는 사용하지 않습니다. 현재 아키텍처는 doPost 방식입니다.');
}

function syncFirestoreToSheets() {
  var PROJECT_ID = 'hic-major-booking-656ca';
  var token = ScriptApp.getOAuthToken();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();

  // ── Sheets 현재 상태 로드 ──
  var sessSheet  = _getOrCreateSheet(ss, 'SessionPreReg',  ['이름','학번','학과','연락처','이메일','설명회명','등록일시']);
  var boothSheet = _getOrCreateSheet(ss, 'BoothReservations', ['이름','학번','학과','이메일','연락처','프로그램','시간','문의내용','서명','상태','코멘트','예약일시']);

  // 시트에 있는 (학번+설명회명) 키 세트
  var existSessKeys = {};
  if (sessSheet.getLastRow() > 1) {
    sessSheet.getDataRange().getValues().slice(1).forEach(function(r) {
      existSessKeys[r[1].toString().trim() + '|' + r[5].toString().trim()] = true;
    });
  }
  // 시트에 있는 (학번+프로그램) 키 세트 (취소된 행 제외)
  var existBoothKeys = {};
  if (boothSheet.getLastRow() > 1) {
    boothSheet.getDataRange().getValues().slice(1).forEach(function(r) {
      var st = r[9] ? r[9].toString().trim() : '';
      if (st !== '취소') existBoothKeys[r[1].toString().trim() + '|' + r[5].toString().trim()] = true;
    });
  }

  // ── Firestore sessionPreReg 전체 조회 ──
  var newSessMail = {}; // 새로 추가된 항목: sid → {name,email,sessions[],dept,phone}
  var sessUrl = 'https://firestore.googleapis.com/v1/projects/' + PROJECT_ID + '/databases/(default)/documents/sessionPreReg?pageSize=500';
  var sessDocs = _firestoreFetchAll(sessUrl, token);
  sessDocs.forEach(function(doc) {
    var f = doc.fields || {};
    var sid      = _fsStr(f.studentId);
    var sessName = _fsStr(f.sessionName);
    var name     = _fsStr(f.name);
    var dept     = _fsStr(f.dept);
    var phone    = _fsStr(f.phone);
    var email    = _fsStr(f.email);
    if (!sid || !sessName) return;
    var key = sid + '|' + sessName;
    if (existSessKeys[key]) return; // 이미 시트에 있음
    sessSheet.appendRow([name, sid, dept, phone, email, sessName, now]);
    existSessKeys[key] = true;
    // 메일 발송 대상 수집
    if (!newSessMail[sid]) newSessMail[sid] = {name:name, email:email, dept:dept, phone:phone, sessions:[], booths:[]};
    newSessMail[sid].sessions.push(sessName);
  });

  // ── Firestore boothReservations 전체 조회 ──
  var boothUrl = 'https://firestore.googleapis.com/v1/projects/' + PROJECT_ID + '/databases/(default)/documents/boothReservations?pageSize=500';
  var boothDocs = _firestoreFetchAll(boothUrl, token);
  boothDocs.forEach(function(doc) {
    var f = doc.fields || {};
    var sid     = _fsStr(f.studentId);
    var prog    = _fsStr(f.program);
    var time    = _fsStr(f.time);
    var name    = _fsStr(f.name);
    var dept    = _fsStr(f.dept);
    var phone   = _fsStr(f.phone);
    var email   = _fsStr(f.email);
    var memo    = _fsStr(f.memo);
    var status  = _fsStr(f.status);
    if (!sid || !prog) return;
    if (status === '취소' || status === '상담취소') return;
    var key = sid + '|' + prog;
    if (existBoothKeys[key]) return;
    boothSheet.appendRow([name, sid, dept, email, phone, prog, time, memo, '', status||'예약완료', '', now]);
    existBoothKeys[key] = true;
    if (!newSessMail[sid]) newSessMail[sid] = {name:name, email:email, dept:dept, phone:phone, sessions:[], booths:[]};
    newSessMail[sid].booths.push({program:prog, time:time});
  });

  // ── 신규 항목에 대해 확인 메일 발송 ──
  Object.keys(newSessMail).forEach(function(sid) {
    var m = newSessMail[sid];
    if (!m.email) return;
    try {
      var hasLdc = m.booths.some(function(b){ return b.program === '라이프디자인센터'; });
      sendConfirmMail(m.email, m.name, m.sessions, m.booths, hasLdc);
      Utilities.sleep(200);
    } catch(e) { Logger.log('메일 발송 실패 (' + sid + '): ' + e.message); }
  });
}

// Firestore REST API 페이지네이션 처리
function _firestoreFetchAll(url, token) {
  var docs = [];
  var nextUrl = url;
  while (nextUrl) {
    var res = UrlFetchApp.fetch(nextUrl, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) {
      Logger.log('Firestore 오류: ' + res.getContentText());
      break;
    }
    var data = JSON.parse(res.getContentText());
    if (data.documents) docs = docs.concat(data.documents);
    nextUrl = data.nextPageToken ? url + '&pageToken=' + data.nextPageToken : null;
  }
  return docs;
}

// Firestore 필드값 문자열 추출
function _fsStr(field) {
  if (!field) return '';
  return (field.stringValue || field.integerValue || '').toString().trim();
}

// 시트가 없으면 생성
function _getOrCreateSheet(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function sendRemindMails() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var SESSION_TIMES={'미래사회디자인':'09:30~09:55','융합의과학/융합의공학':'10:00~10:25','인지융합과학':'10:30~10:55','혁신공학경영':'11:00~11:25','미래반도체공학':'11:30~11:55'};
  var SENDER='intercollege@hanyang.ac.kr';
  var subject='[한양YK인터칼리지] 내일 행사가 있습니다! 최종 예약 내역을 확인해 주세요';
  var personMap={};
  var sessSheet=ss.getSheetByName('SessionPreReg');
  if (sessSheet&&sessSheet.getLastRow()>1) {
    var sessRows=sessSheet.getDataRange().getValues();
    for (var i=1;i<sessRows.length;i++) {
      var r=sessRows[i], sid=r[1]?r[1].toString().trim():'', email=r[4]?r[4].toString().trim():'';
      if (!sid||!email) continue;
      if (!personMap[sid]) personMap[sid]={name:r[0].toString().trim(),email:email,sessions:[],booths:[],hasLdcTest:false};
      personMap[sid].sessions.push(r[5].toString().trim());
    }
  }
  var resvSheet=ss.getSheetByName('BoothReservations');
  if (resvSheet&&resvSheet.getLastRow()>1) {
    var resvRows=resvSheet.getDataRange().getValues();
    for (var j=1;j<resvRows.length;j++) {
      var rv=resvRows[j], status=rv[9]?rv[9].toString().trim():'';
      if (status==='취소') continue;
      var sid2=rv[1]?rv[1].toString().trim():'', email2=rv[3]?rv[3].toString().trim():'';
      if (!sid2||!email2) continue;
      if (!personMap[sid2]) personMap[sid2]={name:rv[0].toString().trim(),email:email2,sessions:[],booths:[],hasLdcTest:false};
      personMap[sid2].booths.push({program:rv[5]?rv[5].toString().trim():'',time:rv[6]?rv[6].toString().trim():''});
      if (rv[7]&&rv[7].toString().indexOf('진로적성검사')!==-1) personMap[sid2].hasLdcTest=true;
    }
  }
  var count=0;
  for (var sid3 in personMap) {
    var p=personMap[sid3];
    if (!p.email) continue;
    try { sendConfirmMail(p.email, p.name, p.sessions, p.booths, p.hasLdcTest); count++; Utilities.sleep(200); } catch(e) { Logger.log('메일 발송 실패 ('+p.email+'): '+e.message); }
  }
  Logger.log('리마인드 메일 발송 완료: '+count+'명');
}

// ── 워밍업 (콜드 스타트 방지) ──────────────────────────────────────
function warmup() {
  var now = new Date();
  // 행사 당일(2026-05-08) 이후엔 트리거 자동 삭제
  if (now > new Date('2026-05-08T18:00:00+09:00')) {
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === 'warmup') ScriptApp.deleteTrigger(t);
    });
    return;
  }
  Logger.log('warmup ping ' + now.toISOString());
}

// 30분 간격 반복 트리거 1개 등록 (GAS 트리거 한도 절약)
function setupWarmupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'warmup') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('warmup').timeBased().everyMinutes(30).create();
  Logger.log('워밍업 트리거 등록 완료 (30분 간격, 행사 후 자동 삭제)');
}

function setupRemindTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t){ if (t.getHandlerFunction()==='sendRemindMails') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendRemindMails').timeBased().at(new Date('2026-05-07T09:00:00+09:00')).create();
  Logger.log('리마인드 트리거 등록 완료');
}

function debugSessionReservations() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getSheetByName('SessionPreReg');
  if (!sheet) { Logger.log('SessionReservations 없음'); return; }
  var rows=sheet.getDataRange().getValues();
  Logger.log('총 행: '+rows.length);
  for (var i=0;i<Math.min(rows.length,6);i++) Logger.log('행'+(i+1)+': '+JSON.stringify(rows[i]));
}

// ─────────────────────────────────────────────
// 고유 검증 코드 생성 유틸
// 형식: YK2026-XXXXXX (영숫자 6자리)
// ─────────────────────────────────────────────
function generateVerifyCode() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // 혼동 문자(0,O,1,I) 제외
  var code = '';
  for (var i = 0; i < 6; i++) {
    code += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return 'YK2026-' + code;
}

// ─────────────────────────────────────────────
// 학생 종합 정보 조회 (gift.html용 — 개선된 버전)
// 체크인 + 스탬프 + 수령여부 + 사전예약 설명회 모두 포함
// ─────────────────────────────────────────────
function getStudentFullStatus(studentId) {
  var sid = studentId.toString().trim();
  if (!/^\d{10}$/.test(sid)) throw new Error('학번은 숫자 10자리여야 합니다.');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ciSheet = ss.getSheetByName('CheckIns');
  var grSheet = ss.getSheetByName('GiftReceipts');
  var prSheet = ss.getSheetByName('SessionPreReg');
  var brSheet = ss.getSheetByName('BoothReservations');
  var ciRows = (ciSheet && ciSheet.getLastRow() > 1) ? ciSheet.getDataRange().getValues() : [];
  var grRows = (grSheet && grSheet.getLastRow() > 1) ? grSheet.getDataRange().getValues() : [];
  var prRows = (prSheet && prSheet.getLastRow() > 1) ? prSheet.getDataRange().getValues() : [];
  var brRows = (brSheet && brSheet.getLastRow() > 1) ? brSheet.getDataRange().getValues() : [];

  // 1. 체크인 여부
  var checkIn = null;
  for (var i = 1; i < ciRows.length; i++) {
    if (ciRows[i][1].toString().trim() === sid) {
      checkIn = {
        checkedAt: ciRows[i][0] ? ciRows[i][0].toString() : '',
        name:  ciRows[i][2] ? ciRows[i][2].toString().trim() : '',
        dept:  ciRows[i][3] ? ciRows[i][3].toString().trim() : '',
        type:  ciRows[i][6] ? ciRows[i][6].toString().trim() : ''
      };
      break;
    }
  }

  // 2. 기념품 수령 여부
  var alreadyReceived = false, existingStampCount = 0;
  for (var j = 1; j < grRows.length; j++) {
    if (grRows[j][1].toString().trim() === sid) {
      alreadyReceived = true;
      if (!existingStampCount) existingStampCount = grRows[j][5] ? Number(grRows[j][5]) : 0;
    }
  }

  // 3. 사전예약 설명회
  var sessions = [], prName = '', prDept = '';
  for (var k = 1; k < prRows.length; k++) {
    if (prRows[k][1].toString().trim() === sid) {
      if (!prName) prName = prRows[k][0].toString().trim();
      if (!prDept) prDept = prRows[k][2].toString().trim();
      sessions.push(prRows[k][5].toString().trim());
    }
  }

  // 4. 부스 예약 — 전체(표시용) + 상담완료(자격 검증용) 분리
  var booths = [], completedBooths = [];
  for (var b = 1; b < brRows.length; b++) {
    var bSid    = brRows[b][1] ? brRows[b][1].toString().trim() : '';
    var bStatus = brRows[b][9] ? brRows[b][9].toString().trim() : '';
    if (bSid !== sid || bStatus === '취소' || bStatus === '상담취소') continue;
    var bProg = brRows[b][5] ? brRows[b][5].toString().trim() : '';
    var bTime = brRows[b][6] ? toTimeStr(brRows[b][6]) : '';
    if (bProg) booths.push(bProg + (bTime ? ' ' + bTime : ''));
    if (bStatus === '상담완료' && bProg) completedBooths.push(bProg);
  }

  var name = checkIn ? checkIn.name : prName;
  var dept = checkIn ? checkIn.dept : prDept;
  var completedCount = completedBooths.length;

  return {
    studentId:        sid,
    name:             name,
    dept:             dept,
    isIntercollege:   dept === INTERCOLLEGE_DEPT,
    checkedIn:        !!checkIn,
    checkInTime:      checkIn ? checkIn.checkedAt : '',
    checkInType:      checkIn ? checkIn.type : '',
    preRegSessions:   sessions,
    preRegBooths:     booths,
    completedBooths:  completedBooths,   // 상담완료 부스 목록
    completedCount:   completedCount,    // 상담완료 건수 (자격 기준)
    stamps:           [],
    stampCount:       existingStampCount,
    stampRequired:    BOOTH_STAMP_REQUIRED,
    eligible:         !!checkIn && completedCount >= BOOTH_STAMP_REQUIRED,
    alreadyReceived:  alreadyReceived,
    bonusSubjects:    BONUS_SUBJECTS
  };
}

// ─────────────────────────────────────────────
// 기념품 수령 저장 (검증코드 포함 — 개선된 버전)
// GiftReceipts: 수령시각|학번|이름|학과|스탬프수|가산점과목|서명|검증코드|검증코드
// ─────────────────────────────────────────────
function saveGiftReceiptV2(data) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var sid = data.studentId.toString().trim();
  if (!/^\d{10}$/.test(sid)) throw new Error('학번은 숫자 10자리여야 합니다.');

  // 조건 재검증
  var status = getStudentFullStatus(sid);
  if (!status.checkedIn)     throw new Error('설명회 체크인 기록이 없습니다.');
  var incomingCount = (data.booths && data.booths.length) ? data.booths.length : (data.stampCount || 0);
  if (incomingCount < BOOTH_STAMP_REQUIRED)
    throw new Error('부스 스탬프가 ' + BOOTH_STAMP_REQUIRED + '개 이상이어야 합니다. (현재 ' + incomingCount + '개)');
  if (status.alreadyReceived) throw new Error('이미 기념품을 수령하셨습니다.');

  var sheet = ss.getSheetByName('GiftReceipts');
  if (!sheet) {
    sheet = ss.insertSheet('GiftReceipts');
    sheet.getRange(1,1,1,9).setValues([['수령시각','학번','이름','학과','방문부스','스탬프수','가산점과목','서명','검증코드']]);
    sheet.getRange(1,1,1,9).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    // lock 내 중복 재확인
    if (sheet.getLastRow() > 1) {
      var rows = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][1].toString().trim() === sid) throw new Error('이미 기념품을 수령하셨습니다.');
      }
    }
    // 기존 9컬럼 시트면 그대로, 8컬럼이면 헤더에 검증코드 추가
    var lastCol = sheet.getLastColumn();
    if (lastCol < 9 && sheet.getLastRow() > 0) {
      sheet.getRange(1, 5).setValue('방문부스');
      sheet.getRange(1, 9).setValue('검증코드');
      sheet.getRange(1, 5, 1, 5).setFontWeight('bold');
    }

    var verifyCode = generateVerifyCode();
    var booths   = (data.booths && data.booths.length) ? data.booths : [];
    var stampCnt = booths.length || (data.stampCount || 0);
    var name3    = data.name.toString().trim();
    var dept3    = data.dept.toString().trim();
    var subj3    = data.bonusSubject ? data.bonusSubject.toString().trim() : '';
    var sig3     = data.signature.toString();
    var now3     = new Date();
    // 부스당 1행 저장 (부스 없으면 빈칸으로 1행)
    if (booths.length === 0) {
      sheet.appendRow([now3, sid, name3, dept3, '', stampCnt, subj3, sig3, verifyCode]);
    } else {
      booths.forEach(function(booth) {
        sheet.appendRow([now3, sid, name3, dept3, booth.toString().trim(), stampCnt, subj3, sig3, verifyCode]);
      });
    }
    return { message: '기념품 수령이 완료되었습니다.', verifyCode: verifyCode };
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────
// 검증 코드로 수령 정보 조회 (공개 — 인증 불필요)
// ?page=verify 전용
// ─────────────────────────────────────────────
function verifyByCode(code) {
  if (!code || code.toString().trim().length < 5) throw new Error('올바른 검증 코드를 입력해주세요.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('GiftReceipts');
  if (!sheet || sheet.getLastRow() <= 1) return null;
  var rows = sheet.getDataRange().getValues();
  var codeUpper = code.toString().trim().toUpperCase();
  // 1단계: 검증코드로 학번 찾기
  var targetSid = null, baseRow = null;
  for (var i = 1; i < rows.length; i++) {
    var rowCode = rows[i][8] ? rows[i][8].toString().trim() : '';
    if (rowCode === codeUpper) {
      targetSid = rows[i][1] ? rows[i][1].toString().trim() : '';
      baseRow = rows[i];
      break;
    }
  }
  if (!targetSid) return { valid: false };
  // 2단계: 해당 학번의 모든 부스 행 수집 (부스당 1행 구조)
  var boothList = [];
  for (var j = 1; j < rows.length; j++) {
    if (rows[j][1].toString().trim() === targetSid) {
      var b = rows[j][4] ? rows[j][4].toString().trim() : '';
      if (b) boothList.push(b);
    }
  }
  return {
    valid:          true,
    name:           baseRow[2] ? baseRow[2].toString().trim() : '',
    studentId:      targetSid,
    dept:           baseRow[3] ? baseRow[3].toString().trim() : '',
    isIntercollege: baseRow[3] ? baseRow[3].toString().trim() === INTERCOLLEGE_DEPT : false,
    booths:         boothList.join(', '),
    stampCount:     baseRow[5] ? Number(baseRow[5]) : 0,
    bonusSubject:   baseRow[6] ? baseRow[6].toString().trim() : '',
    receivedAt:     baseRow[0] ? baseRow[0].toString() : '',
    verifyCode:     codeUpper
  };
}

// ─────────────────────────────────────────────
// 참여확인서 데이터 조회
// 데이터 출처: CheckIns + GiftReceipts 만 사용
// SessionPreReg는 참석유형(사전예약/당일방문) 판별에도 사용하지 않음
// — 참석유형은 CheckIns G열(col[6])에 이미 저장되어 있음
// ─────────────────────────────────────────────
function getCertificateData(password, studentId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var admins = ss.getSheetByName('AdminUsers').getDataRange().getValues();
  var authorized = false;
  for (var i = 1; i < admins.length; i++) {
    if (admins[i][2].toString().trim() === '전체관리' &&
        admins[i][3].toString().trim() === password.toString().trim()) {
      authorized = true; break;
    }
  }
  if (!authorized) throw new Error('권한이 없습니다.');

  var sid = studentId.toString().trim();
  if (!/^\d{10}$/.test(sid)) throw new Error('학번은 숫자 10자리여야 합니다.');

  // ── CheckIns 시트에서 모든 데이터 수집 ──────────────────────────
  // 컬럼: [0]체크인시각 [1]학번 [2]이름 [3]학과 [4]연락처 [5]이메일 [6]참석유형 [7]참석설명회
  var checkInTime = '', checkInType = '', name = '', dept = '', phone = '', email = '';
  var attendedSessions = [];

  var SESSION_ENDS = {
    '미래사회디자인':        '09:55',
    '융합의과학/융합의공학': '10:25',
    '인지융합과학':          '10:55',
    '혁신공학경영':          '11:25',
    '미래반도체공학':        '11:55'
  };

  var ciSheet = ss.getSheetByName('CheckIns');
  var hasCheckInRow = false;
  if (ciSheet && ciSheet.getLastRow() > 1) {
    var ciRows = ciSheet.getDataRange().getValues();
    for (var c = 1; c < ciRows.length; c++) {
      if (ciRows[c][1].toString().trim() !== sid) continue;
      hasCheckInRow = true;
      // 첫 번째 행에서 기본 정보 수집
      if (!name) {
        var rawTime = ciRows[c][0];
        checkInTime = rawTime ? rawTime.toString() : '';
        name        = ciRows[c][2] ? ciRows[c][2].toString().trim() : '';
        dept        = ciRows[c][3] ? ciRows[c][3].toString().trim() : '';
        phone       = ciRows[c][4] ? ciRows[c][4].toString().trim() : '';
        email       = ciRows[c][5] ? ciRows[c][5].toString().trim() : '';
        checkInType = ciRows[c][6] ? ciRows[c][6].toString().trim() : '';
      }
      // H열: 설명회명 — 저장 시 이미 시각 검증 완료된 값
      var sessName = ciRows[c][7] ? ciRows[c][7].toString().trim() : '';
      if (sessName && attendedSessions.indexOf(sessName) === -1) {
        attendedSessions.push(sessName);
      }
    }
  }

  var checkIn = hasCheckInRow ? {
    checkedAt: checkInTime,
    type:      checkInType,
    name:      name,
    dept:      dept
  } : null;

  // ── GiftReceipts 시트에서 수령 정보 수집 ─────────────────────────
  // 컬럼: [0]수령시각 [1]학번 [2]이름 [3]학과 [4]방문부스(행별 1개) [5]스탬프수 [6]가산점과목 [7]서명 [8]검증코드
  var giftInfo = null;
  var boothList = [];

  var giftSheet = ss.getSheetByName('GiftReceipts');
  if (giftSheet && giftSheet.getLastRow() > 1) {
    var gRows = giftSheet.getDataRange().getValues();
    for (var g = 1; g < gRows.length; g++) {
      if (gRows[g][1].toString().trim() !== sid) continue;
      if (!giftInfo) {
        // 이름/학과를 GiftReceipts에서도 확인 (CheckIns보다 최신일 수 있음)
        if (!name) name = gRows[g][2] ? gRows[g][2].toString().trim() : '';
        if (!dept) dept = gRows[g][3] ? gRows[g][3].toString().trim() : '';
        giftInfo = {
          receivedAt:   gRows[g][0] ? gRows[g][0].toString() : '',
          stampCount:   gRows[g][5] ? Number(gRows[g][5]) : 0,
          bonusSubject: gRows[g][6] ? gRows[g][6].toString().trim() : '',
          signature:    gRows[g][7] ? gRows[g][7].toString() : '',
          verifyCode:   gRows[g][8] ? gRows[g][8].toString().trim() : ''
        };
      }
      var booth = gRows[g][4] ? gRows[g][4].toString().trim() : '';
      if (booth) boothList.push(booth);
    }
    if (giftInfo) giftInfo.booths = boothList.join(', ');
  }

  return {
    studentId:        sid,
    name:             name,
    dept:             dept,
    isIntercollege:   dept === INTERCOLLEGE_DEPT,
    checkIn:          checkIn,          // CheckIns 기반 (시각, 참석유형)
    attendedSessions: attendedSessions, // CheckIns H열 기반 실제 참석 설명회
    stamps:           [],               // StampLogs 미사용 (GiftReceipts.booths로 대체)
    giftInfo:         giftInfo,         // GiftReceipts 기반 (부스, 서명, 검증코드 등)
    sealImageUrl:     SEAL_IMAGE_URL,
    verifyBaseUrl:    VERIFY_PAGE_URL
  };
}

// ─────────────────────────────────────────────
// 참여확인서 이메일 발송 (admin.html에서 호출)
// 인터칼리지 학생에게 HTML 메일로 발송
// ─────────────────────────────────────────────
// ─────────────────────────────────────────────
// 이메일 주소 조회 헬퍼
// CheckIns col[5] 만 사용 (SessionPreReg 미참조)
// ─────────────────────────────────────────────
function _getStudentEmail(studentId) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var sid = studentId.toString().trim();
  var ciSheet = ss.getSheetByName('CheckIns');
  if (ciSheet && ciSheet.getLastRow() > 1) {
    var ciRows = ciSheet.getDataRange().getValues();
    for (var c = 1; c < ciRows.length; c++) {
      if (ciRows[c][1].toString().trim() === sid) {
        var em = ciRows[c][5] ? ciRows[c][5].toString().trim() : '';
        if (em) return em;
      }
    }
  }
  return '';
}

// ─────────────────────────────────────────────
// 참여확인서 HTML 반환 (admin에서 미리보기/다운로드용)
// 체크인 + GiftReceipts 모두 확인 후 생성
// ─────────────────────────────────────────────
function getCertificateHtmlForStudent(password, studentId) {
  var data = getCertificateData(password, studentId);
  if (!data.name)
    throw new Error('해당 학번의 참여 정보를 찾을 수 없습니다.');
  if (!data.checkIn)
    throw new Error('설명회 체크인 기록이 없습니다. 참여확인서를 발급할 수 없습니다.');
  if (!data.giftInfo)
    throw new Error('기념품 수령 기록이 없습니다. 참여확인서를 발급할 수 없습니다.');
  if (data.giftInfo.stampCount < BOOTH_STAMP_REQUIRED)
    throw new Error('부스 스탬프가 ' + BOOTH_STAMP_REQUIRED + '개 미만입니다.');
  // 부스 목록이 비어있으면 GiftReceipts 재조회 경고
  if (!data.giftInfo.booths || data.giftInfo.booths.trim() === '')
    Logger.log('[WARN] 학번 ' + studentId + ': GiftReceipts에 방문부스(E열) 데이터 없음');
  return _buildCertificateHtml(data);
}

// ─────────────────────────────────────────────
// 참여확인서 PDF 이메일 발송
// GAS에서 HTML → PDF 변환 후 첨부 발송
// ─────────────────────────────────────────────
function sendCertificateMail(password, studentId) {
  var data = getCertificateData(password, studentId);
  if (!data.name) throw new Error('해당 학번의 참여 정보를 찾을 수 없습니다.');
  if (!data.isIntercollege) throw new Error('참여확인서는 한양인터칼리지학부 학생에게만 발송됩니다.');
  if (!data.giftInfo) throw new Error('기념품 수령 기록이 없습니다. 참여확인서를 발행할 수 없습니다.');

  var email = _getStudentEmail(studentId);
  if (!email) throw new Error('이메일 주소를 찾을 수 없습니다. 학번을 확인해주세요.');

  var html = _buildCertificateHtml(data);

  // HTML → PDF 변환 (Google Drive 임시 파일 경유)
  var pdfBlob;
  try {
    var htmlBlob = Utilities.newBlob(html, 'text/html', 'cert.html');
    var tempFile = DriveApp.createFile(htmlBlob);
    var pdfFile  = tempFile.getAs('application/pdf');
    pdfBlob = pdfFile;
    pdfBlob.setName('융합전공소개행사_참여확인서_' + data.name + '.pdf');
    tempFile.setTrashed(true); // 임시 파일 삭제
  } catch(e) {
    // PDF 변환 실패 시 HTML 첨부로 fallback
    pdfBlob = Utilities.newBlob(html, 'text/html', '참여확인서_' + data.name + '.html');
  }

  GmailApp.sendEmail(email,
    '[한양YK인터칼리지] 2026 융합전공 소개행사 참여확인서',
    data.name + '님의 2026 융합전공 소개행사 참여확인서를 첨부합니다.\n\n한양YK인터칼리지 드림',
    {
      attachments: [pdfBlob],
      name:        '한양YK인터칼리지',
      replyTo:     'intercollege@hanyang.ac.kr',
      htmlBody:    '<p>' + data.name + '님의 2026 융합전공 소개행사 참여확인서를 첨부합니다.</p><p>한양YK인터칼리지 드림</p>'
    }
  );
  return email + ' 로 발송 완료';
}

// ─────────────────────────────────────────────
// 참여확인서 HTML 빌더 — 공문서 양식, 국영문 병기
// ─────────────────────────────────────────────
function _buildCertificateHtml(data) {
  var SESSION_TIMES = {
    '미래사회디자인':        '09:30–09:55',
    '융합의과학/융합의공학': '10:00–10:25',
    '인지융합과학':          '10:30–10:55',
    '혁신공학경영':          '11:00–11:25',
    '미래반도체공학':        '11:30–11:55'
  };
  var ALL_SESSIONS = ['미래사회디자인','융합의과학/융합의공학','인지융합과학','혁신공학경영','미래반도체공학'];
  var ALL_BOOTHS   = ['미래반도체공학','혁신공학경영','융합의과학/융합의공학','미래사회디자인','인지융합과학','라이프디자인센터'];

  function fmt(str) {
    if (!str) return '';
    return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  }
  function fmtDate(s) {
    if (!s) return '—';
    try { var d=new Date(s); if(isNaN(d)) return s;
      return d.getFullYear()+'. '+(d.getMonth()+1)+'. '+d.getDate()+'.  '+
             d.getHours().toString().padStart(2,'0')+':'+d.getMinutes().toString().padStart(2,'0');
    } catch(e){ return s; }
  }

  var preNames   = data.attendedSessions || []; // CheckIns H열 기반 실제 참석 설명회
  // 부스 방문 목록: GiftReceipts.booths 우선 (gift.html에서 직접 체크한 데이터)
  var stampNames = [];
  if (data.giftInfo && data.giftInfo.booths) {
    stampNames = data.giftInfo.booths.split(',').map(function(s){ return s.trim(); }).filter(Boolean);
  } else if (data.stamps && data.stamps.length) {
    stampNames = data.stamps.map(function(s){ return s.boothName; });
  }
  var verifyCode = data.giftInfo ? data.giftInfo.verifyCode : '';
  var issueDate  = '2026년 5월 8일 / May 8, 2026';
  var logoUrl    = 'https://drive.google.com/thumbnail?id=1gAd8LC2fRzr7reuiJV2FHSQCbqbz5DoR&sz=w300';
  var sealUrl    = data.sealImageUrl || '';
  var checkInTime = data.checkIn ? fmtDate(data.checkIn.checkedAt) : '—';
  var bonusSubject = (data.giftInfo && data.giftInfo.bonusSubject) ? data.giftInfo.bonusSubject : '';

  // ── 설명회 행 ──
  var sessRows = ALL_SESSIONS.map(function(s) {
    var ok = preNames.indexOf(s) !== -1;
    return '<tr>'+
      '<td class="tbl-chk">'+(ok?'<span class="chk-y">●</span>':'<span class="chk-n">○</span>')+'</td>'+
      '<td class="tbl-prog">'+fmt(s)+'</td>'+
      '<td class="tbl-time">'+(SESSION_TIMES[s]||'')+'</td>'+
    '</tr>';
  }).join('');

  // ── 부스 행 ──
  var boothRows = ALL_BOOTHS.map(function(b) {
    var ok = stampNames.indexOf(b) !== -1;
    return '<tr>'+
      '<td class="tbl-chk">'+(ok?'<span class="chk-y">●</span>':'<span class="chk-n">○</span>')+'</td>'+
      '<td class="tbl-prog" colspan="2">'+fmt(b)+'</td>'+
    '</tr>';
  }).join('');

  // ── 직인 영역: 박스 없이 학장 이름/직위 텍스트만 표시 ──
  var sealBlock =
    '<div style="text-align:center;line-height:1.8;padding:8px 0;">'
    + '<div style="font-size:12px;font-weight:900;color:#1B263B;letter-spacing:-.01em;">한양YK인터칼리지학장</div>'
    + '<div style="font-size:9px;color:#7A8899;">Dean, Hanyang YK Intercollege</div>'
    + '</div>';

  var html = '<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">'
    +'<style>'
    +'*{box-sizing:border-box;margin:0;padding:0;}'
    +'body{font-family:"Malgun Gothic","맑은 고딕","Apple SD Gothic Neo","Noto Sans KR",sans-serif;'
    +'  background:#e8e8e8;padding:24px;}'
    +'@media print{body{background:#fff;padding:0;}}'
    +'.page{max-width:680px;margin:0 auto;background:#fff;}'
    /* 상단 컬러 바 */
    +'.top-bar{height:6px;background:#1B263B;}'
    /* 헤더 */
    +'.header{padding:22px 32px 16px;display:flex;justify-content:space-between;align-items:center;'
    +'  border-bottom:2px solid #1B263B;}'
    +'.header-left .doc-type{font-size:11px;color:#7A8899;letter-spacing:.12em;margin-bottom:4px;}'
    +'.header-left .doc-title{font-size:22px;font-weight:900;color:#1B263B;letter-spacing:-.02em;}'
    +'.header-left .doc-title-sub{font-size:13px;font-weight:700;color:#1B263B;margin-top:3px;letter-spacing:.02em;}'
    +'.header-left .doc-title-ko{font-size:11px;color:#7A8899;margin-top:2px;}'
    +'.header-right img{height:48px;width:auto;object-fit:contain;}'
    /* 서두 */
    +'.preamble{padding:14px 32px;background:#F7F8FA;border-bottom:1px solid #E0E4EA;'
    +'  font-size:12.5px;color:#333;line-height:1.8;}'
    +'.preamble .name-highlight{font-weight:900;color:#1B263B;font-size:14px;}'
    /* 본문 */
    +'.body{padding:20px 32px;}'
    +'.section{margin-bottom:18px;}'
    +'.sec-title{font-size:10px;font-weight:700;color:#1B263B;letter-spacing:.1em;text-transform:uppercase;'
    +'  border-left:3px solid #1B263B;padding-left:8px;margin-bottom:8px;}'
    +'.sec-title .sec-title-ko{font-size:12px;font-weight:900;display:block;letter-spacing:0;text-transform:none;}'
    /* 인적사항 테이블 */
    +'.info-table{width:100%;border-collapse:collapse;font-size:12px;}'
    +'.info-table td{padding:6px 10px;border:1px solid #D0D5DD;vertical-align:middle;}'
    +'.info-table .lbl{background:#F0F2F5;font-weight:700;color:#1B263B;width:90px;white-space:nowrap;}'
    +'.info-table .lbl-en{font-size:9px;color:#9AA5B4;font-weight:400;display:block;}'
    /* 체크 테이블 */
    +'.chk-table{width:100%;border-collapse:collapse;font-size:12px;}'
    +'.chk-table th{background:#1B263B;color:#fff;padding:6px 10px;font-size:10px;font-weight:600;text-align:left;}'
    +'.chk-table td{padding:5px 10px;border-bottom:1px solid #EAECEF;}'
    +'.tbl-chk{width:28px;text-align:center;}'
    +'.tbl-prog{}'
    +'.tbl-time{color:#7A8899;font-size:11px;width:80px;white-space:nowrap;}'
    +'.chk-y{color:#1B263B;font-size:13px;}'
    +'.chk-n{color:#CCC;font-size:13px;}'
    /* 신청과목 박스 */
    +'.subj-box{background:#EEF2F7;border-left:3px solid #1B263B;padding:8px 14px;'
    +'  font-size:13px;font-weight:700;color:#1B263B;border-radius:0 4px 4px 0;}'
    +'.subj-box .subj-en{font-size:10px;color:#7A8899;font-weight:400;margin-top:1px;}'
    /* 검증코드 */
    +'.verify{border:1px solid #9AA5B4;padding:7px 14px;text-align:center;margin-top:12px;border-radius:3px;}'
    +'.verify .v-label{font-size:9px;color:#9AA5B4;letter-spacing:.1em;margin-bottom:2px;}'
    +'.verify .v-code{font-size:13px;font-weight:700;color:#5A6778;letter-spacing:.15em;}'
    /* 하단 서명란 */
    +'.footer{padding:16px 32px 24px;border-top:1px solid #D0D5DD;'
    +'  display:flex;justify-content:space-between;align-items:flex-end;}'
    +'.footer-left{font-size:11px;color:#7A8899;line-height:1.8;}'
    +'.sign-block{text-align:center;}'
    +'.sign-block .sig-label{font-size:9px;color:#9AA5B4;margin-bottom:4px;letter-spacing:.05em;}'
    +'.sign-block img.sig-img{max-width:130px;max-height:40px;object-fit:contain;display:block;margin:0 auto;}'
    +'.issuer-block{text-align:right;}'
    +'</style></head><body>'
    +'<div class="page">'
      +'<div class="top-bar"></div>'
      /* 헤더 */
      +'<div class="header">'
        +'<div class="header-left">'
          +'<div class="doc-type">CERTIFICATE OF PARTICIPATION · 공식 확인서</div>'
          +'<div class="doc-title">참&nbsp;여&nbsp;확&nbsp;인&nbsp;서</div>'
          +'<div class="doc-title-sub">2026 HIC Majors Fair</div>'
          +'<div class="doc-title-ko">2026 융합전공 소개행사 — Convergence Universe : 6개 융합의 별을 연결하다</div>'
        +'</div>'
        +'<div class="header-right"><img src="'+logoUrl+'" alt="YK Intercollege Logo"></div>'
      +'</div>'
      /* 서두 문장 */
      +'<div class="preamble">'
        +'위 학생은 아래와 같이 한양YK인터칼리지 주관 행사에 참여하였음을 확인합니다.<br>'
        +'<span style="font-size:11px;color:#7A8899;">'
        +'This is to certify that the student below has participated in the event organized by Hanyang YK Intercollege.'
        +'</span>'
      +'</div>'
      +'<div class="body">'
        /* 1. 참여자 정보 */
        +'<div class="section">'
          +'<div class="sec-title"><span class="sec-title-ko">1. 참여자 정보</span>Student Information</div>'
          +'<table class="info-table">'
            +'<tr>'
              +'<td class="lbl">성명<span class="lbl-en">Name</span></td>'
              +'<td style="font-weight:700;font-size:14px;">'+fmt(data.name)+'</td>'
              +'<td class="lbl">학번<span class="lbl-en">Student ID</span></td>'
              +'<td>'+fmt(data.studentId)+'</td>'
            +'</tr>'
            +'<tr>'
              +'<td class="lbl">소속<span class="lbl-en">Department</span></td>'
              +'<td colspan="3">'+fmt(data.dept)+'</td>'
            +'</tr>'
          +'</table>'
        +'</div>'
        /* 2. 행사 정보 */
        +'<div class="section">'
          +'<div class="sec-title"><span class="sec-title-ko">2. 행사 정보</span>Event Information</div>'
          +'<table class="info-table">'
            +'<tr>'
              +'<td class="lbl">행사명<span class="lbl-en">Event</span></td>'
              +'<td colspan="3">2026 융합전공 소개행사 — Convergence Universe</td>'
            +'</tr>'
            +'<tr>'
              +'<td class="lbl">일시<span class="lbl-en">Date &amp; Time</span></td>'
              +'<td>2026. 5. 8.(금) 09:30–17:00</td>'
              +'<td class="lbl">장소<span class="lbl-en">Venue</span></td>'
              +'<td>HIT 1F, 양민용 커리어라운지</td>'
            +'</tr>'
            +'<tr>'
              +'<td class="lbl">참석유형<span class="lbl-en">Type</span></td>'
              +'<td>'+(data.checkIn ? fmt(data.checkIn.type) : '—')+'</td>'
              +'<td class="lbl">체크인<span class="lbl-en">Check-in</span></td>'
              +'<td>'+checkInTime+'</td>'
            +'</tr>'
          +'</table>'
        +'</div>'
        /* 3. 설명회 참여 */
        +'<div class="section">'
          +'<div class="sec-title"><span class="sec-title-ko">3. 설명회 참여 현황</span>Session Attendance</div>'
          +'<table class="chk-table">'
            +'<thead><tr>'
              +'<th class="tbl-chk">참여<br><span style="font-size:9px;font-weight:400;">Attend</span></th>'
              +'<th>프로그램 / Program</th>'
              +'<th class="tbl-time">시간 / Time</th>'
            +'</tr></thead>'
            +'<tbody>'+sessRows+'</tbody>'
          +'</table>'
        +'</div>'
        /* 4. 부스 방문 */
        +'<div class="section">'
          +'<div class="sec-title"><span class="sec-title-ko">4. 부스 방문 현황</span>Booth Visits&nbsp;&nbsp;'
          +'<span style="font-size:10px;font-weight:400;color:#7A8899;">('+stampNames.length+' / 6개)</span></div>'
          +'<table class="chk-table">'
            +'<thead><tr>'
              +'<th class="tbl-chk">방문<br><span style="font-size:9px;font-weight:400;">Visit</span></th>'
              +'<th colspan="2">부스 / Booth</th>'
            +'</tr></thead>'
            +'<tbody>'+boothRows+'</tbody>'
          +'</table>'
        +'</div>'
        /* 5. 신청과목 (해당자만) */
        +(bonusSubject
          ? '<div class="section">'
              +'<div class="sec-title"><span class="sec-title-ko">5. 신청 과목</span>Applied Course</div>'
              +'<div class="subj-box">'+fmt(bonusSubject)
              +'<div class="subj-en">Course applied for academic credit incentive</div></div>'
            +'</div>'
          : '')
        /* 검증코드 */
        +(verifyCode
          ? '<div class="verify">'
              +'<div class="v-label">VERIFICATION CODE · 진위 확인 코드</div>'
              +'<div class="v-code">'+fmt(verifyCode)+'</div>'
            +'</div>'
          : '')
      +'</div>'/* /body */
      /* 푸터 */
      +'<div class="footer">'
        +'<div class="footer-left">'
          +'발급일&nbsp;/&nbsp;Date of Issue : '+issueDate+'<br>'
          +'발급기관 : 한양대학교 한양YK인터칼리지<br>'
          +'<span style="font-size:10px;">Hanyang YK Intercollege, Hanyang University</span>'
        +'</div>'
        // 직인만 표시 (참여자 서명 생략)
        +'<div style="display:flex;align-items:flex-end;gap:24px;">'
          + sealBlock
        +'</div>'
      +'</div>'
    +'</div>'/* /page */
    +'</body></html>';

  return html;
}


// ─────────────────────────────────────────────
// 인터칼리지 학생 전체 목록 조회 (admin 일괄 발송용)
// GiftReceipts 기준으로 인터칼리지 학생만
// ─────────────────────────────────────────────
function getIntercollegeStudentList(password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var admins = ss.getSheetByName('AdminUsers').getDataRange().getValues();
  var authorized = false;
  for (var i = 1; i < admins.length; i++) {
    if (admins[i][2].toString().trim() === '전체관리' &&
        admins[i][3].toString().trim() === password.toString().trim()) {
      authorized = true; break;
    }
  }
  if (!authorized) throw new Error('권한이 없습니다.');

  var sheet = ss.getSheetByName('GiftReceipts');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  // 체크인 학번 맵
  var ciMap2 = {};
  var ciSheet2 = ss.getSheetByName('CheckIns');
  if (ciSheet2 && ciSheet2.getLastRow() > 1) {
    var ciRows2 = ciSheet2.getDataRange().getValues();
    for (var ci2 = 1; ci2 < ciRows2.length; ci2++) {
      var cs2 = ciRows2[ci2][1] ? ciRows2[ci2][1].toString().trim() : '';
      if (cs2) ciMap2[cs2] = true;
    }
  }
  var rows = sheet.getDataRange().getValues();

  // 1패스: 학번 기준 전체 집계
  var byStudentIC = {};
  for (var j = 1; j < rows.length; j++) {
    var sid_ic  = rows[j][1] ? rows[j][1].toString().trim() : '';
    var dept_ic = rows[j][3] ? rows[j][3].toString().trim() : '';
    if (!sid_ic || dept_ic !== INTERCOLLEGE_DEPT) continue;
    if (!byStudentIC[sid_ic]) {
      byStudentIC[sid_ic] = {
        name:         rows[j][2] ? rows[j][2].toString().trim() : '',
        dept:         dept_ic,
        stampCount:   rows[j][5] ? Number(rows[j][5]) : 0,
        bonusSubject: rows[j][6] ? rows[j][6].toString().trim() : '',
        verifyCode:   rows[j][8] ? rows[j][8].toString().trim() : '',
        boothList:    []
      };
    }
    var booth_ic = rows[j][4] ? rows[j][4].toString().trim() : '';
    if (booth_ic) byStudentIC[sid_ic].boothList.push(booth_ic);
  }

  // 2패스: 자격 검증 후 결과 생성
  var result = [];
  Object.keys(byStudentIC).forEach(function(sid_ic) {
    var s = byStudentIC[sid_ic];
    if (!ciMap2[sid_ic] || s.stampCount < BOOTH_STAMP_REQUIRED) return;
    result.push({
      no:           0,
      studentId:    sid_ic,
      name:         s.name,
      dept:         s.dept,
      boothList:    s.boothList,
      booths:       s.boothList.join(', '),
      stampCount:   s.stampCount,
      bonusSubject: s.bonusSubject,
      verifyCode:   s.verifyCode
    });
  });
  result.forEach(function(r, i) { r.no = i + 1; });
  return result;
}