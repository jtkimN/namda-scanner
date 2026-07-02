/**
 * NAMDA 셀프 체크인 — 구글 백엔드 (Google Apps Script)
 * =====================================================
 * 역할 3가지
 *  (1) 구글폼 제출 시 자동 실행 → 참가자에게 "고정 고유번호 + 보안토큰" 부여, 시트 기록
 *  (2) 뿌리오 알림톡으로 개인 QR 링크 발송
 *  (3) 스캐너(index.html)의 체크인 요청(JSONP)을 받아 입장 처리 (2대 동시 안전)
 *
 * 설계 원칙 (지난번 오류 원천 차단)
 *  · 번호는 "행 위치"가 아니라 "값"으로 조회 → 시트 정렬돼도 엉뚱한 사람 안 뜸
 *  · LockService 로 동시쓰기 직렬화 → 스캐너 2대 동시에도 데이터 안 깨짐
 *  · 알림톡 발송(네트워크)은 락 밖에서 처리 → 현장 대량 동시등록에도 번호부여·체크인이 안 밀림
 *  · 토큰 검증 → 위조/타인 QR 차단
 *  · JSONP 반환 → 브라우저 CORS/리다이렉트 문제 회피 (지난번 "처리 오류" 주범)
 *  · 열은 "머리글 이름"으로 찾음 → 열 순서 바뀌어도 안전
 *
 * ⚠️ 비밀값(API키·스태프키)은 코드에 쓰지 말고 [프로젝트 설정 > 스크립트 속성]에 저장합니다.
 *    setup() 실행 시 필요한 속성 목록을 로그로 안내합니다.
 */

// ===== 시트 머리글 이름 =====
// NAME/PHONE 은 "포함" 매칭으로 찾습니다. 폼 질문 제목이 "성함을 입력해주세요." 처럼 길어도 인식됩니다.
var COL = {
  NAME:    '성함',       // 아래 NAME_KEYS 중 하나를 포함하는 열
  PHONE:   '전화번호',   // 아래 PHONE_KEYS 중 하나를 포함하는 열
  ID:      '번호',       // 스크립트가 추가 (정확 일치)
  TOKEN:   '토큰',       // 스크립트가 추가 (정확 일치)
  CHECKED: '입장여부',   // 스크립트가 추가 (정확 일치)
  TIME:    '입장시각',   // 스크립트가 추가 (정확 일치)
  SEND:    '발송상태'    // 스크립트가 추가 (정확 일치)
};
var NAME_KEYS  = ['성함', '이름'];              // 이름 열 판별 키워드
var PHONE_KEYS = ['전화', '휴대', '핸드폰'];    // 전화 열 판별 키워드
var ADDED_COLS = [COL.ID, COL.TOKEN, COL.CHECKED, COL.TIME, COL.SEND];

// ===== 설정 읽기 (스크립트 속성) =====
function cfg_(){
  var p = PropertiesService.getScriptProperties();
  return {
    ppurioBase:   p.getProperty('PPURIO_BASE') || 'https://message.ppurio.com',
    ppurioAcct:   p.getProperty('PPURIO_ACCOUNT') || '',
    ppurioKey:    p.getProperty('PPURIO_APIKEY') || '',
    sender:       p.getProperty('PPURIO_SENDER') || '',      // 발신프로필키
    template:     p.getProperty('PPURIO_TEMPLATE') || '',    // 알림톡 템플릿코드
    fromNumber:   p.getProperty('PPURIO_FROM') || '',        // (선택) 실패시 문자 대체 발신번호
    staffKey:     p.getProperty('STAFF_KEY') || '',
    qrBase:       p.getProperty('QR_BASE') || 'https://checkin.namda.site/qr.html',
    eventName:    p.getProperty('EVENT_NAME') || '보이스피싱 방지 인증기술 세미나 및 시연회',
    smsFallback:  (p.getProperty('SMS_FALLBACK') || 'N') === 'Y'
  };
}

// =====================================================
//  최초 1회 실행: 열 추가 + 트리거 설치 + 설정 점검
// =====================================================
function setup(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = responseSheet_(); // 폼 응답 탭 자동 선택
  ensureColumns_(sheet);

  // 폼 제출 트리거 (중복 방지: 기존 삭제 후 생성)
  var triggers = ScriptApp.getProjectTriggers();
  for (var i=0;i<triggers.length;i++){
    if (triggers[i].getHandlerFunction()==='onFormSubmit') ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();

  var c = cfg_();
  var missing = [];
  ['ppurioAcct','ppurioKey','sender','template','staffKey'].forEach(function(k){
    if (!c[k]) missing.push(k);
  });
  Logger.log('=== NAMDA 체크인 setup 완료 ===');
  Logger.log('추가된 열: ' + ADDED_COLS.join(', '));
  Logger.log('QR 링크 기본주소(QR_BASE): ' + c.qrBase);
  if (missing.length){
    Logger.log('⚠️ [스크립트 속성] 에 아래 값을 채워야 알림톡 발송·체크인이 됩니다:');
    Logger.log('   PPURIO_ACCOUNT / PPURIO_APIKEY / PPURIO_SENDER(발신프로필키) / PPURIO_TEMPLATE / STAFF_KEY');
  } else {
    Logger.log('설정값 모두 존재 ✅');
  }
}

// 이미 접수된(트리거 이전) 응답들에 번호·토큰 소급 부여 + 알림톡 발송
function assignExisting(sendAlimtalk){
  var sheet = responseSheet_();
  var hm, rows = [];

  // 1) 락 안: 번호·토큰 일괄 부여 (짧게, 네트워크 없음)
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    ensureColumns_(sheet);
    hm = headerMap_(sheet);
    var last = sheet.getLastRow();
    if (last < 2) { Logger.log('데이터 없음'); lock.releaseLock(); return; }
    for (var r=2; r<=last; r++){
      var id = String(sheet.getRange(r, hm[COL.ID]).getValue()).trim();
      if (!id){ id = nextId_(sheet, hm); sheet.getRange(r, hm[COL.ID]).setValue(id); }
      var tok = String(sheet.getRange(r, hm[COL.TOKEN]).getValue()).trim();
      if (!tok){ tok = randomToken_(); sheet.getRange(r, hm[COL.TOKEN]).setValue(tok); }
      rows.push({ r:r, id:id, tok:tok });
    }
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  Logger.log('번호 부여 ' + rows.length + '건 완료');

  // 2) 락 밖: 알림톡 발송 (rate-limit 여유)
  if (!sendAlimtalk) return;
  var sent = 0;
  for (var i=0;i<rows.length;i++){
    var row = rows[i].r;
    var name = String(sheet.getRange(row, hm[COL.NAME]).getValue()).trim();
    var phone = String(sheet.getRange(row, hm[COL.PHONE]).getValue()).trim();
    var res = sendAlimtalk_(name, phone, rows[i].id, rows[i].tok);
    sheet.getRange(row, hm[COL.SEND]).setValue(res.ok ? ('발송 ' + nowStr_()) : ('실패: ' + res.msg));
    if (res.ok) sent++;
    Utilities.sleep(120);
  }
  Logger.log('알림톡 발송 ' + sent + '/' + rows.length + '건');
}

// =====================================================
//  (1)(2) 폼 제출 트리거
// =====================================================
function onFormSubmit(e){
  var sheet = e.range.getSheet();
  var row = e.range.getRow();
  var hm, id, token, name, phone;

  // 1) 락 안: 번호·토큰 부여 + 기록만 (짧게 — 네트워크 없음)
  //    현장에서 수십 명이 동시 제출해도 락 점유가 0.1초라 번호 누락·체크인 밀림이 없음.
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    ensureColumns_(sheet);
    hm = headerMap_(sheet);
    name  = String(sheet.getRange(row, hm[COL.NAME]).getValue()).trim();
    phone = String(sheet.getRange(row, hm[COL.PHONE]).getValue()).trim();
    id = String(sheet.getRange(row, hm[COL.ID]).getValue()).trim();
    if (!id){ id = nextId_(sheet, hm); sheet.getRange(row, hm[COL.ID]).setValue(id); }
    token = String(sheet.getRange(row, hm[COL.TOKEN]).getValue()).trim();
    if (!token){ token = randomToken_(); sheet.getRange(row, hm[COL.TOKEN]).setValue(token); }
    SpreadsheetApp.flush();
  } catch (err){
    Logger.log('onFormSubmit(번호부여) 오류: ' + err);
    try { lock.releaseLock(); } catch(e2){}
    return;
  }
  lock.releaseLock();

  // 2) 락 밖: 알림톡 발송(네트워크) + 상태 기록. 실패해도 번호·토큰은 이미 부여됨.
  try {
    var res = sendAlimtalk_(name, phone, id, token);
    sheet.getRange(row, hm[COL.SEND]).setValue(res.ok ? ('발송 ' + nowStr_()) : ('실패: ' + res.msg));
  } catch (err){
    Logger.log('onFormSubmit(발송) 오류: ' + err);
  }
}

// =====================================================
//  (3) 체크인 API — 스캐너가 JSONP 로 호출
//     GET ...?mode=checkin&id=005&t=토큰&key=스태프키&callback=cb
// =====================================================
function doGet(e){
  var p = (e && e.parameter) || {};
  var mode = p.mode || '';
  var cb = p.callback || '';

  if (mode !== 'checkin'){
    // 배포 확인용 상태 페이지
    return reply_({ status:'up', service:'NAMDA 체크인 API', time: nowStr_() }, cb);
  }

  var c = cfg_();
  if (!c.staffKey || p.key !== c.staffKey){
    return reply_({ status:'denied', message:'스태프 키 불일치' }, cb);
  }

  var id = String(p.id || '').trim();
  var token = String(p.t || '').trim();
  if (!id) return reply_({ status:'invalid', message:'번호 없음' }, cb);

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = responseSheet_();
    var hm = headerMap_(sheet);
    var last = sheet.getLastRow();
    if (last < 2) return reply_({ status:'invalid', message:'등록되지 않은 QR입니다', id:id }, cb);

    // 번호 "값"으로 행 찾기 (정렬돼도 안전)
    var ids = sheet.getRange(2, hm[COL.ID], last-1, 1).getValues();
    var rowIdx = -1;
    for (var i=0;i<ids.length;i++){
      if (String(ids[i][0]).trim() === id){ rowIdx = i + 2; break; }
    }
    if (rowIdx === -1) return reply_({ status:'invalid', message:'등록되지 않은 QR입니다', id:id }, cb);

    var name = String(sheet.getRange(rowIdx, hm[COL.NAME]).getValue()).trim();
    var stored = String(sheet.getRange(rowIdx, hm[COL.TOKEN]).getValue()).trim();

    // 토큰 검증 (위조/타인 QR 차단). 토큰이 비어있는 예외 행은 통과 허용.
    if (stored && token && stored !== token){
      return reply_({ status:'invalid', message:'QR이 일치하지 않습니다', id:id }, cb);
    }

    var checked = String(sheet.getRange(rowIdx, hm[COL.CHECKED]).getValue()).trim();
    var timeCell = sheet.getRange(rowIdx, hm[COL.TIME]);
    if (checked){
      return reply_({ status:'duplicate', name:name, id:id, time:String(timeCell.getValue()) }, cb);
    }

    var t = nowStr_();
    sheet.getRange(rowIdx, hm[COL.CHECKED]).setValue('입장');
    timeCell.setValue(t);
    SpreadsheetApp.flush();
    return reply_({ status:'ok', name:name, id:id, time:t }, cb);
  } catch (err){
    return reply_({ status:'error', message:String(err) }, cb);
  } finally {
    lock.releaseLock();
  }
}

// =====================================================
//  뿌리오 알림톡 (message.ppurio.com API)
// =====================================================
function getPpurioToken_(){
  var c = cfg_();
  var cache = CacheService.getScriptCache();
  var cached = cache.get('ppurio_token');
  if (cached) return cached;

  var basic = Utilities.base64Encode(c.ppurioAcct + ':' + c.ppurioKey);
  var res = UrlFetchApp.fetch(c.ppurioBase + '/v1/token', {
    method: 'post',
    headers: { 'Authorization': 'Basic ' + basic },
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  var body = res.getContentText();
  if (code !== 200) throw new Error('토큰 발급 실패 (' + code + '): ' + body);
  var json = JSON.parse(body);
  if (!json.token) throw new Error('토큰 응답에 token 없음: ' + body);
  cache.put('ppurio_token', json.token, 60 * 60 * 6); // 6시간 캐시 (토큰 유효 24h)
  return json.token;
}

/**
 * 알림톡 1건 발송.
 * 뿌리오 템플릿 변수는 [*1*]~[*8*] 형식. changeWord 의 var1..var8 이 각각 [*1*]..[*8*] 을 채웁니다.
 * ⭐ 링크 길이(변수 50자) 문제 회피: 템플릿 본문에 고정 URL을 넣고( ...qr.html?[*2*] ),
 *    [*2*] 에는 "id=..&t=.." 짧은 쿼리만 넣습니다. → var1=성함, var2=쿼리
 */
function sendAlimtalk_(name, phone, id, token){
  var c = cfg_();
  if (!c.ppurioAcct || !c.ppurioKey || !c.sender || !c.template){
    return { ok:false, msg:'뿌리오 설정 미완료(스크립트 속성 확인)' };
  }
  var to = String(phone).replace(/[^0-9]/g, '');
  if (!to){ return { ok:false, msg:'전화번호 없음' }; }

  var query = qrQuery_(id, token);                 // [*2*] = id=..&t=..  (짧게)
  var fullLink = buildLink_(c.qrBase, id, token);  // 문자 대체발송용 전체 URL

  try {
    var accessToken = getPpurioToken_();
    var payload = {
      account: c.ppurioAcct,
      messageType: 'ALT',                 // 알림톡 텍스트(ALT). 이미지형 템플릿이면 'ALI'
      senderProfile: c.sender,            // 발신프로필
      templateCode: c.template,
      duplicateFlag: 'N',                 // 수신번호 중복 제거
      targetCount: 1,
      targets: [{
        to: to,
        name: name,
        changeWord: { var1: name, var2: query }   // [*1*]=성함, [*2*]=쿼리(id=..&t=..)
      }],
      refKey: ('namda' + Utilities.getUuid().replace(/-/g, '')).substring(0, 32), // 32자 이내 고유값
      isResend: c.smsFallback ? 'Y' : 'N'
    };
    if (c.smsFallback && c.fromNumber){
      // 알림톡 실패 시 문자 대체발송. URL 포함이라 LMS. 발신번호(PPURIO_FROM)는 뿌리오 사전 등록 필요.
      payload.resend = {
        messageType: 'LMS',
        from: c.fromNumber.replace(/[^0-9]/g, ''),
        subject: '입장 QR 안내',
        content: name + '님 입장 QR: ' + fullLink
      };
    }
    var res = UrlFetchApp.fetch(c.ppurioBase + '/v1/kakao', {
      method: 'post',
      contentType: 'application/json;charset=UTF-8',
      headers: { 'Authorization': 'Bearer ' + accessToken },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    var body = res.getContentText();
    var j = {};
    try { j = JSON.parse(body); } catch(e){}
    if (code === 200 && (j.code === '1000' || /success/i.test(j.description||''))){
      return { ok:true, msg:'ok', key:j.messageKey };
    }
    return { ok:false, msg:'(' + code + ') ' + (j.description || body) };
  } catch (err){
    return { ok:false, msg:String(err) };
  }
}

// 발송 1건 시험 (스크립트 속성 세팅 후 본인 번호로 테스트)
function testSend(){
  var res = sendAlimtalk_('테스트', '01000000000', '999', 'testtok0'); // ← 전화번호를 본인 번호로 바꿔 실행
  Logger.log(JSON.stringify(res));
}

// =====================================================
//  유틸
// =====================================================
function buildLink_(base, id, token){
  return String(base).replace(/\/+$/,'') + '?' + qrQuery_(id, token);
}

// 알림톡 [*2*] 용 짧은 쿼리 (id=..&t=..). 템플릿의 고정 URL 뒤에 붙습니다.
function qrQuery_(id, token){
  return 'id=' + encodeURIComponent(id) + '&t=' + encodeURIComponent(token);
}

function nextId_(sheet, hm){
  var last = sheet.getLastRow();
  var maxId = 0;
  if (last >= 2){
    var vals = sheet.getRange(2, hm[COL.ID], last-1, 1).getValues();
    for (var i=0;i<vals.length;i++){
      var n = parseInt(String(vals[i][0]).trim(), 10);
      if (!isNaN(n) && n > maxId) maxId = n;
    }
  }
  return pad3_(maxId + 1);
}

function pad3_(n){ n = String(n); while (n.length < 3) n = '0' + n; return n; }
function randomToken_(){ return Utilities.getUuid().replace(/-/g,'').substring(0,8); }

function nowStr_(){
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
}
function nowStamp_(){
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMddHHmmssSSS');
}

// 머리글 → 열번호(1-based) 매핑
// - 추가 관리열(번호/토큰/입장여부/입장시각/발송상태): 정확 일치
// - 이름/전화: 키워드 "포함" 매칭 (폼 질문 제목이 길어도 인식)
function headerMap_(sheet){
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var exact = {};
  for (var i=0;i<headers.length;i++){
    var h = String(headers[i]).trim();
    if (h) exact[h] = i + 1;
  }
  var map = {};
  [COL.ID, COL.TOKEN, COL.CHECKED, COL.TIME, COL.SEND].forEach(function(k){
    if (exact[k]) map[k] = exact[k];
  });
  map[COL.NAME]  = findHeaderCol_(headers, NAME_KEYS);
  map[COL.PHONE] = findHeaderCol_(headers, PHONE_KEYS);

  // 필수 열 검증
  [COL.NAME, COL.PHONE, COL.ID, COL.TOKEN, COL.CHECKED, COL.TIME].forEach(function(k){
    if (!map[k]) throw new Error('시트에서 "' + k + '" 에 해당하는 열을 찾지 못했습니다. '
      + '(setup() 실행 여부 / 폼 질문 제목 확인 — 이름 열엔 "성함" 또는 "이름", 전화 열엔 "전화/휴대"가 들어가야 함)');
  });
  return map;
}

// headers 배열에서 keys 중 하나를 "포함"하는 첫 열의 1-based 인덱스 (없으면 0)
function findHeaderCol_(headers, keys){
  for (var k=0;k<keys.length;k++){
    for (var i=0;i<headers.length;i++){
      if (String(headers[i]).indexOf(keys[k]) !== -1) return i + 1;
    }
  }
  return 0;
}

// 폼 응답 탭 자동 선택: "이름(성함/이름)" 열이 있는 첫 탭. (응답 탭이 첫 번째가 아니어도 안전)
function responseSheet_(){
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var s=0;s<sheets.length;s++){
    var lc = sheets[s].getLastColumn();
    if (lc < 1) continue;
    var headers = sheets[s].getRange(1, 1, 1, lc).getValues()[0];
    if (findHeaderCol_(headers, NAME_KEYS)) return sheets[s];
  }
  return sheets[0];
}

// 없는 관리 열은 오른쪽에 추가
function ensureColumns_(sheet){
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, Math.max(lastCol,1)).getValues()[0];
  var have = {};
  for (var i=0;i<headers.length;i++){ var h=String(headers[i]).trim(); if(h) have[h]=true; }
  var col = lastCol;
  ADDED_COLS.forEach(function(name){
    if (!have[name]){ col++; sheet.getRange(1, col).setValue(name); }
  });
}

// 반환 (callback 있으면 JSONP, 없으면 JSON)
function reply_(obj, cb){
  var json = JSON.stringify(obj);
  if (cb){
    return ContentService.createTextOutput(cb + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
