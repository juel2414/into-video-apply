/**
 * 인투마케팅 영상분야 훈련생 지원서 접수
 *
 * 하는 일
 *  1) '지원자 목록' 시트에 한 줄씩 추가 (접수 시각, 접수 번호, 각 문항, PDF 링크)
 *  2) 지원자별 A4 지원서 PDF를 내 드라이브 '인투마케팅 영상분야 지원서' 폴더에 저장
 *
 * 설치
 *  - 구글 스프레드시트를 새로 만들고 [확장 프로그램 > Apps Script]에 이 코드를 붙여 넣는다.
 *  - [배포 > 새 배포 > 웹 앱] 실행 사용자: 나 / 액세스 권한: 모든 사용자
 *  - 나온 웹 앱 URL(…/exec)을 index.html 의 SCRIPT_URL 에 넣는다.
 */

var SHEET_NAME = '지원자 목록';
var FOLDER_NAME = '인투마케팅 영상분야 지원서';
var TIMEZONE = 'Asia/Seoul';
var AGE_BASE_DATE = '2026-11-11'; // 만 나이 판단 기준일: 교육 시작일
var MIN_AGE = 16;
var MAX_MOTIVATION = 800;

var HEADERS = [
  '접수 시각', '접수 번호', '이름', '연락처', '생년월일', '리뉴젠 입소 날짜',
  '영상 경험', '영상 링크', '지원 동기', '담당 사역자', '사역자에게 알림', '개인정보 동의', 'PDF 링크'
];

var EXPERIENCES = [
  '전혀 없어요',
  '스마트폰으로 가볍게 해봤어요',
  '편집 프로그램(캡컷, 프리미어 등)을 써봤어요',
  '직접 만든 작품이나 채널이 있어요'
];

function doGet() {
  return json_({ ok: true, service: 'into-video-apply' });
}

function doPost(e) {
  var data;
  try {
    data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'bad_json' });
  }

  // 숨김 칸이 채워졌으면 봇으로 보고 저장 없이 성공처럼 응답
  if (data.website) return json_({ ok: true });

  var problem = validate_(data);
  if (problem) return json_({ ok: false, error: problem });

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    return json_({ ok: false, error: 'busy' });
  }

  try {
    var now = new Date();
    var sheet = getSheet_();
    var receiptNo = nextReceiptNo_(sheet, now);
    var app = clean_(data);
    var pdf = createPdf_(app, receiptNo, now);

    sheet.appendRow([
      Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd HH:mm:ss'),
      receiptNo,
      app.name,
      app.phone,
      app.birth,
      app.renewgen,
      app.experience,
      app.videoLink,
      app.motivation,
      app.minister,
      app.notified ? '네, 알렸어요' : '',
      app.privacy ? '동의' : '',
      pdf.getUrl()
    ].map(safeCell_));
    SpreadsheetApp.flush();

    return json_({ ok: true, receiptNo: receiptNo });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'server' });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 검증 ---------- */

function validate_(d) {
  var str = function (v) { return typeof v === 'string' ? v.trim() : ''; };
  if (!str(d.name)) return 'name';
  if (!/^010-\d{4}-\d{4}$/.test(str(d.phone))) return 'phone';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(str(d.birth))) return 'birth';
  if (ageOn_(d.birth, AGE_BASE_DATE) < MIN_AGE) return 'age';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(str(d.entry))) return 'entry';
  if (EXPERIENCES.indexOf(d.exp) < 0) return 'exp';
  if (!str(d.motive)) return 'motive';
  if (!str(d.pastor)) return 'pastor';
  if (d.notified !== true) return 'notified';
  if (d.agree !== true) return 'agree';
  return '';
}

// 지원 페이지(index.html)의 입력 이름을 시트·PDF에서 쓰는 이름으로 옮긴다
function clean_(d) {
  var link = String(d.link || '').trim().slice(0, 500);
  if (link && !/^https?:\/\//i.test(link)) link = '';
  return {
    name: String(d.name).trim().slice(0, 30),
    phone: String(d.phone).trim(),
    birth: String(d.birth).trim(),
    renewgen: String(d.entry).trim(),
    experience: d.exp,
    videoLink: link,
    motivation: String(d.motive).trim().slice(0, MAX_MOTIVATION),
    minister: String(d.pastor).trim().slice(0, 30),
    notified: true,
    privacy: true
  };
}

function ageOn_(birthYMD, baseYMD) {
  var b = birthYMD.split('-').map(Number);
  var t = baseYMD.split('-').map(Number);
  var age = t[0] - b[0];
  if (t[1] < b[1] || (t[1] === b[1] && t[2] < b[2])) age--;
  return age;
}

/* ---------- 시트 ---------- */

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    ss = SpreadsheetApp.openById(id);
  }
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold').setBackground('#1c2752').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(9, 360);
  }
  return sheet;
}

// INTO-V-20261105-001 : 날짜별 순번
function nextReceiptNo_(sheet, now) {
  var prefix = 'INTO-V-' + Utilities.formatDate(now, TIMEZONE, 'yyyyMMdd') + '-';
  var last = sheet.getLastRow();
  var max = 0;
  if (last > 1) {
    var ids = sheet.getRange(2, 2, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      var v = String(ids[i][0]);
      if (v.indexOf(prefix) === 0) max = Math.max(max, Number(v.slice(prefix.length)) || 0);
    }
  }
  return prefix + ('00' + (max + 1)).slice(-3);
}

// 수식으로 해석될 수 있는 값(=, +, -, @ 시작)은 글자로 저장
function safeCell_(v) {
  if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
  return v;
}

/* ---------- PDF ---------- */

function getFolder_() {
  var it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function korDate_(ymd) {
  var p = ymd.split('-');
  return p[0] + '년 ' + Number(p[1]) + '월 ' + Number(p[2]) + '일';
}

function createPdf_(app, receiptNo, now) {
  var NAVY = '#1c2752';
  var LABEL_BG = '#eef0f6';
  var title = receiptNo + '_' + app.name + '_지원서';

  var doc = DocumentApp.create(title);
  var body = doc.getBody();
  body.setPageWidth(595.28).setPageHeight(841.89); // A4
  body.setMarginTop(50).setMarginBottom(50).setMarginLeft(56).setMarginRight(56);
  body.clear();

  var base = {};
  base[DocumentApp.Attribute.FONT_FAMILY] = 'Noto Sans KR';
  base[DocumentApp.Attribute.FONT_SIZE] = 10;
  base[DocumentApp.Attribute.FOREGROUND_COLOR] = '#1a1d29';
  body.setAttributes(base);

  var no = body.getParagraphs()[0];
  no.setText('접수 번호  ' + receiptNo)
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
    .editAsText().setFontSize(8).setForegroundColor('#6b7088');

  var h = body.appendParagraph('지 원 서');
  h.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(4).setSpacingAfter(4);
  h.editAsText().setFontSize(24).setBold(true).setForegroundColor(NAVY);

  var sub = body.appendParagraph('인투마케팅 영상분야 훈련생 · 배우고, 만들고, 수익까지 내는 영상 마케터 과정');
  sub.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(14);
  sub.editAsText().setFontSize(9).setBold(false).setForegroundColor('#6b7088');

  function section(text) {
    var p = body.appendParagraph(text);
    p.setSpacingBefore(12).setSpacingAfter(4);
    p.editAsText().setFontSize(11).setBold(true).setForegroundColor(NAVY);
  }

  function table(rows, labelWidth) {
    var t = body.appendTable(rows);
    t.setBorderColor('#9aa0b8').setBorderWidth(0.5);
    for (var r = 0; r < t.getNumRows(); r++) {
      var row = t.getRow(r);
      for (var c = 0; c < row.getNumCells(); c++) {
        var cell = row.getCell(c);
        cell.setPaddingTop(5).setPaddingBottom(5).setPaddingLeft(7).setPaddingRight(7);
        cell.editAsText().setFontSize(10).setBold(false).setForegroundColor('#1a1d29');
        if (c % 2 === 0 && row.getNumCells() > 1) {
          cell.setBackgroundColor(LABEL_BG);
          cell.editAsText().setBold(true).setForegroundColor(NAVY);
          if (labelWidth) cell.setWidth(labelWidth);
        }
      }
    }
    return t;
  }

  section('1. 인적사항');
  table([
    ['성명', app.name, '연락처', app.phone],
    ['생년월일', korDate_(app.birth), '리뉴젠 입소', korDate_(app.renewgen)]
  ], 70);

  section('2. 영상 경험');
  table([
    ['촬영·편집 경험', app.experience],
    ['영상 링크', app.videoLink || '없음']
  ], 90);

  section('3. 지원 동기');
  var mt = body.appendTable([[app.motivation]]);
  mt.setBorderColor('#9aa0b8').setBorderWidth(0.5);
  var mc = mt.getCell(0, 0);
  mc.setPaddingTop(8).setPaddingBottom(8).setPaddingLeft(9).setPaddingRight(9);
  mc.editAsText().setFontSize(10).setBold(false).setForegroundColor('#1a1d29');
  mc.getChild(0).asParagraph().setLineSpacing(1.5);

  section('4. 담당 사역자 확인');
  table([
    ['패밀리 담당 사역자', app.minister],
    ['지원 사실 알림', '☑ 네, 알렸어요']
  ], 110);

  section('5. 개인정보 수집·이용 동의');
  table([
    ['수집 항목', '이름, 연락처, 생년월일, 리뉴젠 입소 날짜, 담당 사역자'],
    ['이용 목적', '훈련생 선발 및 교육 안내'],
    ['보관 기간', '선발 종료 후 6개월'],
    ['동의 여부', '☑ 동의함']
  ], 90);

  var pledge = body.appendParagraph('위 기재 사항이 사실과 다름없음을 확인하며, 인투마케팅 영상분야 훈련생으로 지원합니다.');
  pledge.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(22);
  pledge.editAsText().setFontSize(10).setBold(false).setForegroundColor('#1a1d29');

  var date = body.appendParagraph(Utilities.formatDate(now, TIMEZONE, 'yyyy년 M월 d일'));
  date.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(14);
  date.editAsText().setFontSize(11).setBold(false);

  var sign = body.appendParagraph('지원자  ' + app.name + '   (서명)');
  sign.setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setSpacingBefore(16);
  sign.editAsText().setFontSize(11).setBold(false);

  var to = body.appendParagraph('인투마케팅 귀중');
  to.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(26);
  to.editAsText().setFontSize(15).setBold(true).setForegroundColor(NAVY);

  doc.saveAndClose();

  var docFile = DriveApp.getFileById(doc.getId());
  var pdf = getFolder_().createFile(docFile.getAs(MimeType.PDF).setName(title + '.pdf'));
  docFile.setTrashed(true); // PDF로 만든 뒤 임시 문서는 휴지통으로
  return pdf;
}

/* ---------- 공통 ---------- */

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 편집기에서 직접 실행해 권한 승인 + 시트/PDF 생성 확인용
function testSubmit() {
  var res = doPost({ postData: { contents: JSON.stringify({
    website: '',
    name: '테스트',
    phone: '010-1234-5678',
    birth: '2005-03-15',
    entry: '2024-02-01',
    exp: EXPERIENCES[2],
    link: 'https://youtube.com/@example',
    motive: '테스트 지원 동기입니다. 영상으로 메시지를 전하는 법을 배우고 싶어요.',
    pastor: '홍길동',
    notified: true,
    agree: true
  }) } });
  console.log(res.getContent());
}
