/**
 * 한문 출정기 기록 수신용 Apps Script
 * 구글 시트 → 확장 프로그램 → Apps Script 에 이 코드를 그대로 붙여 넣고,
 * 「웹 앱」으로 배포한 주소(…/exec)를 게임의 RECORD_URL 에 넣으면 됩니다.
 * 처음 기록이 들어오면 '기록', '오답', '요약', '많이 틀린 문제', '명렬' 시트가 자동으로 만들어집니다.
 * '명렬' 시트 A~D열(학년·반·번호·이름)에 학생 명단을 붙여 넣으면 E열부터 학생별 기록이 자동으로 채워집니다.
 */
const RECORD_HEAD = ['받은 시각','학년','반','번호','이름','결과','완주(1)','점수','맞힌 수','푼 수','정답률','도달','물리친 적','최고 연속','출제 범위','플레이(초)','틀린 문제'];
const WRONG_HEAD = ['받은 시각','학년','반','번호','이름','틀린 문제'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    setup_(ss);
    const now = new Date();
    const who = [num_(d.grade), num_(d.cls), num_(d.num), clean_(d.name)];
    const answered = Number(d.answered) || 0, correct = Number(d.correct) || 0;
    const wrong = Array.isArray(d.wrong) ? d.wrong.map(clean_) : [];
    ss.getSheetByName('기록').appendRow([now, ...who, clean_(d.result), d.done ? 1 : 0, Number(d.score) || 0,
      correct, answered, answered ? correct / answered : '', clean_(d.reached), Number(d.kills) || 0,
      Number(d.maxCombo) || 0, clean_(d.cats), Number(d.seconds) || 0, wrong.join('\n')]);
    if (wrong.length) {
      const ws = ss.getSheetByName('오답');
      ws.getRange(ws.getLastRow() + 1, 1, wrong.length, WRONG_HEAD.length)
        .setValues(wrong.map(w => [now, ...who, w]));
    }
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

// 브라우저로 웹 앱 주소를 열었을 때 확인용 + 게임의 '오늘의 순위' (…/exec?action=rank)
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'rank') {
    return ContentService.createTextOutput(JSON.stringify(todayRank_()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput('한문 출정기 기록을 받을 준비가 되었습니다.');
}

// 오늘 기록만 모아 학생별 최고 점수 순으로. 이름은 가운데 글자를 *로 가려서 보냄
function todayRank_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const sh = ss.getSheetByName('기록');
  if (!sh || sh.getLastRow() < 2) return {date: today, plays: 0, players: []};
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, RECORD_HEAD.length).getValues();
  const best = {};
  let plays = 0;
  rows.forEach(r => {
    if (!(r[0] instanceof Date) || Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') !== today) return;
    plays++;
    const key = [r[1], r[2], r[3], r[4]].join('|');
    const b = best[key] || (best[key] = {grade: String(r[1]), cls: String(r[2]), num: String(r[3]),
      name: mask_(String(r[4])), score: -1, result: '', done: 0, plays: 0});
    b.plays++;
    if (Number(r[6]) === 1) b.done = 1;
    const score = Number(r[7]) || 0;
    if (score > b.score) { b.score = score; b.result = String(r[5]); }
  });
  const players = Object.keys(best).map(k => best[k]).sort((a, b) => b.score - a.score).slice(0, 200);
  return {date: today, plays: plays, players: players};
}

function mask_(n) {
  if (n.length <= 1) return n;
  if (n.length === 2) return n[0] + '*';
  return n[0] + '*'.repeat(n.length - 2) + n[n.length - 1];
}

// 학년·반·번호는 숫자로 저장 (명렬과 맞춰 보기 위해)
function num_(v) {
  const s = String(v == null ? '' : v).trim();
  return /^\d+$/.test(s) ? Number(s) : clean_(s);
}

// 수식으로 오인되지 않게 앞의 =,+,-,@ 를 막고 길이를 제한
function clean_(v) {
  const s = String(v == null ? '' : v).slice(0, 500);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function setup_(ss) {
  if (!ss.getSheetByName('기록')) {
    const sh = ss.insertSheet('기록', 0);
    sh.appendRow(RECORD_HEAD);
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('K:K').setNumberFormat('0%');
    sh.getRange(1, 1, 1, RECORD_HEAD.length).setFontWeight('bold').setBackground('#f6f1e7');
  }
  if (!ss.getSheetByName('오답')) {
    const sh = ss.insertSheet('오답');
    sh.appendRow(WRONG_HEAD);
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange(1, 1, 1, WRONG_HEAD.length).setFontWeight('bold').setBackground('#f6f1e7');
  }
  if (!ss.getSheetByName('요약')) {
    const sh = ss.insertSheet('요약');
    sh.getRange('A1').setFormula(
      "=QUERY('기록'!A2:Q, \"select B, C, D, E, count(A), sum(G), max(H), avg(K), max(A) " +
      "where E is not null group by B, C, D, E order by B, C, D " +
      "label B '학년', C '반', D '번호', E '이름', count(A) '플레이 횟수', sum(G) '완주 횟수', " +
      "max(H) '최고 점수', avg(K) '평균 정답률', max(A) '마지막 플레이'\", 0)");
    sh.getRange('H:H').setNumberFormat('0%');
    sh.getRange('I:I').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.setFrozenRows(1);
  }
  if (!ss.getSheetByName('명렬')) {
    // A~D: 선생님이 붙여 넣는 명단 / E~J: 기록 시트에서 자동 계산 (학년·반·번호로 맞춤)
    const sh = ss.insertSheet('명렬');
    sh.getRange(1, 1, 1, 10).setValues([['학년', '반', '번호', '이름', '플레이 횟수', '완주 횟수', '최고 점수', '평균 정답률', '오늘 최고 점수', '마지막 플레이']])
      .setFontWeight('bold').setBackground('#f6f1e7');
    sh.getRange(1, 5, 1, 6).setBackground('#e6efe9');
    const key = "'기록'!B:B,g,'기록'!C:C,c,'기록'!D:D,n";
    const f = (body) => '=MAP(A2:A,B2:B,C2:C,LAMBDA(g,c,n,IF(n="","",' + body + ')))';
    sh.getRange('E2').setFormula(f('COUNTIFS(' + key + ')'));
    sh.getRange('F2').setFormula(f('SUMIFS(\'기록\'!G:G,' + key + ')'));
    sh.getRange('G2').setFormula(f('IFERROR(1/(1/MAXIFS(\'기록\'!H:H,' + key + ')),"-")'));
    sh.getRange('H2').setFormula(f('IFERROR(AVERAGEIFS(\'기록\'!K:K,' + key + '),"-")'));
    sh.getRange('I2').setFormula(f('IFERROR(1/(1/MAXIFS(\'기록\'!H:H,' + key + ",'기록'!A:A,\">=\"&TODAY())),\"-\")"));
    sh.getRange('J2').setFormula(f('IFERROR(1/(1/MAXIFS(\'기록\'!A:A,' + key + ')),"-")'));
    sh.getRange('H:H').setNumberFormat('0%');
    sh.getRange('J:J').setNumberFormat('mm-dd hh:mm');
    sh.setFrozenRows(1);
  }
  if (!ss.getSheetByName('많이 틀린 문제')) {
    const sh = ss.insertSheet('많이 틀린 문제');
    sh.getRange('A1').setFormula(
      "=QUERY('오답'!F2:F, \"select F, count(F) where F is not null group by F order by count(F) desc " +
      "label F '문제', count(F) '틀린 횟수'\", 0)");
    sh.setColumnWidth(1, 520);
    sh.setFrozenRows(1);
  }
}

// 편집기에서 한 번 실행하면 시트 4개를 미리 만들어 둡니다 (선택).
function 시트_준비() {
  setup_(SpreadsheetApp.getActiveSpreadsheet());
}
