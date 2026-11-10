/**
 * 드라마본부 회의 탭 — 편집 트리거 (기존 L열 진입일 기록 + 연동 드롭다운)
 *
 * 진입점은 기존과 같은 onEditInstalled 하나다. 이미 걸려 있는 설치형 트리거를 그대로 쓴다.
 *  1) 연동 드롭다운: G열(구분①)에 따라 H열(구분②)·I열(구분③), 제작일 때 W열(대분류)·X열(소분류)
 *     목록을 행마다 직접 건다. '목록' 탭 Z~BJ열 보조 영역(6행 ↔ 6행)은 행 이동·복사 때
 *     드롭다운과 짝이 어긋나므로(2026-10-21 확인: 약 30행) 쓰지 않고 이름 범위를 직접 참조한다.
 *  2) 현재단계 진입일: G열 또는 I열이 실제로 바뀌면 L열에 오늘 날짜(한국시간)를 기록한다. (기존 로직 그대로)
 *
 * 설치
 *  1) 확장 프로그램 > Apps Script 의 기존 코드를 이 파일 내용으로 바꾼다.
 *  2) ensureEditTrigger 를 한 번 실행한다. (onEditInstalled 트리거가 이미 있으면 그대로 둔다)
 *  3) rebuildAllDropdowns 를 한 번 실행해 지금 어긋난 행들을 바로잡는다.
 *  '목록' 탭 A~W열 목록을 고친 뒤에도 rebuildAllDropdowns 를 한 번 실행하면 된다.
 */

var DD = {
  SHEET: '드라마본부 회의',
  FIRST_ROW: 6,
  COL_G: 7,  // 구분① 단계구분
  COL_H: 8,  // 구분②
  COL_I: 9,  // 구분③ 세부 진행
  COL_L: 12, // 현재단계 진입일
  COL_W: 23, // 대분류 (제작만)
  COL_X: 24, // 소분류 (제작만)
};

function onEditInstalled(e) {
  if (!e || !e.range) return;
  if (e.range.getSheet().getName() !== DD.SHEET) return;

  handleDropdownEdit_(e);
  updateStageDate_(e);
}

/** onEditInstalled 설치형 트리거가 없으면 만든다. 중복 트리거는 하나만 남긴다. */
function ensureEditTrigger() {
  var ss = SpreadsheetApp.getActive();
  var found = false;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() !== 'onEditInstalled') return;
    if (found) ScriptApp.deleteTrigger(t);
    found = true;
  });
  if (!found) ScriptApp.newTrigger('onEditInstalled').forSpreadsheet(ss).onEdit().create();
}

// ───────────────────────── 현재단계 진입일 (기존 로직) ─────────────────────────

function updateStageDate_(e) {
  var LAST_ROW = 150;

  var sh = e.range.getSheet();
  var row = e.range.getRow();
  var col = e.range.getColumn();

  // 데이터 범위 확인
  if (row < DD.FIRST_ROW || row > LAST_ROW) return;

  // G열 또는 I열 수정 시에만 실행
  if (col !== DD.COL_G && col !== DD.COL_I) return;

  // 여러 셀 일괄 수정은 제외
  // 기존 진입일이 잘못 덮어써지는 것을 방지
  if (
    e.range.getNumRows() !== 1 ||
    e.range.getNumColumns() !== 1
  ) return;

  // 이전 값과 현재 값 비교
  var oldValue = e.oldValue == null
    ? ''
    : String(e.oldValue);

  var newValue = e.range.getDisplayValue();

  // 실제 값이 동일하면 날짜 유지
  if (oldValue === newValue) return;

  // G열 단계구분이 공란이면 날짜 미기록
  var stage = sh.getRange(row, DD.COL_G)
    .getDisplayValue().trim();

  if (stage === '') return;

  // 한국시간 기준 날짜 생성
  var ymd = Utilities.formatDate(
    new Date(),
    'Asia/Seoul',
    'yyyy-MM-dd'
  );

  var today = new Date(ymd + 'T00:00:00+09:00');

  // 실제 단계가 변경된 경우에만 L열 날짜 갱신
  sh.getRange(row, DD.COL_L)
    .setNumberFormat('yyyy-mm-dd')
    .setValue(today);
}

// ───────────────────────── 연동 드롭다운 ─────────────────────────

/** G나 W가 바뀐 행만 다시 건다. */
function handleDropdownEdit_(e) {
  var range = e.range;
  var sheet = range.getSheet();

  var r0 = Math.max(range.getRow(), DD.FIRST_ROW);
  var r1 = range.getLastRow();
  if (r1 < r0) return;

  var c0 = range.getColumn();
  var c1 = range.getLastColumn();
  var touchesG = c0 <= DD.COL_G && DD.COL_G <= c1;
  var touchesW = c0 <= DD.COL_W && DD.COL_W <= c1;
  if (!touchesG && !touchesW) return;

  // G나 W 한 열만 고친 경우에만 새 목록에 없는 하위 값을 비운다.
  // 행 전체 붙여넣기처럼 여러 열을 한꺼번에 넣은 경우는 넣은 값을 그대로 둔다.
  var singleCol = c0 === c1;
  applyRows_(sheet, r0, r1 - r0 + 1, {
    clearHI: singleCol && touchesG,
    clearWX: singleCol && (touchesG || touchesW),
  });
}

/** 전체 행을 다시 건다. 값은 지우지 않는다. */
function rebuildAllDropdowns() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(DD.SHEET);
  var last = sheet.getLastRow();
  if (last < DD.FIRST_ROW) return;
  var n = last - DD.FIRST_ROW + 1;
  var bad = applyRows_(sheet, DD.FIRST_ROW, n, {});
  ss.toast(
    n + '행 드롭다운을 다시 걸었습니다.' +
      (bad.length ? ' 목록에 없는 값: ' + bad.join(', ') : ''),
    '드롭다운',
    10
  );
}

function applyRows_(sheet, startRow, numRows, opts) {
  var lists = loadLists_();
  var gVals = sheet.getRange(startRow, DD.COL_G, numRows, 1).getValues();
  var hiRange = sheet.getRange(startRow, DD.COL_H, numRows, 2);
  var wxRange = sheet.getRange(startRow, DD.COL_W, numRows, 2);
  var hiVals = hiRange.getValues();
  var wxVals = wxRange.getValues();

  var hiRules = [];
  var wxRules = [];
  var bad = [];
  var hiChanged = false;
  var wxChanged = false;

  for (var i = 0; i < numRows; i++) {
    var row = startRow + i;
    var g = String(gVals[i][0]).trim();
    var isProd = g === '제작';

    var l2 = lists['구분2_' + g];
    var l3 = lists['구분3_' + g];
    var lw = isProd ? lists['대분류_제작'] : null;
    var lx = isProd ? lists['소분류_' + String(wxVals[i][0]).trim()] : null;

    hiRules.push([rule_(l2), rule_(l3)]);
    wxRules.push([rule_(lw), rule_(lx)]);

    // H, I, W, X 순서로 확인
    var checks = [
      [hiVals, 0, l2, opts.clearHI, 'H'],
      [hiVals, 1, l3, opts.clearHI, 'I'],
      [wxVals, 0, lw, opts.clearWX, 'W'],
      [wxVals, 1, lx, opts.clearWX, 'X'],
    ];
    for (var k = 0; k < checks.length; k++) {
      var vals = checks[k][0], j = checks[k][1], list = checks[k][2];
      var clear = checks[k][3], col = checks[k][4];
      var v = String(vals[i][j]).trim();
      if (v === '' || (list && list.values.indexOf(v) !== -1)) continue;
      if (clear) {
        vals[i][j] = '';
        if (col === 'H' || col === 'I') hiChanged = true; else wxChanged = true;
      } else if (g !== '') {
        bad.push(col + row);
      }
    }
  }

  hiRange.setDataValidations(hiRules);
  wxRange.setDataValidations(wxRules);
  if (hiChanged) hiRange.setValues(hiVals);
  if (wxChanged) wxRange.setValues(wxVals);
  return bad;
}

/** '목록' 탭 이름 범위를 { 이름: { range, values } } 로 읽는다. 시트 범위 이름('목록'!구분2_기획)도 같은 이름으로 취급. */
function loadLists_() {
  var out = {};
  SpreadsheetApp.getActive().getNamedRanges().forEach(function (nr) {
    var name = nr.getName().replace(/^'?[^'!]*'?!/, '');
    if (!/^(구분2_|구분3_|대분류_|소분류_)/.test(name) || out[name]) return;
    var range = nr.getRange();
    var values = range.getValues()
      .reduce(function (a, r) { return a.concat(r); }, [])
      .map(function (v) { return String(v).trim(); })
      .filter(function (v) { return v !== ''; });
    out[name] = { range: range, values: values };
  });
  return out;
}

/** 목록이 없으면 null(드롭다운 제거). 범위를 직접 참조하므로 '목록' 탭 값을 고치면 바로 반영된다. */
function rule_(list) {
  if (!list || !list.values.length) return null;
  return SpreadsheetApp.newDataValidation()
    .requireValueInRange(list.range, true)
    .setAllowInvalid(false)
    .build();
}
