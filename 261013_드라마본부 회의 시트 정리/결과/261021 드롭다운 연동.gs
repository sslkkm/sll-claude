/**
 * 드라마본부 회의 탭 — 연동 드롭다운
 *
 * G열(구분①) 값에 따라 H열(구분②)·I열(구분③), 그리고 제작일 때 W열(대분류)·X열(소분류)의
 * 드롭다운 목록을 행마다 직접 건다.
 *
 * 기존 방식은 '목록' 탭 Z~BJ열 보조 영역(6행 ↔ 6행)을 거쳤는데, 행 이동·복사·삭제 때
 * 보조 영역의 참조와 드롭다운의 참조가 서로 다르게 움직여 행 짝이 어긋났다
 * (2026-10-21 확인: 약 30개 행이 다른 행의 G값을 기준으로 목록을 받고 있었음).
 * 이 스크립트는 보조 영역을 쓰지 않고 '목록' 탭의 이름 범위(구분2_기획, 구분3_제작 …)를
 * 직접 참조하므로 행을 옮겨도 그 행의 G·W 값 기준으로 다시 맞춰진다.
 *
 * 설치
 *  1) 확장 프로그램 > Apps Script 에 이 파일을 붙여넣는다.
 *  2) 기존 코드에 드롭다운(H·I·W·X)을 다시 거는 부분이 있으면 지운다. (둘이 서로 덮어쓴다)
 *  3) setupDropdownTrigger 를 한 번 실행한다. (권한 승인 필요)
 *  4) rebuildAllDropdowns 를 한 번 실행해 지금 어긋난 행들을 바로잡는다.
 *  '목록' 탭 A~W열 목록을 고친 뒤에도 rebuildAllDropdowns 를 한 번 실행하면 된다.
 */

const DD = {
  SHEET: '드라마본부 회의',
  FIRST_ROW: 6,
  COL_G: 7,  // 구분①
  COL_H: 8,  // 구분②
  COL_I: 9,  // 구분③
  COL_W: 23, // 대분류 (제작만)
  COL_X: 24, // 소분류 (제작만)
};

/** onEdit 설치형 트리거를 건다. 기존 onEdit 함수와 이름이 겹치지 않도록 별도 함수로 둔다. */
function setupDropdownTrigger() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'handleDropdownEdit')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('handleDropdownEdit').forSpreadsheet(ss).onEdit().create();
}

/** 편집 시: G나 W가 바뀐 행만 다시 건다. */
function handleDropdownEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== DD.SHEET) return;

  const r0 = Math.max(range.getRow(), DD.FIRST_ROW);
  const r1 = range.getLastRow();
  if (r1 < r0) return;

  const c0 = range.getColumn();
  const c1 = range.getLastColumn();
  const touchesG = c0 <= DD.COL_G && DD.COL_G <= c1;
  const touchesW = c0 <= DD.COL_W && DD.COL_W <= c1;
  if (!touchesG && !touchesW) return;

  // G나 W 한 열만 고친 경우에만 새 목록에 없는 하위 값을 비운다.
  // 행 전체 붙여넣기처럼 여러 열을 한꺼번에 넣은 경우는 넣은 값을 그대로 둔다.
  const singleCol = c0 === c1;
  applyRows_(sheet, r0, r1 - r0 + 1, {
    clearHI: singleCol && touchesG,
    clearWX: singleCol && (touchesG || touchesW),
  });
}

/** 전체 행을 다시 건다. 값은 지우지 않는다. */
function rebuildAllDropdowns() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(DD.SHEET);
  const last = sheet.getLastRow();
  if (last < DD.FIRST_ROW) return;
  const n = last - DD.FIRST_ROW + 1;
  const bad = applyRows_(sheet, DD.FIRST_ROW, n, {});
  ss.toast(
    n + '행 드롭다운을 다시 걸었습니다.' +
      (bad.length ? ' 목록에 없는 값: ' + bad.join(', ') : ''),
    '드롭다운',
    10
  );
}

function applyRows_(sheet, startRow, numRows, opts) {
  const lists = loadLists_();
  const gVals = sheet.getRange(startRow, DD.COL_G, numRows, 1).getValues();
  const hiRange = sheet.getRange(startRow, DD.COL_H, numRows, 2);
  const wxRange = sheet.getRange(startRow, DD.COL_W, numRows, 2);
  const hiVals = hiRange.getValues();
  const wxVals = wxRange.getValues();

  const hiRules = [];
  const wxRules = [];
  const bad = [];
  let hiChanged = false;
  let wxChanged = false;

  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    const g = String(gVals[i][0]).trim();
    const isProd = g === '제작';

    const l2 = lists['구분2_' + g];
    const l3 = lists['구분3_' + g];
    const lw = isProd ? lists['대분류_제작'] : null;
    const lx = isProd ? lists['소분류_' + String(wxVals[i][0]).trim()] : null;

    hiRules.push([rule_(l2), rule_(l3)]);
    wxRules.push([rule_(lw), rule_(lx)]);

    // H, I, W, X 순서로 확인
    [[hiVals, 0, l2, opts.clearHI, 'H'], [hiVals, 1, l3, opts.clearHI, 'I'],
     [wxVals, 0, lw, opts.clearWX, 'W'], [wxVals, 1, lx, opts.clearWX, 'X']]
      .forEach(([vals, j, list, clear, col]) => {
        const v = String(vals[i][j]).trim();
        if (v === '' || (list && list.values.includes(v))) return;
        if (clear) {
          vals[i][j] = '';
          if (col === 'H' || col === 'I') hiChanged = true; else wxChanged = true;
        } else if (g !== '') {
          bad.push(col + row);
        }
      });
  }

  hiRange.setDataValidations(hiRules);
  wxRange.setDataValidations(wxRules);
  if (hiChanged) hiRange.setValues(hiVals);
  if (wxChanged) wxRange.setValues(wxVals);
  return bad;
}

/** '목록' 탭 이름 범위를 { 이름: { range, values } } 로 읽는다. 시트 범위 이름('목록'!구분2_기획)도 같은 이름으로 취급. */
function loadLists_() {
  const out = {};
  SpreadsheetApp.getActive().getNamedRanges().forEach(nr => {
    const name = nr.getName().replace(/^'?[^'!]*'?!/, '');
    if (!/^(구분2_|구분3_|대분류_|소분류_)/.test(name) || out[name]) return;
    const range = nr.getRange();
    const values = range.getValues().flat().map(v => String(v).trim()).filter(v => v !== '');
    out[name] = { range, values };
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
