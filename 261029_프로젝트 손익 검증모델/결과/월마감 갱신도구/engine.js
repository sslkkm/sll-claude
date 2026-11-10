// 프로젝트 손익 계산 엔진 (9월본 수식 구조를 따름). 금액 단위 억원
// 월 index 0 = 2025-01, 107 = 2033-12
const N_MONTH = 108;
const idxOf = (y, m) => (y - 2025) * 12 + (m - 1);
const ymOf = i => ({ y: 2025 + Math.floor(i / 12), m: (i % 12) + 1 });
const keyOf = i => { const { y, m } = ymOf(i); return y + '-' + String(m).padStart(2, '0'); };

// 유통매출 누적 경험률 (손상 시험)
const EXP_JH = [0.08422232, 0.34868543, 0.55013644, 0.6516992, 0.74233293, 0.81978338, 0.86942677, 0.90162307, 0.93388664, 0.94848005, 0.95854635, 0.9640058, 0.97514994, 0.98411827, 0.98992135, 0.99894386, 1.01592789, 1];
const EXP_EX = [0.596860279201816, 0.8687679703470667, 0.9446420650164873, 0.955182712879788, 0.9599230676144511, 0.9634912600029675];
// 손상 시험 결산월. 원본(9월본)은 2026·2027년 블록만 있다. 라인업 시뮬레이션은 2029년까지 늘려 쓴다
const testMonths = until => { const a = []; for (let y = 2026; y <= until; y++) [3, 6, 9, 12].forEach(m => a.push(idxOf(y, m))); return a; };

function addMonths(d, n) { // d = [y,m,day]
  const t = d[0] * 12 + (d[1] - 1) + n;
  return [Math.floor(t / 12), (t % 12) + 1, d[2]];
}
function monthEnd(i) { const { y, m } = ymOf(i); return [y, m, new Date(y, m, 0).getDate()]; }
function cmpDate(a, b) { return (a[0] - b[0]) || (a[1] - b[1]) || (a[2] - b[2]); }
function datedifM(a, b) { // 엑셀 DATEDIF(a,b,"M")
  let n = (b[0] - a[0]) * 12 + (b[1] - a[1]);
  if (b[2] < a[2]) n -= 1;
  return Math.max(n, 0);
}
const r8 = x => Math.round(x * 1e8) / 1e8;

// 기준 월별 흐름을 배열로
function series(obj, shift, mul) {
  const a = new Array(N_MONTH).fill(0);
  for (const k in (obj || {})) {
    const [y, m] = k.split('-').map(Number);
    const i = idxOf(y, m) + shift;
    if (i >= 0 && i < N_MONTH) a[i] += obj[k] * mul;
  }
  return a;
}

// t = 타이틀 기준 정보, s = 시나리오, g = 공통 가정
function runTitle(t, s, g) {
  const shift = s.delay || 0;
  const start = s.start || addMonths(t.start, shift); // 라인업 화면은 방영일을 직접 넘긴다
  const startIdx = idxOf(start[0], start[1]);
  const type = s.type;
  // 편성: 월별 방영 회차. 9월본과 다르면 종영월도 마지막 방영월로 바뀐다
  const segs = s.segs || t.segs;
  const segChanged = segs.join(',') !== t.segs.join(',');
  const eps = segs.reduce((a, b) => a + b, 0);
  let end = s.end || addMonths(t.end, shift);
  if (segChanged && !s.end) {
    let lastK = 0; segs.forEach((n, k) => { if (n > 0) lastK = k; });
    const em = addMonths([start[0], start[1], 1], lastK);
    end = [em[0], em[1], Math.min(t.end[2], new Date(em[0], em[1], 0).getDate())];
  }
  const fx = 1 + t.usdShare * ((s.fx || t.baseFx) / t.baseFx - 1);
  const ovsMul = s.ovsMul * fx;

  const L = {};
  // 방영권료 = 회당 방영권료 × 그 달 방영 회차, 같은 금액이 제작매출원가
  L.방영권료 = new Array(N_MONTH).fill(0);
  segs.forEach((n, k) => { const i = startIdx + k; if (i < N_MONTH) L.방영권료[i] += s.rightsPerEp * n; });
  L.제작매출원가 = L.방영권료.slice();
  // 인식 기간: 1이면 원본처럼 그 달에 한 번에, n이면 그 달부터 n개월에 나눠 인식(가정)
  const spread = (a, n) => {
    if (!n || n <= 1) return a;
    const o = new Array(N_MONTH).fill(0);
    a.forEach((v, i) => { if (v) for (let j = 0; j < n && i + j < N_MONTH; j++) o[i + j] += v / n; });
    return o;
  };
  // 편성을 바꾸면 9월본 합계를 새 회차 비중대로 다시 나눈다(근사).
  // 방영월에 잡히는 항목은 그 달 회차 비중, 국내유통은 방영월분과 익월 정산분 비율(α)을 9월본에서 맞춰 나눈다
  const byEp = (obj, lag, mul) => {
    const total = Object.values(obj || {}).reduce((a, b) => a + b, 0) * mul;
    const a = new Array(N_MONTH).fill(0);
    segs.forEach((n, k) => { const i = startIdx + k + lag; if (i < N_MONTH) a[i] += total * n / eps; });
    return a;
  };
  const domAlpha = (() => {
    const bs = idxOf(t.start[0], t.start[1]), n = t.segs.length + 1, e0 = t.segs.reduce((a, b) => a + b, 0);
    const o = series(t.base.국내유통매출, 0, 1), tot = o.reduce((a, b) => a + b, 0) || 1;
    let num = 0, den = 0;
    for (let d = 0; d < n; d++) {
      const a = (t.segs[d] || 0) / e0, b = (d > 0 ? t.segs[d - 1] || 0 : 0) / e0, ob = o[bs + d] / tot;
      num += (ob - b) * (a - b); den += (a - b) * (a - b);
    }
    return den ? Math.min(1, Math.max(0, num / den)) : 1;
  })();
  const domMix = (obj, mul) => byEp(obj, 0, mul * domAlpha).map((v, i) => v + byEp(obj, 1, mul * (1 - domAlpha))[i]);
  if (!segChanged) {
    L.국내유통 = series(t.base.국내유통매출, shift, s.domMul);
    L.해외유통 = series(t.base.해외유통매출, shift, ovsMul);
    L.협찬매출 = series(t.base.협찬매출, shift, s.spMul);
    L.협찬원가 = series(t.base.협찬원가, shift, s.spMul);
    L.저작권료국내 = series(t.base['저작권료(국내)'], shift, s.domMul);
    L.저작권료해외 = series(t.base['저작권료(해외)'], shift, ovsMul);
    L.재제작비 = series(t.base.재제작비, shift, 1);
    L.기타원가 = series(t.base.기타원가, shift, 1);
    L.외주RS = series(t.base.외주RS, shift, 1); // 정산월에 한꺼번에 잡혀 같은 개월 수만큼 옮긴다
  } else {
    L.국내유통 = domMix(t.base.국내유통매출, s.domMul);
    L.저작권료국내 = domMix(t.base['저작권료(국내)'], s.domMul);
    L.해외유통 = byEp(t.base.해외유통매출, 0, ovsMul);
    L.저작권료해외 = byEp(t.base['저작권료(해외)'], 0, ovsMul);
    L.협찬매출 = byEp(t.base.협찬매출, 0, s.spMul);
    L.협찬원가 = byEp(t.base.협찬원가, 0, s.spMul);
    // 재제작비는 회당 금액이라(9월본 신작은 회차 × 0.03억) 회차를 바꾸면 같은 비율로 바꾼다
    const rt = Object.values(t.base.재제작비 || {}).reduce((a, b) => a + b, 0) * eps / t.segs.reduce((a, b) => a + b, 0);
    L.재제작비 = new Array(N_MONTH).fill(0);
    const ei = idxOf(end[0], end[1]); if (ei < N_MONTH) L.재제작비[ei] = rt;
    L.기타원가 = series(t.base.기타원가, shift, 1);
    L.외주RS = series(t.base.외주RS, shift, 1); // 정산월에 한꺼번에 잡혀 같은 개월 수만큼 옮긴다
  }
  L.국내유통 = spread(L.국내유통, s.domSpread);
  L.해외유통 = spread(L.해외유통, s.ovsSpread);
  L.저작권료국내 = spread(L.저작권료국내, s.domSpread);
  L.저작권료해외 = spread(L.저작권료해외, s.ovsSpread);

  // 판권 = (회당 표제비 − 회당 방영권료) × 회차, 방영월별 구간으로 나눔
  const perEp = s.feePerEp - s.rightsPerEp;
  const segA = segs.map(n => perEp * n);
  const totalA = segA.reduce((a, b) => a + b, 0);
  const saveIdx = t.saving ? idxOf(...t.saving.month) + shift : Infinity;
  const reg = segs.map(() => new Array(N_MONTH).fill(0));
  const imp = segs.map(() => new Array(N_MONTH).fill(0));
  const tests = [];
  let impCum = 0, adjN = 0;
  const TM = testMonths(g.testUntil || 2027);

  for (let i = 0; i < N_MONTH; i++) {
    segs.forEach((n, k) => {
      const A = segA[k] - (t.saving && i >= saveIdx ? t.saving.amount * segs[k] / eps : 0);
      const mc = i - (startIdx + k) + 1; // 월차
      if (mc <= 0) return;
      let past = 0;
      for (let j = 0; j < i; j++) past += reg[k][j] + imp[k][j];
      const rem = A - past;
      let v;
      if (rem <= 0) v = rem;
      else if (type === '독점') v = mc > g.exMonths ? rem : rem / (g.exMonths + 1 - mc);
      else if (mc > g.jhTotal) v = rem;
      else if (mc <= g.jhInit) v = A * g.jhRatio * (mc / g.jhInit) - past;
      else v = A * g.jhRatio + A * (1 - g.jhRatio) * (mc - g.jhInit) / (g.jhTotal - g.jhInit) - past;
      reg[k][i] = v;
    });

    if (TM.includes(i)) {
      const D = monthEnd(i), { y, m } = ymOf(i);
      const G = cmpDate(start, D) > 0 ? 0 : datedifM(start, D) + (start[2] < 15 ? 1 : 0); // 방영 전이면 원본 IFERROR처럼 0
      const ended = cmpDate(end, D) <= 0;
      const target = type === '조합' || m === 6 || m === 12; // 손상 시험: 독점은 반기(6·12월), 조합은 분기
      const over = type === '독점' ? G > 6 : G > 18;
      const base = type === '조합' ? totalA * g.jhTestBase : totalA;
      let N = 0; segs.forEach((_, k) => { for (let j = 0; j <= i; j++) N += reg[k][j]; });
      // 장부가 = 기준가 − 누적 손상 − 누적 상각 − 반영된 제작비 절감(9월본 가속상각 P열 '감소금액')
      const book = r8(base - impCum - N - (t.saving && i >= saveIdx ? t.saving.amount : 0));
      let P = 0, Cr = 0;
      for (let j = idxOf(2026, 1); j <= i; j++) { P += L.국내유통[j] + L.해외유통[j]; Cr += L.저작권료국내[j]; }
      let Q = 0;
      if (over) Q = 1;
      else { const tb = type === '독점' ? EXP_EX : EXP_JH; Q = (G >= 1 && G <= tb.length) ? tb[G - 1] : 0; }
      const R = Q ? P / Q : 0;
      const T = P ? Cr / P : 0;
      const U = (R - P) * (1 - T);
      const calc = (ended && target) ? Math.max(r8(book - U), 0) : 0;
      let V = calc;
      if (g.impMode === 'fixed') V = (t.fixedImp && t.fixedImp[keyOf(i)]) || 0;
      // adj: 9월본 조건의 (확정값 − 수식값)을 시험 순서대로 더한다. 조건이 그대로면 확정값과 같고, 바꾸면 바뀐 만큼만 움직인다
      // 손상은 아직 비용이 안 된 판권(판권 금액 − 누적 손상 − 누적 상각 − 반영된 절감)을 넘을 수 없다
      else if (g.impMode === 'adj' && ended && target) V = Math.min(Math.max(calc + ((g.impAdj || [])[adjN] || 0), 0), Math.max(r8(totalA - impCum - N - (t.saving && i >= saveIdx ? t.saving.amount : 0)), 0));
      if (ended && target) adjN++;
      tests.push({ key: keyOf(i), G, ended, target, over, base, N, prevImp: impCum, book, P, Q, expected: U, calc, applied: V });
      impCum += V;
      segs.forEach((n, k) => { imp[k][i] = V * n / eps; });
    }
  }
  L.정기상각 = new Array(N_MONTH).fill(0);
  L.손상 = new Array(N_MONTH).fill(0);
  for (let i = 0; i < N_MONTH; i++) segs.forEach((_, k) => { L.정기상각[i] += reg[k][i]; L.손상[i] += imp[k][i]; });

  // 외주 R/S(원본 프로젝트손익 BL·BQ): 정산대상손익 = 방영권료 + 유통매출 × 인정률(보통 80%) − 총제작비 − 기타제작원가.
  // 이익일 때만 × R/S율(정부지원 Y 작품은 +20%), 종영 3개월 뒤 달에 한 번에
  if (t.rs) {
    const tot = a => a.reduce((x, y) => x + y, 0);
    const settle = tot(L.방영권료) + (tot(L.국내유통) + tot(L.해외유통)) * t.rs.recog - s.feePerEp * eps - tot(L.기타원가);
    L.외주RS = new Array(N_MONTH).fill(0);
    const ri = idxOf(end[0], end[1]) + 3;
    if (settle > 0 && ri < N_MONTH) L.외주RS[ri] = settle * (t.rs.rate + (t.rs.plus20 ? 0.2 : 0));
  }
  L.매출 = L.방영권료.map((_, i) => L.방영권료[i] + L.국내유통[i] + L.해외유통[i] + L.협찬매출[i]);
  L.원가 = L.방영권료.map((_, i) => L.제작매출원가[i] + L.정기상각[i] + L.손상[i] + L.협찬원가[i] + L.저작권료국내[i] + L.저작권료해외[i] + L.재제작비[i] + L.기타원가[i] + L.외주RS[i]);
  L.손익 = L.매출.map((v, i) => v - L.원가[i]);
  return { lines: L, totalA, segA, tests, start, end, startIdx, segs, domAlpha };
}

const PERIODS = { FY26: [idxOf(2026, 1), idxOf(2026, 12)], FY27: [idxOf(2027, 1), idxOf(2027, 12)], 'FY28~': [idxOf(2028, 1), N_MONTH - 1], 전체: [0, N_MONTH - 1] };
function sumP(arr, p) { const [a, b] = PERIODS[p]; let s = 0; for (let i = a; i <= b; i++) s += arr[i]; return s; }

if (typeof module !== 'undefined') module.exports = { runTitle, sumP, PERIODS, idxOf, keyOf, N_MONTH };
