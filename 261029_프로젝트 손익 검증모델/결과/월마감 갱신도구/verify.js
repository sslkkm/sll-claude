// 옮길 수 있는 Captive 작품을 마감본 위치·조건 그대로 계산 엔진으로 다시 계산해, 월별 판권상각·방영권료가 원본과 같은지 본다
// 사용: node verify.js lu.json  → 결과를 JSON으로 출력
const fs = require('fs');
const E = require('./engine.js');
const LU = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const addDays = (d, n) => { const t = new Date(Date.UTC(d[0], d[1] - 1, d[2] + n)); return [t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate()]; };
function eps(p) { // 토일은 토·일 2회, 금금은 금요일 2회
  const o = []; let left = p.eps, w = 0;
  while (left > 0) { const d = addDays(p.start, 7 * w); o.push(d); left--; if (left > 0) { o.push(p.slot === '토일' ? addDays(d, 1) : d); left--; } w++; }
  return o;
}
const G = { exMonths: 6, jhTotal: 60, jhInit: 6, jhRatio: 0.6, jhTestBase: 0.6, impMode: 'fixed', testUntil: 2029 };
const rows = [];
let maxA = 0, maxR = 0, maxAdj = 0; const rsRows = [];
for (const p of LU.filter(x => !x.locked)) {
  const e = eps(p), f = e[0], l = e[e.length - 1];
  const i0 = E.idxOf(f[0], f[1]), i1 = E.idxOf(l[0], l[1]);
  const segs = new Array(i1 - i0 + 1).fill(0); e.forEach(d => segs[E.idxOf(d[0], d[1]) - i0]++);
  const t = { name: p.name, type: p.type, start: p.start, end: p.end, segs, feePerEp: p.fee, rightsPerEp: p.rights, usdShare: 0, baseFx: 1, base: p.base, fixedImp: p.fixedImp, saving: p.saving, rs: p.rs };
  const s = { type: p.type, delay: 0, feePerEp: p.fee, rightsPerEp: p.rights, domMul: 1, ovsMul: 1, spMul: 1, fx: 1, domSpread: 1, ovsSpread: 1, segs };
  const RF = E.runTitle(t, s, G), L = RF.lines;
  // 손상 방식 '확정값 + 조건 변경분'(adj)이 조건 그대로일 때 확정값과 같은지
  const adj = RF.tests.filter(x => x.ended && x.target).map(x => x.applied - x.calc);
  const LA = E.runTitle(t, s, { ...G, impMode: 'adj', impAdj: adj }).lines;
  let dI = 0; for (let i = 0; i < E.N_MONTH; i++) dI = Math.max(dI, Math.abs(LA.손상[i] - L.손상[i]), Math.abs(LA.정기상각[i] - L.정기상각[i]));
  maxAdj = Math.max(maxAdj, dI);
  if (p.rs) { const ri = L.외주RS.findIndex(v => v !== 0); const tot = a => a.reduce((x, y) => x + y, 0);
    const settle = tot(L.방영권료) + (tot(L.국내유통) + tot(L.해외유통)) * p.rs.recog - p.fee * p.eps - tot(L.기타원가);
    rsRows.push({ name: p.name, rate: p.rs.rate, recog: p.rs.recog, bq: p.rs.bq, bl: p.rs.bl, settle, rs: tot(L.외주RS), month: ri >= 0 ? E.keyOf(ri) : '-', mon: Object.values(p.base.외주RS || {}).reduce((x, y) => x + y, 0) }); }
  let dA = 0, dR = 0;
  for (let i = 0; i < E.N_MONTH; i++) { const k = E.keyOf(i); dA = Math.max(dA, Math.abs(L.정기상각[i] + L.손상[i] - (p.raw.판권상각비[k] || 0))); dR = Math.max(dR, Math.abs(L.방영권료[i] - (p.raw.방영권료[k] || 0))); }
  maxA = Math.max(maxA, dA); maxR = Math.max(maxR, dR);
  rows.push({ name: p.name, segs: segs.join('·'), dA, dR });
}
process.stdout.write(JSON.stringify({ maxAmort: maxA, maxRights: maxR, maxAdj, rows, rs: rsRows }));
