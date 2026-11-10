# 월 마감 갱신: 마감 손익계산서 엑셀 하나로 시뮬레이터 HTML과 대사 리포트를 새로 만든다
#   사용: python refresh.py "마감 손익계산서.xlsx"   (또는 실행.bat 위에 엑셀을 끌어다 놓기)
#   결과: 결과/YYMMDD 프로젝트 손익 시뮬레이터(N월본).html, 결과/YYMMDD N월본 갱신 대사 리포트.md
# 행·열은 위치가 아니라 머리글 이름으로 찾는다. 원본 엑셀은 읽기만 한다.
import sys, os, re, json, datetime, subprocess, shutil, warnings
import openpyxl
warnings.filterwarnings('ignore')
try: sys.stdout.reconfigure(encoding='utf-8')
except Exception: pass

HERE = os.path.dirname(os.path.abspath(__file__))
OUTDIR = os.path.dirname(HERE)  # 결과 폴더
n = lambda x: x / 1e8 if isinstance(x, (int, float)) else 0
isdate = lambda x: isinstance(x, (datetime.datetime, datetime.date))
txt = lambda x: '' if x is None else str(x).strip()


def rows_of(wb, name, max_row=None, max_col=None):
    return list(wb[name].iter_rows(max_row=max_row, max_col=max_col, values_only=True))


def month_cols(rows):
    """6행 날짜·7행 구분으로 (연, 월) → 열 번호. 사업계획 열은 빼고, 같은 날짜면 뒤쪽(월 열)을 쓴다"""
    r6, r7 = rows[5], rows[6]
    out = {}
    for j, d in enumerate(r6):
        if isdate(d) and txt(r7[j] if j < len(r7) else None) != '사업계획':
            out[(d.year, d.month)] = j
    return out


def find_row(rows, pred, start=0):
    for i in range(start, len(rows)):
        if pred(rows[i]): return i
    return -1


def col_of(row, label, nth=1):
    k = 0
    for j, v in enumerate(row):
        if txt(v) == label:
            k += 1
            if k == nth: return j
    return -1


# ---------------- 추출 ----------------
def extract(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    D = {'file': os.path.basename(path)}

    # P&L: 마감월, 회사 손익
    pl = rows_of(wb, 'P&L', 200)
    r5, r6 = pl[4], pl[5]
    c26 = next(j for j, v in enumerate(r5) if re.search(r'20\d\d', txt(v)) and '실적' in txt(v) and txt(r6[j]) == '연간')
    year = int(re.search(r'(20\d\d)', txt(r5[c26])).group(1))
    c27 = next((j for j, v in enumerate(r5) if re.search(r'%d년|%d' % (year % 100 + 1, year + 1), txt(v)) and txt(r6[j]) == '연간'), -1)
    close_m = 0
    for j in range(c26 + 1, len(r6)):
        if txt(r5[j]) == '실적' and str(r6[j]).isdigit(): close_m = max(close_m, int(r6[j]))
    if not close_m:
        m = re.search(r'(\d{1,2})\s*월\s*마감', os.path.basename(path)); close_m = int(m.group(1)) if m else 0
    D['close'] = [year, close_m]
    KEYS = {'rev': (3, '매출'), 'capRev': (4, 'Captive매출'), 'old13': (1, '국내유통매출_Cap구작'), 'old16': (1, '해외유통매출_Cap구작'), 'fmt': (6, '포맷/리메이크'),
            'pv': (1, '현할차'), 'capEtc': (5, '기타매출'), 'ncRev': (4, 'Non-cap.매출'), 'buRev': (4, '부가사업매출'), 'etcRev': (4, '기타매출'),
            'cost': (3, '원가'), 'capCost': (4, 'Captive원가'), 'amort': (1, '판권상각비'), 'labor': (1, '제작인건비소계'), 'roy': (5, '판권로열티'),
            'etcCost59': (5, '기타제작원가'), 'rs60': (1, '외주제작사R/S'), 'ncCost': (4, 'Non-cap.원가'), 'buCost': (4, '부가사업원가'), 'etcCost': (4, '기타원가'),
            'gp': (3, '매출이익'), 'sga': (3, '판관비'), 'sgaLabor': (1, '인건비'), 'dep': (1, '감가상각비'), 'intang': (1, '무형자산상각비'), 'bad': (1, '대손상각비'),
            'op': (3, '영업이익'), 'non': (3, '영업외손익'), 'nIn': (4, '영업외수익'), 'nOut': (4, '영업외비용'), 'pbt': (3, '세전이익'), 'ebitda': (3, 'EBITDA')}
    rowk, miss = {}, []
    for k, (c, lab) in KEYS.items():
        i = find_row(pl, lambda r: len(r) > c and txt(r[c]) == lab, 6)
        if i < 0: miss.append(lab)
        rowk[k] = i
    if miss: raise SystemExit('P&L 시트에서 이 행을 찾지 못했습니다: ' + ', '.join(miss))
    D['pl'] = {'y26': {k: n(pl[i][c26]) for k, i in rowk.items()}, 'y27': {k: n(pl[i][c27]) for k, i in rowk.items()} if c27 >= 0 else None}

    # 계산 시트: 작품 × 계정 × 월
    ACC = {('1.제작', '방영권료'): '방영권료', ('3-1.국내유통', '국내유통매출'): '국내유통', ('3-2.해외유통', '해외유통매출'): '해외유통',
           ('2.협찬', '협찬매출'): '협찬매출', ('1.제작', '제작매출원가'): '제작매출원가', ('1.제작', '기타제작원가'): '기타제작원가', ('1.제작', '외주제작사R/S'): '외주RS',
           ('1.제작', '판권상각비'): '판권상각비', ('2.협찬', '협찬원가'): '협찬원가', ('3-1.국내유통', '판권로열티국내'): '판권로열티', ('3-2.해외유통', '판권로열티해외'): '판권로열티',
           ('3-1.국내유통', '저작권료국내'): '저작권료국내', ('3-2.해외유통', '저작권료해외'): '저작권료해외', ('3-2.해외유통', '재제작비'): '재제작비'}
    base = {}
    cache = {}
    for (sh, lab), acc in ACC.items():
        if sh not in cache: cache[sh] = rows_of(wb, sh)
        rows = cache[sh]; mc = month_cols(rows)
        for r in rows[7:]:
            if len(r) > 6 and r[3] == lab and r[6] not in (None, '', '제외'):
                cur = base.setdefault(str(r[6]).strip(), {}).setdefault(acc, {})
                for (y, m), j in mc.items():
                    v = n(r[j]) if j < len(r) else 0
                    if abs(v) > 1e-12: cur[f'{y}-{m:02d}'] = cur.get(f'{y}-{m:02d}', 0) + v
    # 상각raw: 마지막 월 열 다음 해부터의 판권상각 + 제작비 절감
    ar = rows_of(wb, '상각raw')
    hi = find_row(ar, lambda r: '반영시기' in [txt(x) for x in r])
    cV, cW = col_of(ar[hi], '반영시기'), col_of(ar[hi], '감소금액')
    cT = col_of(ar[hi - 1], '타이틀명') if col_of(ar[hi - 1], '타이틀명') >= 0 else col_of(ar[hi], '타이틀명')
    di = find_row(ar, lambda r: sum(1 for x in r if isdate(x)) >= 12, hi)
    last_calc = max(month_cols(cache['1.제작']).keys())
    cols_late = [(j, d) for j, d in enumerate(ar[di]) if isdate(d) and (d.year, d.month) > tuple(last_calc)]
    saving = {}
    for r in ar[di + 1:]:
        if len(r) <= cT or not r[cT]: continue
        t = str(r[cT]).strip()
        if t in base:
            cur = base[t].setdefault('판권상각비', {})
            for j, d in cols_late:
                v = n(r[j]) if j < len(r) else 0
                if abs(v) > 1e-12: cur[f'{d.year}-{d.month:02d}'] = cur.get(f'{d.year}-{d.month:02d}', 0) + v
        if txt(r[1]) == '정기상각' and len(r) > cW and isdate(r[cV]) and isinstance(r[cW], (int, float)) and r[cW]:
            s = saving.setdefault(t, {'amount': 0, 'month': [r[cV].year, r[cV].month]}); s['amount'] += n(r[cW])
    D['saving'] = saving
    # 가속상각: 확정 손상(두 번째 "당기 손상인식액")
    ga = rows_of(wb, '가속상각')
    hi = find_row(ga, lambda r: '콘텐트명' in [txt(x) for x in r])
    cD, cT, cAE = col_of(ga[hi], '대상일자'), col_of(ga[hi], '콘텐트명'), col_of(ga[hi], '당기 손상인식액', 2)
    fixed = {}
    for r in ga:
        if len(r) > cAE and isdate(r[cD]) and r[cT] and isinstance(r[cAE], (int, float)) and abs(r[cAE]) > 0:
            fixed.setdefault(str(r[cT]).strip(), {})[f'{r[cD].year}-{r[cD].month:02d}'] = n(r[cAE])
    D['fixedImp'] = fixed
    D['base'] = base

    # 라인업: Captive 조건, Non-captive R/S
    lu = rows_of(wb, '라인업', 260, 60)
    hi = find_row(lu, lambda r: 'Slot' in [txt(x) for x in r] and '방영시작일' in [txt(x) for x in r])
    h = [txt(x) for x in lu[hi]]
    C = {k: h.index(v) for k, v in dict(fy='연도', title='타이틀명', eps='방영회차', slot='Slot', s='방영시작일', e='방영종료일', type='독점여부', plat='플랫폼').items()}
    C['rights'] = next(j for j, v in enumerate(h) if v.startswith('방영권료'))
    C['fee'] = next(j for j, v in enumerate(h) if '청구표제비' in v)
    # 리쿱율(제작비 연동 판매): 해외 플랫폼 리쿱율, 국내 TVING·쿠팡 리쿱율. 공동제작사(R/S)는 R/S 20% 가산 조건에 쓴다
    C['ovsRate'] = h.index('리쿱율') if '리쿱율' in h else -1
    C['domRates'] = [j for j, v in enumerate(h) if v in ('TVING리쿱율', '쿠팡리쿱율')]
    C['coRS'] = h.index('공동제작사(R/S)') if '공동제작사(R/S)' in h else -1
    num = lambda x: x if isinstance(x, (int, float)) else 0
    lineup = {}
    for r in lu[hi + 1:]:
        if txt(r[C['title']]) == '타이틀명': break
        t = txt(r[C['title']])
        if not t or t == '제외' or not r[C['fy']] or not isdate(r[C['s']]): continue
        lineup[t] = dict(fy=r[C['fy']], eps=r[C['eps']], slot=r[C['slot']], start=[r[C['s']].year, r[C['s']].month, r[C['s']].day],
                         end=[r[C['e']].year, r[C['e']].month, r[C['e']].day] if isdate(r[C['e']]) else None,
                         type=r[C['type']], platform=r[C['plat']], rightsPerEp=n(r[C['rights']]), feePerEp=n(r[C['fee']]), group=txt(r[1]),
                         ovsRate=num(r[C['ovsRate']]) if C['ovsRate'] >= 0 else 0, domRate=sum(num(r[j]) for j in C['domRates']),
                         coRS=num(r[C['coRS']]) if C['coRS'] >= 0 else 0)
    D['lineup'] = lineup
    hn = find_row(lu, lambda r: '공개회차' in [txt(x) for x in r])
    hN = [txt(x) for x in lu[hn]]
    ncl = {}
    if hn >= 0:
        cTn, cP, cRS = hN.index('타이틀명'), hN.index('공동제작사'), hN.index('R/S율')
        hP = [txt(x) for x in lu[hn - 1]]
        cPl = hN.index('플랫폼') if '플랫폼' in hN else (hP.index('플랫폼') if '플랫폼' in hP else -1)
        for r in lu[hn + 1:]:
            t = txt(r[cTn])
            if t and t != '제외': ncl[t] = dict(partner=r[cP], rsRate=r[cRS] if isinstance(r[cRS], (int, float)) else 0, platform=r[cPl] if cPl >= 0 else '')
    # 4.Non-cap.
    nc = rows_of(wb, '4.Non-cap.')
    mc = month_cols(nc)
    hn = find_row(nc, lambda r: '제작시작일' in [txt(x) for x in r])
    hh = [txt(x) for x in nc[hn]]
    cH, cI, cJ, cK = hh.index('청구표제비'), hh.index('제작수수료'), hh.index('제작시작일'), hh.index('제작종료일')
    NACC = {'제작매출': 'rev', '제작수익': 'fee', '해외유통대행': 'agency', '기타유통매출': 'otherRev', '제작원가': 'cost', '공동제작사 R/S': 'rs', '기타제작원가': 'otherCost', '판권로열티': 'royalty'}
    ncb, ncc = {}, {}
    for r in nc:
        if len(r) > 7 and r[3] in NACC and r[6] not in (None, '', '제외', '구작'):
            t = str(r[6]).strip(); cur = ncb.setdefault(t, {}).setdefault(NACC[r[3]], {})
            for (y, m), j in mc.items():
                v = n(r[j]) if j < len(r) else 0
                if abs(v) > 1e-12: cur[f'{y}-{m:02d}'] = cur.get(f'{y}-{m:02d}', 0) + v
            if r[3] == '제작매출':
                ncc[t] = dict(fy=r[4], total=n(r[cH]), commission=n(r[cI]), start=[r[cJ].year, r[cJ].month] if isdate(r[cJ]) else None,
                              end=[r[cK].year, r[cK].month] if isdate(r[cK]) else None, **ncl.get(t, {}))
    D['ncBase'], D['ncCond'] = ncb, ncc
    # 프로젝트손익: 작품별 FY 매출이익
    pj = rows_of(wb, '프로젝트손익', 300)
    hm = find_row(pj, lambda r: '매출이익' in [txt(x) for x in r])
    cm = col_of(pj[hm], '매출이익')
    fyr = pj[hm + 1]
    fyc = {}
    for j in range(cm, min(len(fyr), cm + 8)):  # 매출이익 머리 바로 뒤의 FY 열(첫 번째 것만)
        mt = re.match(r'(FY\d\d)', txt(fyr[j]))
        if mt and mt.group(1) not in fyc: fyc[mt.group(1)] = j
    fy0 = year % 100 - 1
    lab = [f'FY{fy0}', f'FY{fy0 + 1}', f'FY{fy0 + 2}', f'FY{fy0 + 3}']
    D['pjLabels'] = lab[:3] + [lab[3] + '~']
    pjo = {}
    for r in pj[hm + 2:]:
        t = txt(r[4]) if len(r) > 4 else ''
        if t and isinstance(r[fyc.get(lab[0], -1)] if lab[0] in fyc else None, (int, float)):
            pjo[t] = dict(c=r[2], fy=[n(r[fyc[x]]) if x in fyc else 0 for x in lab])
    D['pj'] = pjo
    # 프로젝트손익: 외주제작사R/S 원본 산식(BQ 정산대상손익의 유통 인정률, BP R/S율, BL의 정부지원 20% 가산)
    hr = find_row(pj, lambda r: '(R/S율)' in [txt(x) for x in r] and '(정산대상손익)' in [txt(x) for x in r])
    rsp = {}
    if hr >= 0:
        cBP, cBQ, cBL = col_of(pj[hr], '(R/S율)'), col_of(pj[hr], '(정산대상손익)'), col_of(pj[hr], '외주제작사R/S')
        wf = openpyxl.load_workbook(path, data_only=False, read_only=True)
        pf = rows_of(wf, '프로젝트손익', len(pj))
        ft = lambda x: getattr(x, 'text', x) if x is not None else ''
        gov = {}  # 상각raw 정부지원금 표(5행 타이틀, 9행 Y/N)
        for j in range(len(ar[4])):
            if j < len(ar[8]) and txt(ar[4][j]): gov[txt(ar[4][j])] = txt(ar[8][j])
        for i in range(hr + 1, len(pj)):
            t = txt(pj[i][4]) if len(pj[i]) > 4 else ''
            if t not in D['lineup']: continue
            fq, fl = str(ft(pf[i][cBQ])), str(ft(pf[i][cBL]))
            mo = re.search(r'\)\*\(1-([\d.]+)%\)', fq) or None
            mo2 = re.search(r'\)\*([\d.]+)%', fq)
            recog = 1 - float(mo.group(1)) / 100 if mo else (float(mo2.group(1)) / 100 if mo2 else 0.8)
            plus20 = '$I$9:$X$9' in fl and gov.get(t) == 'Y' and abs(D['lineup'][t]['coRS']) > 0
            rsp[t] = dict(rate=pj[i][cBP] if isinstance(pj[i][cBP], (int, float)) else 0, recog=recog, plus20=plus20,
                          bl=n(pj[i][cBL]), bq=n(pj[i][cBQ]), formula=fq)
    D['rsParam'] = rsp
    return D


# ---------------- 화면 데이터 ----------------
def ym_key(y, m): return f'{y}-{m:02d}'


def link_share(c, bb):
    # 유통매출 중 리쿱율로 정해지는 몫(= 회당 청구표제비 × 리쿱율 × 회차) ÷ 그 유통매출 합계. 나머지(IPTV·현업 입력 등)는 고정
    S = lambda o: sum(o.values()) if o else 0
    eps = c['eps'] or 0
    out = {}
    for side, rate, key in (('dom', c.get('domRate', 0), '국내유통매출'), ('ovs', c.get('ovsRate', 0), '해외유통매출')):
        tot = S(bb.get(key))
        out[side] = round(min(1, max(0, c['feePerEp'] * rate * eps / tot)), 6) if tot > 1e-9 else 0
    return out


def build(D):
    y, m = D['close']
    close = datetime.date(y, m, 1) + datetime.timedelta(days=32)
    close = close.replace(day=1) - datetime.timedelta(days=1)
    base, pj = D['base'], D['pj']
    # Captive
    KEYMAP = {'국내유통': '국내유통매출', '해외유통': '해외유통매출', '협찬매출': '협찬매출', '협찬원가': '협찬원가', '저작권료국내': '저작권료(국내)', '저작권료해외': '저작권료(해외)', '재제작비': '재제작비', '외주RS': '외주RS'}
    LU = []
    for name, c in D['lineup'].items():
        b = base.get(name, {})
        bb = {KEYMAP[k]: v for k, v in b.items() if k in KEYMAP}
        oth = {}
        for k in ('기타제작원가', '판권로열티'):
            for mm, v in b.get(k, {}).items(): oth[mm] = oth.get(mm, 0) + v
        bb['기타원가'] = oth
        p = pj.get(name, {}).get('fy', [0, 0, 0, 0])
        LU.append(dict(name=name, fy=c['fy'], slot=c['slot'], lane='타채널' if '타채널' in name else c['slot'], start=c['start'], end=c['end'], eps=c['eps'],
                       type=c['type'] if c['type'] in ('독점', '조합') else '독점', typeRaw=c['type'] if c['type'] in ('독점', '조합') else '구분 없음', platform=c['platform'],
                       fee=c['feePerEp'], rights=c['rightsPerEp'], locked=datetime.date(*c['start']) <= close, base=bb,
                       raw={k: b.get(k, {}) for k in ('방영권료', '제작매출원가', '판권상각비')},
                       fixedImp=D['fixedImp'].get(name, {}), saving=D['saving'].get(name), pj=p, group=c['group'],
                       link=link_share(c, bb), rs=D['rsParam'].get(name)))
    cap = set(D['lineup'])
    NONCAP = [dict(name=t, fy=v['c'], pj=v['fy']) for t, v in pj.items() if t not in cap and t != '제외']
    # Non-captive
    NC = []
    for t, a in D['ncBase'].items():
        c = D['ncCond'].get(t, {})
        if not c.get('start') or not c.get('end'): continue
        other = {}
        for k, sg in (('agency', 1), ('otherRev', 1), ('otherCost', -1), ('royalty', -1)):
            for mm, v in a.get(k, {}).items(): other[mm] = other.get(mm, 0) + sg * v
        s, e = c['start'], c['end']
        locked = tuple(s) <= (y, m)
        NC.append(dict(name=t, fy=c['fy'], start=s, dur=(e[0] - s[0]) * 12 + e[1] - s[1] + 1, total=c.get('total', 0),
                       fee=c.get('commission', 0) if not locked else sum(a.get('fee', {}).values()), rsRate=c.get('rsRate', 0) or 0,
                       platform=c.get('platform') or '', partner=c.get('partner') or '', locked=locked,
                       base={k: a.get(k, {}) for k in ('rev', 'fee', 'cost', 'rs')} | {'other': other}))
    names = {x['name'] for x in NC}
    years = [y - 1, y, y + 1]
    for t, v in pj.items():
        if t in names or t in cap or t == '제외' or v['c'] in (None,): continue
        if not str(v['c']).startswith('FY'): continue
        o = {ym_key(yy, 12): v['fy'][i] for i, yy in enumerate(years) if abs(v['fy'][i]) > 1e-9}
        if o: NC.append(dict(name=t, fy=v['c'], start=[y - 1, 1], dur=12, total=0, fee=0, rsRate=0, platform='', partner='', locked=True, yearly=True,
                             base=dict(rev={}, fee={}, cost={}, rs={}, other=o)))
    NC.sort(key=lambda x: (x['start'], x['name']))
    # 회사 P&L 다리
    T = list(D['lineup'])
    acc = lambda a, yy: sum(v for t in T for k, v in base.get(t, {}).get(a, {}).items() if int(k[:4]) == yy)
    REV = ['방영권료', '국내유통', '해외유통', '협찬매출']
    COST = ['제작매출원가', '판권상각비', '협찬원가', '외주RS', '기타제작원가', '판권로열티', '저작권료국내', '저작권료해외', '재제작비']
    PLB = {}
    for key, yy in (('y26', y), ('y27', y + 1)):
        P = D['pl'][key]
        if P is None: continue
        proj = sum(acc(a, yy) for a in REV) - sum(acc(a, yy) for a in COST)
        capm = P['capRev'] - P['capCost']
        it = [['구작 유통매출', P['old13'] + P['old16']], ['구작 판권상각', -(P['amort'] - acc('판권상각비', yy))], ['제작인건비', -P['labor']],
              ['판권로열티', -(P['roy'] - acc('판권로열티', yy))], ['포맷·리메이크, 기타매출', P['fmt'] + P['capEtc']], ['현할차(매출 차감)', P['pv']],
              ['건중자 비용화 등', -(P['etcCost59'] - P['rs60'])]]
        it.append(['구작 저작권료 등 집계 차이', capm - proj - sum(v for _, v in it)])
        dep = P['dep'] + P['intang']
        PLB[key] = dict(proj=proj, capm=capm, cap2=it, nc=P['ncRev'] - P['ncCost'], bu=P['buRev'] - P['buCost'], etc=P['etcRev'] - P['etcCost'], gp=P['gp'],
                        sga=[['인건비', P['sgaLabor']], ['감가상각·무형자산상각', dep], ['대손상각', P['bad']], ['그 밖의 판관비', P['sga'] - P['sgaLabor'] - dep - P['bad']]],
                        sgaT=P['sga'], op=P['op'], nonIn=P['nIn'], nonOut=P['nOut'], non=P['non'], pbt=P['pbt'], ebitda=P['ebitda'], projAmort=acc('판권상각비', yy),
                        projRev=sum(acc(a, yy) for a in REV), projCost=sum(acc(a, yy) for a in COST), capRev=P['capRev'], capCost=P['capCost'], ncRev=P['ncRev'], ncCost=P['ncCost'],
                        buRev=P['buRev'], buCost=P['buCost'], etcRev=P['etcRev'], etcCost=P['etcCost'], rev=P['rev'], cost=P['cost'])
    # 받는 숫자: 확인이 필요한 항목 자동 표시
    flags = {}
    a, b = PLB['y26'], PLB.get('y27')
    if b:
        o26, o27 = a['cap2'][0][1], b['cap2'][0][1]
        if o26 and abs(o27 - o26) / abs(o26) > 0.3: flags['c_old'] = dict(status='확인 중', memo=f'FY26 {o26:,.1f} → FY27 {o27:,.1f}로 크게 바뀌는 계획. 근거 확인')
        bd26, bd27 = a['sga'][2][1], b['sga'][2][1]
        if bd27 < 0 or (bd26 and abs(bd27 - bd26) / abs(bd26) > 0.5): flags['s_bad'] = dict(status='확인 중', memo=f'FY26 {bd26:,.1f} → FY27 {bd27:,.1f}. 환입·변동이 계획대로 맞는지 확인')
        if abs(a['nonIn'] - b['nonIn']) < 0.005: flags['n_in'] = dict(status='확인 중', memo='FY27이 FY26과 같은 값(복사로 보임)')
        if abs(a['nonOut'] - b['nonOut']) < 0.005: flags['n_out'] = dict(status='확인 중', memo='FY27이 FY26과 같은 값(복사로 보임). 차입·금리 계획 확인')
    flags['c_gap'] = dict(status='', memo=f'{m}월본 Captive 매출이익에서 위 항목과 신작 {len(LU)}편을 빼고 남는 값')
    return dict(LU=LU, NONCAP=NONCAP, NC=NC, PLB=PLB, FLAGS=flags, close=[y, m])


# ---------------- 대사 ----------------
def reconcile(D, B):
    out = []
    REVC = ['방영권료', '국내유통', '해외유통', '협찬매출']
    COSTC = ['제작매출원가', '판권상각비', '협찬원가', '외주RS', '기타제작원가', '판권로열티', '저작권료국내', '저작권료해외', '재제작비']
    y = D['close'][0]
    yrs = [(y - 1, y - 1), (y, y), (y + 1, y + 1), (y + 2, 2100)]
    cap_bad = []
    for t in D['lineup']:
        b = D['base'].get(t, {})
        fy = [sum(v for k in REVC for mm, v in b.get(k, {}).items() if a <= int(mm[:4]) <= z) - sum(v for k in COSTC for mm, v in b.get(k, {}).items() if a <= int(mm[:4]) <= z) for a, z in yrs]
        p = D['pj'].get(t)
        if not p: cap_bad.append((t, '프로젝트손익에 없음', fy)); continue
        d = [round(fy[i] - p['fy'][i], 3) for i in range(4)]
        if any(abs(x) > 0.01 for x in d): cap_bad.append((t, d, fy))
    nc_bad = []
    NREV, NCOST = ['rev', 'fee', 'agency', 'otherRev'], ['cost', 'rs', 'otherCost', 'royalty']
    for t, a in D['ncBase'].items():
        fy = [sum(v for k in NREV for mm, v in a.get(k, {}).items() if int(mm[:4]) == yy) - sum(v for k in NCOST for mm, v in a.get(k, {}).items() if int(mm[:4]) == yy) for yy in (y - 1, y, y + 1)]
        p = D['pj'].get(t)
        if not p: continue
        d = [round(fy[i] - p['fy'][i], 3) for i in range(3)]
        if any(abs(x) > 0.01 for x in d): nc_bad.append((t, d))
    pl = []
    for key in ('y26', 'y27'):
        P = B['PLB'].get(key)
        if not P: continue
        chk = P['proj'] + sum(v for _, v in P['cap2']) + P['nc'] + P['bu'] + P['etc']
        pl.append((key, chk - P['gp'], (chk - P['sgaT']) - P['op'], (chk - P['sgaT'] + P['nonIn'] - P['nonOut']) - P['pbt']))
    return dict(cap_bad=cap_bad, nc_bad=nc_bad, pl=pl)


def run_node(B, tmp):
    node = shutil.which('node')
    if not node: return None
    p = os.path.join(tmp, 'lu.json')
    json.dump(B['LU'], open(p, 'w', encoding='utf-8'), ensure_ascii=False)
    try:
        r = subprocess.run([node, os.path.join(HERE, 'verify.js'), p], capture_output=True, text=True, encoding='utf-8', timeout=300)
        return json.loads(r.stdout) if r.returncode == 0 else {'error': r.stderr[-500:]}
    except Exception as e:
        return {'error': str(e)}


# ---------------- HTML ----------------
def write_html(D, B, R, V):
    y, m = B['close']
    tpl = open(os.path.join(HERE, 'template.html'), encoding='utf-8').read()
    a = tpl.index('<section id="lineup"'); z = tpl.index('</section>', a)
    sec = tpl[a:z]
    d0 = tpl.index('/*DATA_START*/')
    code = tpl[d0:]
    def relabel(s):
        for x, yv in (('2026년 9월 마감', f'{y}년 {m}월 마감'), ('9월본', f'{m}월본'), ('9월 마감', f'{m}월 마감'),
                      ('29편', f'{len(B["LU"])}편'), ('37편', f'{len(B["NC"])}편'),
                      ("'lineup9-v1'", f"'lineup-{y}{m:02d}-v1'"), ("'nc9-v1'", f"'nc-{y}{m:02d}-v1'"), ("'recv9-v1'", f"'recv-{y}{m:02d}-v1'")):
            s = s.replace(x, yv)
        return s
    notes = notes_html(D, B, R, V)
    sec = relabel(sec)
    sec = re.sub(r'(<div class="lu-pane" data-tab="note">.*?<ul class="changes small"[^>]*>).*?(</ul>)', lambda mo: mo.group(1) + notes + mo.group(2), sec, flags=re.S)
    code = relabel(code)
    nxt = (y + (m // 12), m % 12 + 1)
    sc = (y if m > 3 else y - 1, (m - 4) % 12 + 1)
    const = (f'/*CLOSE_CONST*/const CLOSE_YM = [{y}, {m}], CLOSE_NEXT = [{nxt[0]}, {nxt[1]}, 1], CLOSE_SCROLL = [{sc[0]}, {sc[1]}, 1], CLOSE_NIDX = idxOf({nxt[0]}, {nxt[1]});'
             f' const RECV_FLAGS = {json.dumps(B["FLAGS"], ensure_ascii=False)};/*CLOSE_CONST_END*/')
    html = tpl[:a] + sec + tpl[z:d0] + code
    html = re.sub(r'/\*CLOSE_CONST\*/.*?/\*CLOSE_CONST_END\*/', lambda _: const, html, flags=re.S)
    js = lambda nm, v: f'const {nm} = ' + json.dumps(v, ensure_ascii=False) + ';'
    html = re.sub(r'const PLB = \{.*?\};\n', lambda _: js('PLB', B['PLB']) + '\n', html, count=1, flags=re.S)
    html = re.sub(r'const LU9 = \[.*?\];\nconst NONCAP9 = \[.*?\];\n', lambda _: js('LU9', B['LU']) + '\n' + js('NONCAP9', B['NONCAP']) + '\n', html, count=1, flags=re.S)
    html = re.sub(r'const NC9 = \[.*?\];\n', lambda _: js('NC9', B['NC']) + '\n', html, count=1, flags=re.S)
    today = datetime.date.today().strftime('%y%m%d')
    base = os.path.join(OUTDIR, f'{today} 프로젝트 손익 시뮬레이터({m}월본)')
    out, k = base + '.html', 0
    while os.path.exists(out): k += 1; out = f'{base}_R{k}.html'
    open(out, 'w', encoding='utf-8').write(html)
    return out


def notes_html(D, B, R, V):
    y, m = B['close']
    mov = sum(1 for x in B['LU'] if not x['locked']); ncm = sum(1 for x in B['NC'] if not x['locked'])
    ipn = sum(1 for x in B['NC'] if x['locked'] and not x.get('yearly') and (x['start'][0] - y) * 12 + x['start'][1] - m + x['dur'] - 1 >= 1 and (x['total'] > 0 or x['fee'] > 0))
    li = [f'출처: {D["file"]} ({y}년 {m}월 마감). 갱신 도구가 P&L·라인업·계산 시트·상각raw·가속상각·4.Non-cap.·프로젝트손익 시트를 머리글 이름으로 찾아 뽑았다.',
          f'Captive 신작 {len(B["LU"])}편 중 {m}월 마감 뒤 방영을 시작하는 {mov}편만 옮기거나 조건을 바꿀 수 있고, 나머지는 실적이 있어 고정이다. Non-captive {len(B["NC"])}편 중 {ncm}편은 제작 일정·조건을 바꿀 수 있고, 실적이 있는데 마감 뒤에도 제작이 이어지는 {ipn}편은 시작월을 뺀 기간·금액·R/S율을 바꿀 수 있다(실적 뒤 남은 금액을 남은 개월로 다시 나눔).',
          '확정 손상(가속상각 시트)과 제작비 절감(상각raw 시트 반영시기·감소금액)은 시트 값을 자동으로 읽었다.',
          f'대사: Captive 작품별 연도 손익이 프로젝트손익과 다른 작품 {len(R["cap_bad"])}편, Non-captive {len(R["nc_bad"])}편. 회사 P&L 다리 차이 ' + ', '.join(f'{k} 매출이익 {g:+.4f}·영업이익 {o:+.4f}·세전 {p:+.4f}' for k, g, o, p in R['pl']) + '. 자세한 내용은 같은 날짜의 갱신 대사 리포트를 본다.',
          '다음 해(FY27) 회사 손익은 Captive 신작·Non-captive를 편성·제작 일정으로 계산하고, 판관비·영업외 등 받는 숫자는 P&L 시트의 다음 해 열을 그대로 쓴다.',
          '「작품 하나로 구조 설명」과 「검증용 상세」 화면은 9월본 기준 설명 자료로, 갱신하지 않는다.',
          '움직이는 숫자(Captive 신작, Non-captive)의 계산 규칙과 받는 숫자 입력판 사용법은 이전 판과 같다. 엑셀을 올리면 받는 숫자를 다시 채울 수 있다.']
    if V and 'maxAmort' in V: li.insert(3, f'옮길 수 있는 Captive {mov}편을 마감본 위치에서 다시 계산하면 월별 판권상각 최대 차이 {V["maxAmort"]:.4f}억, 방영권료 최대 차이 {V["maxRights"]:.4f}억이다.')
    return '\n' + '\n'.join(f'        <li>{x}</li>' for x in li) + '\n      '


def write_report(D, B, R, V, html_path):
    y, m = B['close']
    today = datetime.date.today().strftime('%y%m%d')
    p = os.path.join(OUTDIR, f'{today} {m}월본 갱신 대사 리포트.md')
    k = 0
    while os.path.exists(p): k += 1; p = os.path.join(OUTDIR, f'{today} {m}월본 갱신 대사 리포트_R{k}.md')
    L = [f'# {m}월본 갱신 대사 리포트', '',
         f'{D["file"]}({y}년 {m}월 마감)으로 시뮬레이터를 새로 만들었다. 결과 파일은 `{os.path.basename(html_path)}`이다. 이 리포트는 뽑은 숫자가 원본과 맞는지 확인한 결과다.', '',
         '## 1. 회사 P&L 다리', '',
         '프로젝트 손익에서 세전이익까지 다시 쌓은 값과 P&L 시트 값의 차이다. 0이면 맞다.', '',
         '| 연도 | 매출이익 차이 | 영업이익 차이 | 세전이익 차이 |', '|---|---|---|---|']
    L += [f'| {"FY" + str(y % 100) if k == "y26" else "FY" + str(y % 100 + 1)} | {g:+.4f} | {o:+.4f} | {pp:+.4f} |' for k, g, o, pp in R['pl']]
    L += ['', '## 2. Captive 신작 작품별 연도 손익', '',
          f'계산 시트 월별 값을 더한 연도 손익과 프로젝트손익 시트 값을 비교했다. 차이가 0.01억을 넘는 작품: {len(R["cap_bad"])}편.', '']
    if R['cap_bad']:
        L += ['| 작품 | 차이(' + ' / '.join(D['pjLabels']) + ') |', '|---|---|']
        L += [f'| {t} | {d if isinstance(d, str) else " / ".join(f"{x:+.2f}" for x in d)} |' for t, d, _ in R['cap_bad']]
        L += ['', '9월본에서는 옥씨부인전 FY28~(원본 프로젝트손익에 상각 누락), 27미정1 FY28~(원본 칸이 상각만 담음), 신의구슬 FY27 국내유통 집계 차이가 이런 원인이었다. 화면은 월별 값 합계를 쓴다.']
    L += ['', '## 3. 옮길 수 있는 작품 재계산', '']
    if V is None: L += ['node가 없어 이 검증은 건너뛰었다. 다른 PC에서 돌렸다면 node 설치 뒤 다시 실행한다.']
    elif 'error' in V: L += ['검증 실행 오류: ' + V['error']]
    else:
        L += [f'마감본 위치·조건 그대로 계산 엔진으로 다시 계산한 결과다. 월별 판권상각 최대 차이 {V["maxAmort"]:.4f}억, 방영권료 최대 차이 {V["maxRights"]:.4f}억.', '',
              '| 작품 | 월별 회차 | 판권상각 최대 차이 | 방영권료 최대 차이 |', '|---|---|---|---|']
        L += [f'| {x["name"]} | {x["segs"]} | {x["dA"]:.4f} | {x["dR"]:.4f} |' for x in V['rows']]
    if V and 'rs' in V:
        L += ['', '### 외주제작사R/S 원본 산식 대조', '', '정산대상손익 = 방영권료 + 유통매출 × 인정률 − 총제작비 − 기타제작원가, 이익일 때만 × R/S율(정부지원 Y면 +20%), 종영 3개월 뒤.', '',
              '| 작품 | R/S율 | 인정률 | 원본 정산대상손익 | 계산 정산대상손익 | 원본 R/S | 계산 R/S | 인식월 | 9월본 월별 반영 |', '|---|---|---|---|---|---|---|---|---|']
        L += [f'| {x["name"]} | {x["rate"]:.0%} | {x["recog"]:.0%} | {x["bq"]:.2f} | {x["settle"]:.2f} | {x["bl"]:.2f} | {x["rs"]:.2f} | {x["month"]} | {x["mon"]:.2f} |' for x in V['rs']]
        L += ['', '원본 정산대상손익은 FY25~27 합계만 더해, 2028년에 들어오는 유통매출이 있으면 계산값과 조금 다르다. 9월본 월별 반영이 0인데 원본 R/S가 있으면 인식월이 계산 시트 기간(2027년) 밖이라 9월본 손익에 빠진 것이다.']
    L += ['', '### 유통매출 제작비 연동 몫', '', '리쿱율로 정해지는 몫(회당 청구표제비 × 리쿱율 × 회차) ÷ 유통매출 합계. 회당 제작비를 바꾸면 이 몫만 같이 바뀐다.', '']
    L += [f'- {x["name"]}: 국내 {x["link"]["dom"]:.0%}, 해외 {x["link"]["ovs"]:.0%}' for x in B['LU'] if not x['locked'] and x.get('link')]
    L += ['', '## 4. Non-captive', '', f'작품별 연도 손익이 프로젝트손익과 다른 작품: {len(R["nc_bad"])}편.']
    L += [f'- {t}: {" / ".join(f"{x:+.2f}" for x in d)}' for t, d in R['nc_bad']]
    L += ['', '## 5. 자동으로 읽은 값', '',
          '- 확정 손상: ' + (', '.join(f'{t} {k} {v:.2f}억' for t, kv in D['fixedImp'].items() for k, v in kv.items() if t in D['lineup'] and not next(x for x in B['LU'] if x['name'] == t)['locked']) or '옮길 수 있는 작품에는 없음'),
          '- 제작비 절감: ' + (', '.join(f'{t} {v["amount"]:.2f}억({v["month"][0]}-{v["month"][1]:02d})' for t, v in D['saving'].items()) or '없음'),
          '- 받는 숫자 중 확인 필요로 표시한 항목: ' + (', '.join(f'{k}({v["memo"]})' for k, v in B['FLAGS'].items() if v.get('status')) or '없음'), '',
          '## 6. 확인할 것', '', '- 위 1~4에서 차이가 있으면 원인을 확인한 뒤 시뮬레이터를 쓴다.', '- 새로 생긴 작품이나 빠진 작품이 맞는지 편성표에서 확인한다.',
          '- 이전 판에서 브라우저에 저장한 배치는 마감월이 달라 새 판에 따라오지 않는다(새로 시작).']
    open(p, 'w', encoding='utf-8').write('\n'.join(L) + '\n')
    return p


def main():
    if len(sys.argv) < 2: raise SystemExit('사용: python refresh.py "마감 손익계산서.xlsx"')
    path = sys.argv[1]
    print('읽는 중:', os.path.basename(path), '(1~2분 걸릴 수 있음)')
    D = extract(path)
    print(f'마감월 {D["close"][0]}년 {D["close"][1]}월 · Captive {len(D["lineup"])}편 · Non-captive {len(D["ncCond"])}편')
    B = build(D)
    R = reconcile(D, B)
    import tempfile
    V = run_node(B, tempfile.gettempdir())
    html = write_html(D, B, R, V)
    rep = write_report(D, B, R, V, html)
    print('만든 파일:'); print(' -', html); print(' -', rep)
    print('대사: Captive 차이', len(R['cap_bad']), '편 / Non-captive 차이', len(R['nc_bad']), '편 / P&L 다리', [(k, round(g, 4), round(o, 4), round(p, 4)) for k, g, o, p in R['pl']])
    if V and 'maxAmort' in V: print(f'재계산: 판권상각 최대 차이 {V["maxAmort"]:.4f}, 방영권료 {V["maxRights"]:.4f}, 손상 방식(확정값+조건 변경분) 최대 차이 {V["maxAdj"]:.4f}')
    if V and 'rs' in V: print('R/S:', [(x['name'], round(x['bl'], 3), round(x['rs'], 3), x['month']) for x in V['rs'] if x['bl'] or x['rs']])


if __name__ == '__main__':
    main()
