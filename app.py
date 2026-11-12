"""Consumo de insumos SC-120 + % de láminas sobre hemogramas.
Uso:  pip install streamlit pandas numpy openpyxl  ->  streamlit run app.py
Carga los CSV Review_/Log_ de cada SC-120 y los Review_ de los analizadores.
Pestaña 'Láminas por alarma': cruza láminas del SC-120 con las alarmas de los analizadores (requiere openpyxl para el Excel de revisión)."""
import csv
import re
from datetime import datetime
import numpy as np
import pandas as pd
import streamlit as st

R = ['Metanol', 'Color', 'Tampón', 'Agua desioniz', 'Diluy.', 'Porta']
NM = {'Metanol': 'Metanol', 'Color': 'Coloración', 'Tampón': 'Tampón', 'Agua desioniz': 'Agua desioniz.',
      'Diluy.': 'Diluyente', 'Porta': 'Porta'}
MODES = ['SM+ST', 'SM', 'ST', 'PR']
DEF = {'SM+ST': set(R), 'SM': {'Metanol', 'Diluy.', 'Porta'},
       'ST': {'Metanol', 'Color', 'Tampón', 'Agua desioniz'}, 'PR': set()}
EN = {'Slide': 'Porta', 'Stain': 'Color', 'Buffer': 'Tampón', 'Methanol': 'Metanol',
      'Deionized': 'Agua desioniz', 'Diluent': 'Diluy.'}
DATE = re.compile(r'(\d{1,2})[-/](\d{1,2})[-/](\d{4})(?:\s+(\d{1,2}):(\d{2}))?')
VOL = re.compile(r'Volum\w*[\s:"]*([\d.,]+)\s*(?:->|→|=>|›|>)\s*([\d.,]+)')
DAYS = lambda s: s.strftime('%Y-%m-%d')


# ---------------------------------------------------------------- lectura
def decode(b):
    return b.decode('utf-16') if b[:2] in (b'\xff\xfe', b'\xfe\xff') else b.decode('utf-8-sig', errors='replace')


def lines(t):
    return [l for l in re.split(r'\r\n|\r|\n', t) if l.strip()]


def pd_(s, f):
    m = DATE.search(s or '')
    if not m:
        return None
    p1, p2, y = int(m[1]), int(m[2]), int(m[3])
    d, mo = (p1, p2) if f == 'DMY' else (p2, p1)
    try:
        return datetime(y, mo, d, int(m[4] or 0), int(m[5] or 0))
    except ValueError:
        return None


def fmt_of(vals, exp):
    """Detecta si las fechas vienen día-mes o mes-día (por archivo)."""
    a = b = 0
    for s in vals:
        m = re.match(r'^\s*(\d{1,2})[-/](\d{1,2})[-/](\d{4})', s or '')
        if m:
            a += int(m[1]) > 12
            b += int(m[2]) > 12
    if a or b:
        return 'DMY' if a >= b else 'MDY'
    return 'DMY' if all((pd_(s, 'DMY') or exp) <= exp for s in vals) else 'MDY'


def cells(l):
    return [x.strip().strip('"') for x in l.split('\t')]


def parse_review(t, exp):
    L = lines(t)
    H = cells(L[0])
    ix = lambda p: next(i for i, h in enumerate(H) if re.search(p, h))
    iM, iF, iH, iN, iS = ix('^Modo cito'), ix('^Fec'), ix('^Hora'), ix('portas'), ix('^Estado')
    rows = [cells(l) + [''] * len(H) for l in L[1:]]
    f = fmt_of([r[iF] for r in rows], exp)
    out = []
    for r in rows:
        d = pd_(r[iF] + ' ' + (r[iH] or '0:00'), f)
        if d:
            m = re.match(r'\d+', r[iN])
            out.append((d, DAYS(d), r[iM], (int(m[0]) if m else 1) or 1, r[iS].startswith('Final')))
    return pd.DataFrame(out, columns=['t', 'day', 'mode', 'n', 'ok'])


def parse_log(t, exp):
    L = lines(t)
    iF = next(i for i, h in enumerate(cells(L[0])) if 'Fech' in h)
    rows = [cells(l) for l in L[1:]]
    f = fmt_of([r[iF] for r in rows if len(r) > iF], exp)
    ev = []
    for r in reversed(rows):                      # el log viene del más nuevo al más antiguo
        dt = ' '.join(r)
        v = VOL.search(dt)
        if not v:
            continue
        x = next((k for k in R if k in dt), None) or next((EN[k] for k in EN if k in dt), None)
        if not x:
            continue
        d = pd_(r[iF] if DATE.search(r[iF]) else next((c for c in r if DATE.search(c)), ''), f)
        if d:
            ev.append((x, d, float(v[1].replace(',', '.')), float(v[2].replace(',', '.'))))
    return pd.DataFrame(ev, columns=['x', 't', 'frm', 'to']).sort_values('t', kind='stable').reset_index(drop=True)


def parse_an(t, exp):
    L = lines(t)
    H = L[0].split('\t')
    wi = next((i for i, h in enumerate(H) if re.fullmatch(r'"?WBC"?', h.strip())), -1)
    dre = re.compile(r'\s*\d{1,2}[-/]\d{1,2}[-/]\d{4}\s*')
    raw = []
    for l in L[1:]:
        c = l.split('\t')
        i = next((k for k, x in enumerate(c) if dre.fullmatch(x)), -1)
        if i < 3:                                  # el ID puede traer tabuladores: se ubica todo respecto a la fecha
            continue
        k = wi + i - 3
        w = c[k].strip() if wi >= 0 and k < len(c) else 'x'
        raw.append((c[i].strip(), (c[i + 1] if i + 1 < len(c) else '0:00').strip(), c[i - 1].strip(),
                    ''.join(c[:i - 2]).replace('"', '').strip(), w))
    f = fmt_of([r[0] for r in raw], exp)
    out = []
    for d, h, p, id_, w in raw:
        dd = pd_(d + ' ' + h, f)
        if dd:
            out.append((dd, DAYS(dd), p, id_, w in ('', '****')))
    return pd.DataFrame(out, columns=['t', 'day', 'panel', 'id', 'nw'])


@st.cache_data(show_spinner=False)
def load(name, data):
    m = re.search(r'([A-Za-z0-9]{2,4}-[A-Za-z0-9]{6,})', name)
    ser = m[1] if m else name.rsplit('.', 1)[0]
    e = re.search(r'(\d{4})_(\d\d)_(\d\d)', name)
    exp = datetime(int(e[1]), int(e[2]), int(e[3]), 23, 59) if e else datetime.now()
    t = decode(data)
    h = lines(t)[0]
    if 'Modo recue' in h:
        return 'an', ser, parse_an(t, exp)
    if 'Modo cito' in h:
        return 'rev', ser, parse_review(t, exp)
    if re.search('Detall|Detail', h):
        return 'log', ser, parse_log(t, exp)
    return None, ser, None


# ---------------------------------------------------------------- cálculo
def calc(rows, ev, mat, only_ok):
    """Consumo real (de los cambios del log) repartido entre las láminas que usan cada insumo."""
    rs = (rows[rows.ok] if only_ok else rows).reset_index(drop=True)
    out = pd.DataFrame(index=sorted(rs.day.unique()))
    g = rs.groupby('day')
    out['tot'], out['m'] = g.n.sum(), g.size()
    for m in MODES:
        out['sl_' + m] = rs[rs['mode'] == m].groupby('day').n.sum()
    out = out.fillna(0)
    yl = {}
    for x in R:
        u = rs.n * rs['mode'].map(lambda m: x in mat.get(m, ())).astype(int)
        out['u_' + x] = u.groupby(rs.day).sum()
        e = ev[ev.x == x]
        te = e.t.values
        inside = np.zeros(len(rs), bool)
        share = np.zeros(len(rs))
        cs = us = 0.0
        unas = {}
        if len(te) > 1:
            cint = np.maximum(0, e.to.values[:-1] - e.frm.values[1:])
            k = np.searchsorted(te, rs.t.values, side='right') - 1
            inside = (k >= 0) & (k < len(te) - 1) & (u.values > 0)
            U = np.zeros(len(te) - 1)
            np.add.at(U, k[inside], u.values[inside])
            share[inside] = cint[k[inside]] * u.values[inside] / U[k[inside]]
            has = U > 0
            cs, us = cint[has].sum(), U[has].sum()
            for kk in np.where(~has & (cint > 0))[0]:      # consumo sin láminas en el intervalo
                d = DAYS(pd.Timestamp(te[kk + 1]))
                unas[d] = unas.get(d, 0) + cint[kk]
        y = cs / us if us else 0.0
        yl[x] = y
        est = np.where((u.values > 0) & ~inside, u.values * y, 0.0)
        out['c_' + x] = pd.Series(share, index=rs.index).groupby(rs.day).sum()
        for d, v in unas.items():
            out.loc[d, 'c_' + x] = out.loc[d, 'c_' + x] + v if d in out.index else v
        out['es_' + x] = pd.Series(est, index=rs.index).groupby(rs.day).sum()
    return out.fillna(0), yl


def refills(ev):
    """Agrupa los pasos encadenados del mismo día (0,9->1,9->2,9) como una sola recarga."""
    out = []
    for x in R:
        g = None
        for r in ev[ev.x == x].itertuples():
            d = DAYS(r.t)
            if g and g['day'] == d and abs(r.frm - g['to']) < 0.0015:
                g['to'] = r.to
                g['lines'] += 1
            else:
                g = dict(x=x, day=d, frm=r.frm, to=r.to, lines=1)
                out.append(g)
    df = pd.DataFrame(out, columns=['x', 'day', 'frm', 'to', 'lines'])
    df['add'] = df.to - df.frm
    return df


def hemograms(an, names, excl, uniq, nr, d0, d1):
    cols = list(names) + ['H', 'Hi']
    parts = []
    for n in names:
        a = an[n]
        a = a[~a.panel.isin(excl)]
        if nr:
            a = a[~a.nw]
        if uniq:
            a = a[~a.id.isin(['', '0'])].drop_duplicates(['day', 'id'])
        parts.append(a.assign(an=n))
    if not parts:
        return pd.DataFrame(columns=cols)
    a = pd.concat(parts)
    h = a.pivot_table(index='day', columns='an', values='t', aggfunc='count', fill_value=0)
    h = h.reindex(columns=names, fill_value=0)
    h['H'] = h.sum(axis=1)
    h['Hi'] = np.where((h.index >= d0) & (h.index <= d1), h.H, 0)   # % solo donde el SC-120 tiene registros
    return h


def view(S, ans, yld):
    nz = lambda s: s.replace(0, np.nan)
    o = pd.DataFrame(index=S.index)
    o['Láminas'] = S.tot
    if len(ans) > 1:
        for a in ans:
            o[a] = S[a]
    o['Hemogr.'] = S.H
    o['% lám./hemogr.'] = S.tot * 100 / nz(S.Hi)
    o['% muestras c/ lám.'] = S.m * 100 / nz(S.Hi)
    for m in MODES:
        o[m] = S['sl_' + m]
    for x in R:
        v = S['c_' + x] + S['es_' + x]
        o[NM[x]] = (v * (1 if x == 'Porta' else 1000) / nz(S['u_' + x])) if yld else v
    return o.round(2)


def autoassign(eq, an):
    rg = lambda df: (df.t.min(), df.t.max())
    res = {}
    for n, df in an.items():
        a0, a1 = rg(df)
        res[n] = min(eq, key=lambda k: abs(a0 - rg(eq[k]['rev'])[0]) + abs(a1 - rg(eq[k]['rev'])[1]))
    return res


# ---------------------------------------------------------------- alarmas
TRIG = ['Anisocitosis', 'Población dimórfica', 'Histogr RBC anor.', '¿Fragment?', '¿Linf/blast anor?', '¿Blast?',
        '¿Linf atípico?', '¿Gran inmadur?', '¿Desv. izq?', 'Trombocitopenia', 'Neutrofilia', 'Leucocitosis',
        'Leucocitopenia', 'Neutropenia', 'Diagr disp WBC anor', '¿Agr PLT?', 'NRBC present']   # alarmas que mandan a lámina (regla actual)
NUMC = ['RDW-CV', 'WBC', 'PLT', 'HGB', 'MCV', 'MCHC', 'Neu#', 'Linf#', 'Linf%']
ALC = ('WBC', 'RBC', 'PLT', 'VSG')                     # las alarmas van en las últimas columnas WBC/RBC/PLT/VSG


def meta(name):
    m = re.search(r'([A-Za-z0-9]{2,4}-[A-Za-z0-9]{6,})', name)
    e = re.search(r'(\d{4})_(\d\d)_(\d\d)', name)
    return (m[1] if m else name.rsplit('.', 1)[0]), (datetime(int(e[1]), int(e[2]), int(e[3]), 23, 59) if e else datetime.now())


def parse_an2(t, exp):
    """Review del analizador con parámetros y alarmas (columnas finales WBC/RBC/PLT/VSG).
    Se lee con csv (comillas): el ID trae tabuladores. Los parámetros se ubican por encabezado y las alarmas desde el final
    de la fila, porque algunas filas traen celdas extra."""
    R_ = list(csv.reader(lines(t), delimiter='\t', quotechar='"'))
    H = [h.strip() for h in R_[0]]
    pos = lambda k: [i for i, h in enumerate(H) if h == k]
    ia = [pos(k)[-1] - len(H) if len(pos(k)) > 1 else None for k in ALC]      # índice negativo (desde el final)
    ip = {k: (pos(k) or [None])[0] for k in NUMC}
    iI, iP, iF, iH = pos('ID muestr')[0], pos('Panel pr')[0], pos('Fech')[0], pos('Hora')[0]
    raw = []
    for c in R_[1:]:
        c = [x.strip() for x in c]
        m = len(c)
        g = lambda k: c[k] if k is not None and 0 <= k < m else ''
        al = frozenset(a.strip() for k in ia if k is not None and m >= len(H) for a in g(m + k).split(',') if a.strip())
        raw.append([g(iF), g(iH) or '0:00', g(iP), g(iI), al] + [g(ip[k]) for k in NUMC])
    f = fmt_of([r[0] for r in raw], exp)
    out = []
    for r in raw:
        d = pd_(r[0] + ' ' + r[1], f)
        if d:
            out.append([d, DAYS(d), r[2], r[3], r[4]] + r[5:])
    df = pd.DataFrame(out, columns=['t', 'day', 'panel', 'id', 'al'] + NUMC)
    for k in NUMC:
        df[k] = pd.to_numeric(df[k].str.replace(',', '.', regex=False), errors='coerce')
    return df


def parse_rev2(t, exp):
    """Review del SC-120 con ID de muestra y motivo de la citología (reglas de reexamen o manual)."""
    L = lines(t)
    H = cells(L[0])
    ix = lambda p: next(i for i, h in enumerate(H) if re.search(p, h))
    iI, iF, iH, iR, iS, iN = ix('^ID muestr'), ix('^Fec'), ix('^Hora'), ix('^Motivo'), ix('^Estado'), ix('portas')
    rows = [cells(l) + [''] * len(H) for l in L[1:]]
    f = fmt_of([r[iF] for r in rows], exp)
    out = []
    for r in rows:
        d = pd_(r[iF] + ' ' + (r[iH] or '0:00'), f)
        if d:
            m = re.match(r'\d+', r[iN])
            out.append((d, DAYS(d), r[iI].strip(), r[iR].startswith('Contra'), r[iS].startswith('Final'), (int(m[0]) if m else 1) or 1))
    return pd.DataFrame(out, columns=['t', 'day', 'id', 'rules', 'ok', 'n'])


@st.cache_data(show_spinner=False)
def load2(name, data):
    ser, exp = meta(name)
    t = decode(data)
    h = lines(t)[0]
    if 'Modo recue' in h:
        return 'an2', ser, parse_an2(t, exp)
    if 'Modo cito' in h:
        return 'rev2', ser, parse_rev2(t, exp)
    return None, ser, None


def sample_frame(an2, rev2, ans, excl, only_ok=True, only_rules=True):
    """Una fila por muestra de hemograma (sin ID 0 ni paneles excluidos), marcando si tuvo lámina en el SC-120.
    Si la muestra se corrió más de una vez, se toman los parámetros de la primera corrida y las alarmas de todas."""
    A = pd.concat([an2[n].assign(an=n) for n in ans])
    A = A[~A.panel.isin(excl) & ~A.id.isin(['', '0'])]
    A = A[(A.day >= DAYS(rev2.t.min())) & (A.day <= DAYS(rev2.t.max()))]
    A = A.sort_values('t', kind='stable')           # muestra repetida (reexamen): parámetros de la 1.ª corrida, alarmas de todas
    al = A.groupby('id', sort=False).al.agg(lambda x: frozenset().union(*x))
    A = A.drop_duplicates('id', keep='first').assign(al=lambda d: d.id.map(al))
    S = rev2[rev2.ok] if only_ok else rev2
    if only_rules:
        S = S[S.rules]
    return A.assign(slide=A.id.isin(set(S.id))).reset_index(drop=True)


def hard_mask(A, wl=3.5, wh=30.0, pl=100.0, hb=10.0, mv=70.0, mh=30.0):
    return (A.WBC < wl) | (A.WBC > wh) | (A.PLT < pl) | ((A.HGB < hb) & (A.MCV < mv) & (A.MCHC < mh))


def trig_sets(A, trig, silence=(), rcut=0.0):
    """Alarmas disparadoras de cada muestra; 'Anisocitosis' solo cuenta si RDW-CV > rcut."""
    tr = set(trig) - set(silence)
    out = []
    for al, r in zip(A.al, A['RDW-CV']):
        s = al & tr
        out.append(s - {'Anisocitosis'} if 'Anisocitosis' in s and r <= rcut else s)
    return out


def alarm_table(A, trig, hard, ans):
    base = trig_sets(A, trig)
    tot = max(1, int(A.slide.sum()))
    rows = []
    for a in sorted({x for s in A.al for x in s}):
        m = A.al.map(lambda s: a in s)
        r = {'Alarma': a, 'Muestras': int(m.sum()), 'Láminas': int((m & A.slide).sum())}
        r['% con lámina'] = r['Láminas'] * 100 / r['Muestras']
        r['% de las láminas'] = r['Láminas'] * 100 / tot
        r['Solo esta alarma'] = int((pd.Series([s == {a} for s in base], index=A.index) & ~hard & A.slide).sum())
        if len(ans) > 1:
            for n in ans:
                mm = m & (A.an == n)
                r[n + ' muestras'] = int(mm.sum())
                r[n + ' % lám.'] = (mm & A.slide).sum() * 100 / mm.sum() if mm.sum() else np.nan
        rows.append(r)
    return pd.DataFrame(rows).sort_values('Láminas', ascending=False).reset_index(drop=True)


def simulate(A, trig, hard, silence, rcut):
    """Láminas que se evitarían: tenían alguna alarma disparadora y, al silenciarla, ya no queda ninguna ni umbral duro."""
    b, s = trig_sets(A, trig), trig_sets(A, trig, silence, rcut)
    return int((A.slide & ~hard & pd.Series([bool(x) and not y for x, y in zip(b, s)], index=A.index)).sum())


def review_xlsx(df):
    from io import BytesIO
    from openpyxl.worksheet.datavalidation import DataValidation
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w:
        df.to_excel(w, sheet_name='Revisión', index=False)
        ws = w.sheets['Revisión']
        c = df.columns.get_loc('Resultado frotis') + 1
        dv = DataValidation(type='list', formula1='"Sí,No,Dudoso"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f'{ws.cell(2, c).column_letter}2:{ws.cell(2, c).column_letter}{len(df) + 1}')
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = min(40, max(10, max(len(str(x.value or '')) for x in col[:60]) + 2))
        ws.freeze_panes = 'A2'
    return buf.getvalue()


def alarmas(eq2, an2, auto, excl, only_ok):
    if not eq2 or not an2:
        st.info('Sube el Review_ del SC-120 y los Review_ de los analizadores (con las columnas de alarmas al final).')
        return
    sc = st.selectbox('SC-120', list(eq2), key='asc') if len(eq2) > 1 else next(iter(eq2))
    ans = st.multiselect('Analizadores', list(an2), [a for a in an2 if auto.get(a) == sc], key='aan')
    if not ans:
        return
    c1, c2 = st.columns(2)
    only_rules = c1.checkbox('Solo láminas por reglas de reexamen ("Contra reglas")', True, key='arl')
    A = sample_frame(an2, eq2[sc], ans, excl, only_ok, only_rules)
    if A.empty:
        st.warning('No hay hemogramas dentro del período del SC-120.')
        return
    lo, hi = pd.Timestamp(A.day.min()).date(), pd.Timestamp(A.day.max()).date()
    rg = c2.date_input('Rango de fechas', (lo, hi), min_value=lo, max_value=hi, key='arg')
    if len(rg) == 2:
        A = A[(A.day >= str(rg[0])) & (A.day <= str(rg[1]))].reset_index(drop=True)
    with st.expander('Alarmas disparadoras y umbrales duros'):
        trig = st.multiselect('Alarmas que mandan a lámina (según la regla de reexamen)',
                              sorted(set(TRIG) | {x for s in A.al for x in s}), TRIG, key='atr')
        k = st.columns(6)
        hard = hard_mask(A, k[0].number_input('WBC <', value=3.5), k[1].number_input('WBC >', value=30.0),
                         k[2].number_input('PLT <', value=100.0), k[3].number_input('HGB <', value=10.0),
                         k[4].number_input('MCV <', value=70.0), k[5].number_input('MCHC <', value=30.0))
        st.caption('Umbrales duros: WBC bajo/alto, PLT bajo, y HGB+MCV+MCHC bajos a la vez.')
    n, ns = len(A), int(A.slide.sum())
    m = st.columns(3)
    m[0].metric('Hemogramas (muestras únicas)', n)
    m[1].metric('Muestras con lámina', ns)
    m[2].metric('% con lámina', f'{ns * 100 / n:.1f}%')

    st.subheader('Láminas por alarma')
    T = alarm_table(A, trig, hard, ans)
    st.dataframe(T.round(1), use_container_width=True, hide_index=True)
    st.caption('Una muestra puede tener varias alarmas. "Solo esta alarma" = láminas donde era la única disparadora y no había umbral duro.')

    st.subheader('Evolución diaria')
    D = A.groupby('day').agg(Hemogramas=('id', 'size'), Con_lámina=('slide', 'sum'))
    D['% con lámina'] = D.Con_lámina * 100 / D.Hemogramas
    st.line_chart(D['% con lámina'])
    cc = st.columns(2)
    if cc[0].checkbox('Comparar antes / después de un cambio de reglas', key='acmp'):
        dc = str(cc[1].date_input('Fecha del cambio', hi, min_value=lo, max_value=hi, key='acd'))
        for nm, sel in (('Antes', D.index < dc), ('Desde el cambio', D.index >= dc)):
            x = D[sel]
            st.write(f'**{nm}:** {len(x)} días · {int(x.Hemogramas.sum())} hemogramas · {int(x.Con_lámina.sum())} con lámina · '
                     f'**{x.Con_lámina.sum() * 100 / max(1, x.Hemogramas.sum()):.1f}%**')
        st.caption('El día del cambio puede estar incompleto; interpreta con cuidado.')
    st.dataframe(D.round(1), use_container_width=True)

    st.subheader('Simulador')
    s1, s2 = st.columns([3, 1])
    sil = s1.multiselect('Alarmas que dejarían de mandar a lámina', trig, key='asi')
    rc = s2.number_input('Anisocitosis solo si RDW-CV >', value=16.0, key='arc')
    av = simulate(A, trig, hard, sil, rc)
    k = st.columns(3)
    k[0].metric('Láminas evitadas (est.)', av)
    k[1].metric('% actual', f'{ns * 100 / n:.1f}%')
    k[2].metric('% estimado', f'{(ns - av) * 100 / n:.1f}%')
    st.caption('Estimación: asume que cada lámina se debió a las alarmas disparadoras que tenía; las láminas con umbral duro o sin alarma se mantienen.')

    st.subheader('Lista para revisar frotis')
    r1, r2 = st.columns([2, 1])
    al = r1.selectbox('Láminas hechas solo por la alarma', [a for a in trig if a in set(T.Alarma)], key='ara')
    nn = r2.number_input('Muestras', 10, 500, 100, key='arn')
    base = trig_sets(A, trig)
    sel = A[A.slide & ~hard & pd.Series([s == {al} for s in base], index=A.index)]
    st.write(f'{len(sel)} muestras cumplen la condición.')
    if len(sel):
        X = sel.sample(min(int(nn), len(sel)), random_state=7).sort_values('t')
        X = pd.DataFrame({'Analizador': X.an, 'Fecha': X.t.dt.strftime('%d-%m-%Y'), 'Hora': X.t.dt.strftime('%H:%M'),
                          'ID muestra': X.id, 'Panel': X.panel, 'WBC': X.WBC, 'Linf#': X['Linf#'], 'Linf%': X['Linf%'],
                          'Neu#': X['Neu#'], 'HGB': X.HGB, 'PLT': X.PLT, 'RDW-CV': X['RDW-CV'],
                          'Alarmas': X.al.map(lambda s: ', '.join(sorted(s))), 'Resultado frotis': '', 'Observaciones': ''})
        st.download_button('Descargar Excel de revisión', review_xlsx(X), f'revision_{re.sub(r"[^A-Za-z0-9]+", "_", al)}.xlsx',
                           'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    st.subheader('Reporte del día')
    ds = sorted(A.day.unique())
    dd = st.selectbox('Día', ds, len(ds) - 1, key='ard')
    Z = A[A.day == dd]
    cnt = {a: int(Z.slide[Z.al.map(lambda s: a in s)].sum()) for a in trig}
    top = [f'{a} ({v})' for a, v in sorted(cnt.items(), key=lambda x: -x[1]) if v][:5]
    act = st.text_area('Actividades realizadas (opcional)', key='aact', height=100)
    rep = (f'REPORTE DE SEGUIMIENTO – {pd.Timestamp(dd).strftime("%d-%m-%Y")}\n'
           f'Tema: láminas por reglas de reexamen (SC-120 {sc})\n\n'
           + (f'Actividades realizadas:\n{act.strip()}\n\n' if act.strip() else '')
           + f'Hemogramas: {len(Z)} · Muestras con lámina: {int(Z.slide.sum())} '
           f'({Z.slide.sum() * 100 / max(1, len(Z)):.1f}%)\n'
           + (f'Alarmas con más láminas: {", ".join(top)}\n' if top else '')
           + '\nSeguimiento: se repetirá la comparación con la data completa del día siguiente.')
    st.code(rep, language=None)


# ---------------------------------------------------------------- interfaz
def main():
    st.set_page_config('Consumo SC-120', layout='wide')
    st.title('Consumo de insumos · SC-120')
    st.caption('Carga los Review_/Log_ de cada SC-120 y, para el % de láminas, los Review_ de los analizadores conectados.')
    files = st.file_uploader('CSV', type=['csv', 'txt'], accept_multiple_files=True)
    eq, an, eq2, an2 = {}, {}, {}, {}
    for f in files or []:
        data = f.getvalue()
        kind, ser, df = load(f.name, data)
        if kind == 'an':
            an[ser] = df
        elif kind in ('rev', 'log'):
            eq.setdefault(ser, {})[kind] = df
        k2, _, d2 = load2(f.name, data)
        if k2 == 'an2':
            an2[ser] = d2
        elif k2 == 'rev2':
            eq2[ser] = d2
    eq = {k: v for k, v in eq.items() if 'rev' in v}
    for v in eq.values():
        v.setdefault('log', pd.DataFrame(columns=['x', 't', 'frm', 'to']))

    with st.sidebar:
        st.header('Configuración')
        only_ok = st.checkbox('Contar solo láminas "Finaliz"', True)
        st.caption('Insumos que usa cada modo de cito')
        mt = st.data_editor(pd.DataFrame({NM[x]: [x in DEF[m] for m in MODES] for x in R}, index=MODES))
        mat = {m: {x for x in R if mt.loc[m, NM[x]]} for m in MODES}
        st.subheader('Hemogramas')
        uniq = st.checkbox('Solo muestras únicas por día (sin duplicados ni ID 0)', False)
        nr = st.checkbox('Excluir análisis sin resultado (WBC ****)', False)
        panels = sorted({p for df in an.values() for p in df.panel})
        default_ex = [p for p in panels if re.match(r'(?i)^(PLT|Ret|VSG|HMC)', p)]
        excl = st.multiselect('Paneles que NO cuentan como hemograma', panels, default_ex)

    if not eq:
        st.info('Sube al menos un Review_ y su Log_ de un SC-120.')
        return
    auto = autoassign(eq, an) if an else {}
    t1, t2 = st.tabs(['Consumo de insumos', 'Láminas por alarma'])
    with t1:
        consumo(eq, an, auto, mat, only_ok, uniq, nr, excl)
    with t2:
        alarmas(eq2, an2, auto, excl, only_ok)


def consumo(eq, an, auto, mat, only_ok, uniq, nr, excl):
    for tab, sc in zip(st.tabs(list(eq)), eq):
        with tab:
            rows, ev = eq[sc]['rev'], eq[sc]['log']
            ans = st.multiselect('Analizadores de este SC-120', list(an), [a for a in an if auto[a] == sc], key='an' + sc)
            c1, c2, c3 = st.columns(3)
            vw = c1.radio('Vista', ['Mes', 'Día'], horizontal=True, key='v' + sc)
            yld = c2.radio('Mostrar', ['Consumo', 'Rendimiento (mL/lámina)'], horizontal=True, key='y' + sc) != 'Consumo'
            days, yl = calc(rows, ev, mat, only_ok)
            hd = hemograms(an, ans, excl, uniq, nr, DAYS(rows.t.min()), DAYS(rows.t.max()))
            T = days.join(hd, how='outer').fillna(0)
            for c in list(ans) + ['H', 'Hi']:
                if c not in T:
                    T[c] = 0
            if vw == 'Día':
                mes = c3.selectbox('Mes', sorted({d[:7] for d in T.index}), key='m' + sc)
                T = T[T.index.str.startswith(mes)]
                S = T
            else:
                S = T.groupby(T.index.str[:7]).sum()
            S.loc['Total'] = S.sum()
            V = view(S, ans, yld)
            k = st.columns(4)
            k[0].metric('Láminas', int(S.loc['Total', 'tot']))
            k[1].metric('Hemogramas', int(S.loc['Total', 'H']))
            hi = S.loc['Total', 'Hi']
            k[2].metric('% láminas / hemogr.', f"{S.loc['Total', 'tot'] * 100 / hi:.1f}%" if hi else '—')
            k[3].metric('% muestras con lámina', f"{S.loc['Total', 'm'] * 100 / hi:.1f}%" if hi else '—')
            st.subheader('Láminas por modo')
            st.bar_chart(V.drop('Total')[MODES])
            if len(ans) > 1:
                st.subheader('Hemogramas por analizador')
                st.bar_chart(V.drop('Total')[list(ans)])
            st.dataframe(V, use_container_width=True)
            st.caption('Consumo en litros (porta en unidades) = volumen nuevo del cambio anterior − volumen restante '
                       'al cambio siguiente, repartido entre las láminas que usan el insumo. Lo que cae fuera del '
                       'rango del log se estima con el rendimiento promedio.')
            st.subheader('Recargas de insumos')
            rf = refills(ev)
            res = rf.groupby('x').agg(recargas=('day', 'size'), cargado=('add', 'sum'), primera=('day', 'min'),
                                      ultima=('day', 'max')).reindex(R)
            res.index = [NM[x] for x in res.index]
            st.dataframe(res.round(2), use_container_width=True)
            x = st.selectbox('Detalle de', R, format_func=NM.get, key='r' + sc)
            st.dataframe(rf[rf.x == x].drop(columns='x').round(3), use_container_width=True)


if __name__ == '__main__':
    main()
