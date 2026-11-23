#!/usr/bin/env python3
"""Actualiza el Excel LOTERIA PEÑA con los resultados oficiales y los números de cada semana.
Uso: python3 excel_update.py "LOTERIA PEÑA.xlsx" resultados.json boletos.json
- Hoja "SORTEOS APP": una fila por sorteo (fecha, combinación, C, R, apuestas y puntos).
- Hoja2, bloque del mes en curso (B3:F17): puntos de cada sorteo en G / RI / RA.
- Al empezar un mes nuevo: apunta el total del mes cerrado (y el semestre si toca) en el bloque
  anual de la derecha y prepara el bloque del mes nuevo con sus días de sorteo.
"""
import sys, json, datetime as dt
import openpyxl
from openpyxl.styles import Font, Alignment

MESES = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE']
WD = 'LMXJVSD'
INICIO = '2026-11-16'  # primera semana con números guardados por la app

def monday(d): return d - dt.timedelta(days=d.weekday())

def main(xlsx, resp, bolp):
    res = json.load(open(resp, encoding='utf-8'))
    bol = json.load(open(bolp, encoding='utf-8'))
    wb = openpyxl.load_workbook(xlsx)
    log = []

    # ---- puntos por sorteo ----
    rows = []
    for k in sorted(res):
        if k < INICIO: continue
        r = res[k]; d = dt.date.fromisoformat(k)
        b = bol.get(monday(d).isoformat())
        pts = None
        if b and len(b.get('apuestas', [])) >= 3:
            win = set(r['n']); pts = [len([n for n in b['apuestas'][i] if n in win]) for i in range(3)]
        rows.append((d, r, b, pts))

    # ---- hoja SORTEOS APP ----
    ws = wb['SORTEOS APP'] if 'SORTEOS APP' in wb.sheetnames else wb.create_sheet('SORTEOS APP')
    hdr = ['FECHA','DÍA','COMBINACIÓN','C','R','APUESTA GUI','APUESTA RIC','APUESTA RAÚL','REINT.','GUI','RIC','RAÚL']
    if ws['A1'].value != 'FECHA':
        for i, h in enumerate(hdr, 1):
            c = ws.cell(1, i, h); c.font = Font(bold=True); c.alignment = Alignment(horizontal='center')
        for col, w in zip('ABCDEFGHIJKL', [12,6,22,5,5,20,20,20,7,6,6,6]): ws.column_dimensions[col].width = w
    existing = {}
    for rr in range(2, ws.max_row + 1):
        v = ws.cell(rr, 1).value
        if isinstance(v, dt.datetime): v = v.date()
        if isinstance(v, dt.date): existing[v] = rr
    for d, r, b, pts in rows:
        rr = existing.get(d) or ws.max_row + 1
        ap = (b or {}).get('apuestas') or []
        vals = [d, WD[d.weekday()], ' '.join(f'{n:02d}' for n in r['n']), r.get('c'), r.get('r'),
                *[' '.join(f'{n:02d}' for n in ap[i]) if i < len(ap) else None for i in range(3)],
                (b or {}).get('reintegro'), *(pts or [None]*3)]
        for i, v in enumerate(vals, 1):
            c = ws.cell(rr, i, v)
            if i == 1: c.number_format = 'DD/MM/YYYY'
            if i != 3 and i < 6 or i >= 9: c.alignment = Alignment(horizontal='center')
        existing[d] = rr
    log.append(f'SORTEOS APP: {len(rows)} sorteos')

    # ---- Hoja2: bloque del mes en curso ----
    h2 = wb['Hoja2']
    mes_bloque = str(h2['C3'].value or '').strip().upper()
    ultimo = rows[-1][0] if rows else None

    def anio_bloque():
        # el año del bloque es el del último sorteo de ese mes que tengamos
        cand = [d.year for d, *_ in rows if MESES[d.month-1] == mes_bloque]
        return max(cand) if cand else (ultimo.year if ultimo else dt.date.today().year)

    # rollover de mes
    # puntos de los sorteos del mes del bloque (antes de cerrar el mes)
    def apuntar(mb):
        n = 0
        for d, r, b, pts in rows:
            if MESES[d.month-1] != mb or pts is None: continue
            for rr in range(5, 18):
                if h2[f'C{rr}'].value == d.day:
                    for col, v in zip('DEF', pts): h2[f'{col}{rr}'] = v
                    n += 1; break
        return n
    if mes_bloque in MESES: apuntar(mes_bloque)
    incompleto = [h2[f'C{rr}'].value for rr in range(5, 18) if isinstance(h2[f'C{rr}'].value, int) and any(h2[f'{c}{rr}'].value is None for c in 'DEF')]
    if ultimo and mes_bloque in MESES and MESES[ultimo.month-1] != mes_bloque and incompleto:
        log.append(f'No cierro {mes_bloque}: faltan puntos de los días {incompleto}')
    if ultimo and mes_bloque in MESES and MESES[ultimo.month-1] != mes_bloque and not incompleto:
        y = anio_bloque(); m = MESES.index(mes_bloque) + 1
        tot = [0, 0, 0]
        for rr in range(5, 18):
            for j, col in enumerate('DEF'):
                v = h2[f'{col}{rr}'].value
                if isinstance(v, (int, float)): tot[j] += v
        # sección del año en el bloque de la derecha (columna M = año)
        ystart = None
        for rr in range(1, h2.max_row + 2):
            if h2[f'M{rr}'].value == y: ystart = rr
        last = h2.max_row
        if ystart is None:
            ystart = last + 2
            h2[f'M{ystart}'] = y
            for col, t in zip('NOP', ['GUI','RIC','RAUL']): h2[f'{col}{ystart}'] = t
            last = ystart + 1
        else:
            last = ystart + 1
            for rr in range(ystart + 1, h2.max_row + 1):
                if h2[f'M{rr}'].value not in (None, ''): last = rr
                elif rr > last + 1: break
        # no duplicar
        ya = any(str(h2[f'M{rr}'].value).upper() == mes_bloque for rr in range(ystart + 1, last + 1))
        if not ya:
            last += 1
            h2[f'M{last}'] = mes_bloque
            for col, v in zip('NOP', tot): h2[f'{col}{last}'] = v
            log.append(f'Hoja2: total de {mes_bloque} {y} en fila {last}: {tot}')
            if m in (6, 12):
                meses_sem = MESES[0:6] if m == 6 else MESES[6:12]
                st = [0, 0, 0]
                for rr in range(ystart + 1, last + 1):
                    if str(h2[f'M{rr}'].value).upper() in meses_sem or str(h2[f'M{rr}'].value).upper() == 'AGO-SEP':
                        for j, col in enumerate('NOP'):
                            v = h2[f'{col}{rr}'].value
                            if isinstance(v, (int, float)): st[j] += v
                last += 1
                h2[f'M{last}'] = f"{'1º' if m == 6 else '2º'} SEMESTRE {y}"
                for col, v in zip('NOP', st): h2[f'{col}{last}'] = v
                log.append(f'Hoja2: semestre en fila {last}: {st}')
        # preparar el bloque del mes nuevo
        ny, nm = ultimo.year, ultimo.month
        for rr in range(5, 18):
            for col in 'BCDEF': h2[f'{col}{rr}'] = None
        h2['C3'] = MESES[nm-1]
        d = dt.date(ny, nm, 1); rr = 5
        while d.month == nm and rr <= 17:
            if d.weekday() in (0, 3, 5):
                h2[f'B{rr}'] = WD[d.weekday()]; h2[f'C{rr}'] = d.day; rr += 1
            d += dt.timedelta(days=1)
        mes_bloque = MESES[nm-1]
        log.append(f'Hoja2: bloque preparado para {mes_bloque} {ny}')

    # puntos de cada sorteo del mes del bloque
    puestos = 0
    for d, r, b, pts in rows:
        if MESES[d.month-1] != mes_bloque or pts is None: continue
        for rr in range(5, 18):
            if h2[f'C{rr}'].value == d.day:
                for col, v in zip('DEF', pts): h2[f'{col}{rr}'] = v
                puestos += 1; break
    log.append(f'Hoja2: puntos apuntados en {puestos} sorteos de {mes_bloque}')
    falt = [d.isoformat() for d, r, b, pts in rows if pts is None]
    if falt: log.append('Sin números guardados para: ' + ', '.join(falt))

    wb.save(xlsx)
    print('\n'.join(log))

if __name__ == '__main__':
    main(*sys.argv[1:4])
