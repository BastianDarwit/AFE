# -*- coding: utf-8 -*-
"""Avance 2 v6 del TFG · resultados ordenados en embudo.

Por macroproceso: arquitectura → ejemplo del manual → SIPOC; luego la síntesis del estado
del contenido de ambos, el universo del OE2, la estructura de la matriz y la carta Gantt.
Reutiliza las láminas de gen_av5.py y gen_av2.py que no cambian.
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import gen_av2 as g  # noqa: E402
import gen_av5 as v5  # noqa: E402
from gen_av2 import (Slide, table, para, build, WHITE, INK, BLUE_LT, BLUE_S, BLUE_VP,  # noqa: E402
                     BLUE_BG, RULE, TX, TX_SOFT, AUTOR, MX, CW, BODY_T, L_BODY, r, tb, rect,
                     regla, head, imagen, mapper, marca, leyenda, cell, apa, pct)
from gen_av5 import (header, bajada, callout, caja, flecha, TAG1, TAG2, TITLES, T8, T21,  # noqa: E402
                     s_problema1, s_sipoc21, s_ej_c1, s_oe2a, s_oe2b, s_gantt)

# Numeración de figuras en el nuevo orden: el ejemplo C.1 de MP21 pasa a ser la figura 6.
_f5 = v5.fuente


def fuente(s, text, **kw):
    _f5(s, text.replace('Figura 8.', 'Figura 6.'), **kw)


v5.fuente = fuente
g.fuente = fuente

Y0 = BODY_T + 0.08     # inicio del cuerpo cuando no hay bajada


def niveles(s, y, vals):
    """Franja de los cinco niveles del MGO: de macroproceso a tarea."""
    labs = ['Macroproceso', 'Procesos', 'Subprocesos', 'Actividades', 'Tareas']
    gap = 0.30
    w = (CW - 4 * gap) / 5
    h = 0.66
    for k, (lab, v) in enumerate(zip(labs, vals)):
        x = MX + k * (w + gap)
        last = k == 4
        rect(s, 'Nivel ' + lab, x, y, w, h, fill=INK if last else BLUE_VP,
             paras=[para([r(str(v), sz=1800, b=True, color=WHITE if last else INK)],
                         algn='ctr', line=100),
                    para([r(lab, sz=950, color=WHITE if last else INK)], algn='ctr',
                         line=100)],
             ins=(0.02, 0.02, 0.02, 0.02))
        if k < 4:
            flecha(s, 'rightArrow', x + w + 0.05, y + h / 2 - 0.10, gap - 0.10, 0.20,
                   fill=BLUE_S)


# ------------------------------------------------------------------ 6 arquitectura MP8
def s_arq8(num):
    s = Slide(L_BODY, label='OE1 arquitectura MP8')
    header(s, '', num)
    w = 7.60
    box = imagen(s, 'Captura apertura MP8', 'mp8_apertura', MX + (CW - w) / 2, Y0, w=w)
    niveles(s, Y0 + box[3] + 0.20,
            [1, T8['proc'], T8['sub'], T8['act'], T8['task']])
    v5._orig_fuente(s, 'Figura 2. Apertura del macroproceso Gestión de Personas en su manual de '
                       'procesos. Fuente: elaboración propia.')
    return s


# ------------------------------------------------------------------ 7 ejemplo MP8
def s_ej8(num):
    s = Slide(L_BODY, label='OE1 ejemplo MP8')
    header(s, '', num)
    bajada(s, 'Ejemplo del manual MP8: subproceso C.1 Requisición y apertura de vacante.')
    cropf = (60, 55, 1258, 660)
    wf = 4.70
    boxf = imagen(s, 'Captura flujograma C.1', 'mp8_c1_flujograma', MX, 2.60, w=wf)
    f = mapper(boxf, cropf)
    marca(s, f, (266, 331, 316, 354), 1, place='tl', size=0.22)
    marca(s, f, (596, 334, 700, 424), 2, place='tr', size=0.22)
    marca(s, f, (1000, 132, 1252, 632), 3, place='tl', size=0.22)
    cropt = (100, 135, 1220, 462)
    xt = MX + wf + 0.22
    wt = MX + CW - xt
    boxt = imagen(s, 'Captura tareas C.1', 'mp8_c1_tareas', xt, 2.60, w=wt)
    gm = mapper(boxt, cropt)
    marca(s, gm, (840, 176, 1090, 202), 4, place='tr', size=0.22)
    marca(s, gm, (838, 278, 915, 300), 5, place='tr', size=0.22)
    leyenda(s, xt, 2.60 + boxt[3] + 0.10, wt, [
        (1, [r('Código MGO', sz=950, b=True, color=INK),
             r(' de la actividad (C.1.1).', sz=950)], 1),
        (2, [r('Decisión', sz=950, b=True, color=INK),
             r(' del flujo: proceso masivo o individual.', sz=950)], 1),
        (3, [r('Soporte, herramientas, KPI y dueño', sz=950, b=True, color=INK),
             r(' del subproceso.', sz=950)], 1),
        (4, [r('Tarea con responsable', sz=950, b=True, color=INK),
             r(' identificado.', sz=950)], 1),
        (5, [r('Responsable pendiente', sz=950, b=True, color=INK),
             r(' de levantar.', sz=950)], 1),
    ], sz=950, gap=0.02, line_h=0.18)
    callout(s, MX, 2.60 + boxf[3] + 0.22, wf, 0.80,
            [r('Así se documentan los %d subprocesos: ' % T8['sub'], sz=980, b=True, color=INK),
             r('flujograma con sus decisiones y tabla de actividades y tareas, cada una con su '
               'responsable.', sz=980)])
    fuente(s, 'Figura 3. Flujograma y tabla de actividades y tareas del subproceso C.1 (manual '
              'MP8). En el manual, el texto naranjo identifica propuestas por validar.')
    return s


# ------------------------------------------------------------------ 9 arquitectura MP21
def s_arq21(num):
    s = Slide(L_BODY, label='OE1 arquitectura MP21')
    header(s, '', num)
    w = 5.30
    box = imagen(s, 'Captura apertura MP21', 'mp21v2_apertura', MX, Y0 + 0.36, w=w)
    head(s, MX, Y0, w, 'Apertura en el manual: %d procesos y %d subprocesos'
         % (T21['proc'], T21['sub']), sz=1050)
    callout(s, MX, box[1] + box[3] + 0.14, w, 0.78,
            [r('Estructura acordada con el dueño del proceso: ', sz=950, b=True, color=INK),
             r('MP21 agrupa servicios paralelos (aseo, alimentación, seguridad, transporte) que '
               'siguen un mismo ciclo; por eso no se segmentó desde un flujo único, como MP8.',
               sz=950)])
    # ciclo de gestión
    xc = MX + w + 0.35
    wc = MX + CW - xc
    head(s, xc, Y0, wc, 'Ciclo de gestión', sz=1050)
    bw, bh, gx = 1.42, 0.62, wc - 2 * 1.42
    x1, x2 = xc, xc + bw + gx
    ya, ym, yc = Y0 + 0.40, Y0 + 1.36, Y0 + 2.16
    caja(s, 'Proceso A', x1, ya, bw, bh, 'A · Gestión', 'Decide qué servicio', sz1=950, sz2=820)
    caja(s, 'Proceso B', x2, ya, bw, bh, 'B · Planificación', 'Estándar y bases', sz1=950,
         sz2=820)
    caja(s, 'MP19', x2, ym, bw, 0.50, 'MP19 · Compras', 'Licita (fuera del alcance)',
         fill=WHITE, color=INK, line=INK, dash='dash', sz1=880, sz2=760, i2=True)
    caja(s, 'Proceso C', x2, yc, bw, bh, 'C · Administración', 'Asegura y controla', sz1=950,
         sz2=820)
    caja(s, 'Proceso D', x1, yc, bw, bh, 'D · Evaluación', 'Brechas y mejora', sz1=950,
         sz2=820)
    flecha(s, 'rightArrow', x1 + bw + 0.06, ya + bh / 2 - 0.10, gx - 0.12, 0.20)
    flecha(s, 'downArrow', x2 + bw / 2 - 0.10, ya + bh + 0.05, 0.20, ym - ya - bh - 0.10)
    flecha(s, 'downArrow', x2 + bw / 2 - 0.10, ym + 0.50 + 0.05, 0.20, yc - ym - 0.50 - 0.10)
    flecha(s, 'leftArrow', x1 + bw + 0.06, yc + bh / 2 - 0.10, gx - 0.12, 0.20)
    flecha(s, 'upArrow', x1 + bw / 2 - 0.10, ya + bh + 0.05, 0.20, yc - ya - bh - 0.10,
           fill=INK)
    tb(s, 'Retro', x1 + bw / 2 + 0.14, ya + bh + 0.30, bw / 2 + gx - 0.20, 0.60,
       [para([r('D retroalimenta a A y B', sz=820, i=True, color=INK)], line=100)],
       anchor='ctr')
    niveles(s, 5.82, [1, T21['proc'], T21['sub'], T21['act'], T21['task']])
    v5._orig_fuente(s, 'Figura 5. Apertura y ciclo de gestión del macroproceso Gestión de '
                       'Servicios Generales. Fuente: elaboración propia sobre su manual.')
    return s


# ------------------------------------------------------------------ 12 síntesis
def s_sintesis(num):
    s = Slide(L_BODY, label='OE1 sintesis')
    header(s, '', num)
    wl = 5.55
    col = [wl * 0.34, wl * 0.24, wl * 0.24, wl * 0.18]
    hdr = [('Nivel', 'l'), ('MP8', 'ctr'), ('MP21', 'ctr'), ('Total', 'ctr')]

    def fila(lab, a, b, bold=False):
        c = INK if bold else TX
        return [cell(lab, b=True, color=INK), cell(str(a), algn='ctr', b=bold, color=c),
                cell(str(b), algn='ctr', b=bold, color=c),
                cell(str(a + b), algn='ctr', b=True, color=INK)]

    head(s, MX, Y0, wl, 'Estructura en los cinco niveles del MGO', sz=1100)
    rows = [fila('Macroproceso', 1, 1), fila('Procesos', T8['proc'], T21['proc']),
            fila('Subprocesos', T8['sub'], T21['sub']),
            fila('Actividades', T8['act'], T21['act']),
            fila('Tareas', T8['task'], T21['task'], bold=True)]
    s.add(table(s.nid(), 'Tabla estructura', MX, Y0 + 0.34, col, apa(hdr, rows, sz=1000),
                [0.28] + [0.27] * 5))

    def v(n, tot):
        return '%d (%s %%)' % (n, pct(n, tot))

    def fila2(lab, k, b8=False, b21=False):
        return [cell(lab), cell(v(T8[k], T8['task']), algn='ctr', b=b8, color=INK if b8 else TX),
                cell(v(T21[k], T21['task']), algn='ctr', b=b21, color=INK if b21 else TX),
                cell(str(T8[k] + T21[k]), algn='ctr', b=True, color=INK)]

    y2 = Y0 + 0.34 + 0.28 + 0.27 * 5 + 0.22
    head(s, MX, y2, wl, 'Contenido de las tareas', sz=1100)
    rows2 = [fila2('Levantadas', 'task_lev', b8=True), fila2('Propuestas', 'task_prop', b21=True),
             fila2('Con responsable', 'resp_ok', b21=True)]
    s.add(table(s.nid(), 'Tabla contenido', MX, y2 + 0.34, col,
                apa([('Tareas', 'l'), ('MP8', 'ctr'), ('MP21', 'ctr'), ('Total', 'ctr')], rows2,
                    sz=1000), [0.28] + [0.27] * 3))
    # lectura
    xr = MX + wl + 0.30
    wr = MX + CW - xr
    head(s, xr, Y0, wr, 'Lectura', sz=1100)
    callout(s, xr, Y0 + 0.36, wr, 1.02,
            [r('MP8 · levantado en terreno. ', sz=980, b=True, color=INK),
             r('Falta: %d responsables y las fases F15 a F18 (nómina, clima, desvinculación y '
               'datos).' % T8['resp_pend'], sz=980)])
    callout(s, xr, Y0 + 1.52, wr, 1.02,
            [r('MP21 · estructura y responsables definidos. ', sz=980, b=True, color=INK),
             r('Falta: validar %d tareas propuestas con el dueño.' % T21['task_prop'], sz=980)])
    rect(s, 'CAP fondo', xr, Y0 + 2.70, wr, 1.02, fill=WHITE, line=INK, line_w=1.25)
    tb(s, 'CAP', xr + 0.10, Y0 + 2.70, wr - 0.20, 1.02,
       [para([r('CAP: ', sz=980, b=True, color=INK),
              r('estructura completa en ambos; el indicador se declara al cierre de la '
                'validación formal (A1.6).', sz=980)], line=104)], anchor='ctr')
    fuente(s, 'Fuente: elaboración propia a partir de los manuales MP8 y MP21. Levantado: texto '
              'normal; propuesto: texto naranjo.')
    return s


# ------------------------------------------------------------------ 2 contenidos
def s_contenidos(num):
    s = Slide(L_BODY, label='Contenidos')
    header(s, 'Contenidos', num)
    items = [
        ('1', 'Problema', 'Organización, unidad de estudio y descripción del problema', '3–4'),
        ('2', 'Objetivos', 'Objetivo general y objetivos específicos con sus indicadores', '5'),
        ('3', 'Resultados del OE1', 'MP8 y MP21: arquitectura, ejemplo del manual, SIPOC y '
         'estado del contenido', '6–12'),
        ('4', 'Resultados del OE2', 'Universo de tareas y estructura de la matriz de riesgos',
         '13–14'),
        ('5', 'Carta Gantt', 'Nivel de cumplimiento del plan de trabajo', '15'),
    ]
    y = BODY_T + 0.10
    rh = 0.78
    for n, t1, t2, sl in items:
        destacado = n == '3'
        if destacado:
            rect(s, 'Fondo destacado', MX, y, CW, rh - 0.06, fill=BLUE_BG)
        rect(s, 'Num %s' % n, MX + 0.12, y + (rh - 0.06 - 0.46) / 2, 0.46, 0.46, fill=INK,
             geom='ellipse', paras=[para([r(n, sz=1400, b=True, color=WHITE)], algn='ctr')],
             ins=(0, 0, 0, 0))
        tb(s, 'Seccion %s' % n, MX + 0.80, y + 0.04, 6.50, rh - 0.14,
           [para([r(t1, sz=1500, b=True, color=INK)]),
            para([r(t2, sz=1050, color=TX)], space_before=2)], anchor='ctr')
        tb(s, 'Laminas %s' % n, MX + CW - 1.40, y + 0.04, 1.30, rh - 0.14,
           [para([r('Láminas ' if '–' in sl else 'Lámina ', sz=900, color=TX_SOFT),
                  r(sl, sz=1100, b=True, color=INK)], algn='r')], anchor='ctr')
        if not destacado:
            regla(s, y + rh - 0.04)
        y += rh
    return s


# ------------------------------------------------------------------ armado
def make(out):
    seq = [s_problema1, g.s_problema2, g.s_objetivos,
           s_arq8, s_ej8, g.s_oe1c,
           s_arq21, s_ej_c1, s_sipoc21,
           s_sintesis, s_oe2a, s_oe2b, s_gantt]
    TITLES.clear()
    TITLES.update({
        6: (TAG1, 'Resultado: arquitectura del macroproceso MP8'),
        7: (TAG1, 'Resultado: MP8 documentado hasta la tarea'),
        8: (TAG1, 'Resultado: cómo se articula el macroproceso MP8'),
        9: (TAG1, 'Resultado: arquitectura del macroproceso MP21'),
        10: (TAG1, 'Resultado: MP21 controla al contratista en ejecución'),
        11: (TAG1, 'Resultado: cómo se articula el macroproceso MP21'),
        12: (TAG1, 'Resultado: MP8 levantado; MP21 por validar'),
        13: (TAG2, 'Resultado: universo de 248 tareas para el OE2'),
        14: (TAG2, 'Resultado: estructura de la matriz de riesgos'),
    })
    slides = [g.s_portada(), s_contenidos(2)]
    for k, fn in enumerate(seq, 3):
        slides.append(fn(k))
    slides.append(g.s_cierre())
    build(g.BASE, out, slides, title='Avance 2 TFG - ' + AUTOR, author=AUTOR, media=g.MEDIA)
    return slides


if __name__ == '__main__':
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(g.ROOT, 'out', 'Avance_2_v6.pptx')
    sl = make(out)
    print('OK', out, len(sl), 'láminas')
