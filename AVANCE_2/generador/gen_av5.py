# -*- coding: utf-8 -*-
"""Avance 2 v5 del TFG · lámina 2 de contenidos.

Cada lámina de resultados lleva un título «Resultado: …» que enuncia lo que muestra, con
una etiqueta del objetivo específico sobre el título. La presentación no hace referencia a
versiones anteriores de sí misma ni de los manuales. Datos: stats_mp8.json (manual MP8) y
stats_mp21_v2.json (manual MP21).
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import gen_av2 as g  # noqa: E402
from gen_av2 import (Slide, shape, table, para, picture, REL_IMAGE, build, WHITE,  # noqa: E402
                     INK, BLUE, BLUE_LT, BLUE_S, BLUE_P, BLUE_VP, BLUE_BG, RULE, TX,
                     TX_SOFT, TITLE_GRAY, TFG, AUTOR, MX, CW, BODY_T, L_BODY, L_PHOTO,
                     r, tb, rect, regla, fuente, head, imagen, mapper, marca, leyenda, cell,
                     apa, pct)

ROOT = g.ROOT
S8 = g.S8
T8 = g.T8
S21 = json.load(open(os.path.join(ROOT, 'stats_mp21_v2.json')))
T21 = S21['tot']
# Actividades levantadas / propuestas (color del nombre de la actividad en el manual)
ACT8 = (37, 22)
ACT21 = (11, 13)
N_OE1 = 10

# --- títulos de resultados -----------------------------------------------------------
TAG1 = 'Desarrollo y resultados del OE1 · Arquitectura de procesos'
TAG2 = 'Desarrollo y resultados del OE2 · Evaluación de riesgos'
TITLES = {}   # num -> (etiqueta, título)
_orig_header = g.header


def header(s, titulo, num):
    tag, tit = TITLES.get(num, (None, titulo))
    _orig_header(s, tit, num)
    if tag:
        tb(s, 'Etiqueta OE', MX, 1.18, 5.60, 0.26,
           [para([r(tag, sz=1050, b=True, color=BLUE_LT)])], anchor='ctr')


g.header = header

_orig_fuente = g.fuente


def fuente(s, text, **kw):
    _orig_fuente(s, text.replace('manual MP8 v1', 'manual MP8'), **kw)


g.fuente = fuente


def oe1(k):
    return 'Desarrollo y resultados del OE1 (%d de %d)' % (k, N_OE1)


def bajada(s, text, y=None):
    tb(s, 'Bajada', MX, (BODY_T - 0.04) if y is None else y, CW, 0.32,
       [para([r(text, sz=1100, b=True, color=INK)])], anchor='ctr')


def callout(s, x, y, w, h, runs, sz=1000):
    """Recuadro de mensaje: fondo azul claro con barra lateral."""
    rect(s, 'Mensaje fondo', x, y, w, h, fill=BLUE_BG)
    rect(s, 'Mensaje barra', x, y, 0.06, h, fill=INK)
    tb(s, 'Mensaje', x + 0.14, y, w - 0.20, h, [para(runs, line=106)], anchor='ctr')


def caja(s, name, x, y, w, h, t1, t2=None, fill=INK, color=WHITE, line=None, dash=None,
         sz1=1050, sz2=900, i2=False):
    ps = [para([r(t1, sz=sz1, b=True, color=color)], algn='ctr', line=100)]
    if t2:
        ps.append(para([r(t2, sz=sz2, i=i2, color=color)], algn='ctr', line=100,
                       space_before=2))
    s.add(shape(s.nid(), name, x, y, w, h, fill=fill, line=line, line_w=1.25, dash=dash,
                paras=ps, anchor='ctr', ins=(0.06, 0.03, 0.06, 0.03), autofit=False))


def flecha(s, geom, x, y, w, h, fill=BLUE_S):
    s.add(shape(s.nid(), 'Flecha', x, y, w, h, geom=geom, fill=fill, line=None,
                adj=[('1', 50000), ('2', 50000)]))


# ------------------------------------------------------------------ 2 título
def s_titulo():
    s = Slide(L_PHOTO, label='Título del trabajo')
    s.rels.append(('rId9', REL_IMAGE, '../media/image4.jpeg'))
    s.add(picture(s.nid(), 'Imagen institucional', 0.74, 1.82, 3.90, 4.87, 'rId9'))
    tb(s, 'Estudiante', 6.28, 1.13, 2.21, 0.24,
       [para([r('| ', sz=900, color='7F7F7F'), r('Estudiante: ', sz=900, b=True, color=INK),
              r(AUTOR, sz=900, color=INK), r(' |', sz=900, color='7F7F7F')], algn='ctr')],
       anchor='ctr', ins=(0.02, 0.01, 0.02, 0.01))
    tb(s, 'Titulo del trabajo', 5.00, 1.82, 4.50, 2.50,
       [para([r(TFG, sz=2000, b=True, color=TITLE_GRAY)], line=105)])
    regla(s, 4.46, x=5.04, w=4.40)
    tb(s, 'Precision', 5.00, 4.58, 4.50, 2.00,
       [para([r('Título sin modificaciones', sz=1200, b=True, color=INK),
              r(' respecto del anteproyecto aprobado y del Avance 1.', sz=1200)], line=110),
        para([r('El trabajo comprende el ', sz=1200), r('diseño', sz=1200, b=True),
              r(' del sistema de gestión de riesgos; su implementación no forma parte del '
                'alcance.', sz=1200)], line=110, space_before=8),
        para([r('Este avance reporta la arquitectura de procesos consolidada en los '
                'manuales de MP8 y MP21.', sz=1200)], line=110, space_before=8)])
    return s



# ------------------------------------------------------------------ 2 contenidos
def s_contenidos(num):
    s = Slide(L_BODY, label='Contenidos')
    header(s, 'Contenidos', num)
    items = [
        ('1', 'Problema', 'Organización, unidad de estudio y descripción del problema', '3–4'),
        ('2', 'Objetivos', 'Objetivo general y objetivos específicos con sus indicadores', '5'),
        ('3', 'Resultados del OE1', 'Arquitectura de procesos de MP8 y MP21 en cinco niveles',
         '6–15'),
        ('4', 'Resultados del OE2', 'Universo de tareas y estructura de la matriz de riesgos',
         '16–17'),
        ('5', 'Carta Gantt', 'Nivel de cumplimiento del plan de trabajo', '18'),
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

# ------------------------------------------------------------------ 3 problema 1
def s_problema1(num):
    s = Slide(L_BODY, label='Problema 1')
    header(s, 'Problema (1 de 2) · Organización y unidad de estudio', num)
    tb(s, 'Organizacion', MX, BODY_T, CW, 0.72,
       [para([r('Celulosa Arauco y Constitución S.A.', sz=1250, b=True, color=INK),
              r(', Negocio de Celulosa: cinco plantas en Chile (Arauco, Constitución, Licancel, '
                'Nueva Aldea y Valdivia) y una en Argentina. Sus procesos se ordenan en '
                '21 macroprocesos bajo el ', sz=1250),
              r('Modelo de Gestión Operacional (MGO)', sz=1250, b=True),
              r('.', sz=1250)], line=108)])
    crop = (110, 175, 1170, 642)
    box = imagen(s, 'Captura mapa macroprocesos', 'mp8_mapa_macroprocesos', MX, 3.05, w=6.30)
    f = mapper(box, crop)
    marca(s, f, (592, 220, 684, 325), 1, place='tr')
    marca(s, f, (1026, 530, 1122, 635), 2, place='tr')
    tb(s, 'Rotulo figura', MX, 3.05 + box[3] + 0.04, 6.30, 0.26,
       [para([r('Figura 1. ', sz=900, b=True, color=INK),
              r('Mapa de macroprocesos del Negocio de Celulosa según el MGO.', sz=900, i=True,
                color=TX_SOFT)])], anchor='ctr')
    xr = 7.05
    wr = MX + CW - xr
    head(s, xr, 3.05, wr, 'Unidad de estudio', sz=1250)
    leyenda(s, xr, 3.42, wr, [
        (1, [r('MP8 Gestión de Personas. ', sz=1050, b=True, color=INK),
             r('Transversal a los negocios de celulosa, madera y forestal.', sz=1050)], 3),
        (2, [r('MP21 Servicios Generales. ', sz=1050, b=True, color=INK),
             r('Soporte; transversal a celulosa y forestal.', sz=1050)], 3),
    ], gap=0.12)
    tb(s, 'Unidad', xr, 5.05, wr, 1.20,
       [para([r('Unidad de análisis: ', sz=1050, b=True, color=INK),
              r('el nivel corporativo que diseña ambos macroprocesos para las cinco '
                'plantas.', sz=1050)], line=105)])
    fuente(s, 'Fuente: elaboración propia sobre el mapa de macroprocesos de los manuales MP8 y '
              'MP21 y la sección 1.5 del informe.')
    return s


# ------------------------------------------------------------------ OE1 1: niveles
def s_oe1_niveles(num):
    s = Slide(L_BODY, label='OE1 niveles')
    header(s, oe1(1), num)
    tb(s, 'Actividades', MX, BODY_T, CW, 0.70,
       [para([r('Actividades realizadas. ', sz=1150, b=True, color=INK),
              r('A1.3.2 segmentación de MP8 desde su flujo único de información; A1.4 '
                'estructura y segmentación de MP21 con su dueño; A1.5 consolidación en '
                'manuales con el formato MGO aprobado: MP8 (58 láminas) y MP21 '
                '(31 láminas).', sz=1150)], line=106)])
    crop = (150, 135, 1128, 548)
    w = 5.55
    box = imagen(s, 'Captura apertura MP8', 'mp8_apertura', MX, 3.02, w=w)
    f = mapper(box, crop)
    marca(s, f, (158, 143, 346, 222), 1, place='tl')
    marca(s, f, (158, 224, 346, 286), 2, place='bl')
    marca(s, f, (546, 143, 734, 545), 3, place='tr')
    ly = 3.02 + box[3] + 0.10
    leyenda(s, MX, ly, w, [
        (1, [r('Proceso', sz=1000, b=True, color=INK),
             r(' (A a E), con su ficha y SIPOC.', sz=1000)], 1),
        (2, [r('Subproceso', sz=1000, b=True, color=INK),
             r(' (A.1), con flujograma, actividades y tareas.', sz=1000)], 1),
        (3, [r('Atracción y selección', sz=1000, b=True, color=INK),
             r(': cinco subprocesos (fases F3 a F10 del flujo).', sz=1000)], 1),
    ], sz=1000, gap=0.03, line_h=0.19)
    xr = MX + w + 0.30
    wr = MX + CW - xr
    head(s, xr, 3.02, wr, 'Elementos por nivel MGO', sz=1150)
    col = [1.27, 0.88, 0.88]
    rows = []
    for lab, a, b in (('Macroproceso', 1, 1), ('Procesos', T8['proc'], T21['proc']),
                      ('Subprocesos', T8['sub'], T21['sub']),
                      ('Actividades', T8['act'], T21['act']),
                      ('Tareas', T8['task'], T21['task'])):
        last = lab == 'Tareas'
        rows.append([cell(lab, b=True, color=INK),
                     cell(str(a), algn='ctr', b=last, color=INK if last else TX),
                     cell(str(b), algn='ctr', b=last, color=INK if last else TX)])
    t = apa([('Nivel', 'l'), ('MP8', 'ctr'), ('MP21', 'ctr')], rows)
    s.add(table(s.nid(), 'Tabla niveles', xr, 3.38, col, t, [0.32] + [0.30] * 5))
    tb(s, 'Cobertura', xr, 5.32, wr, 1.18,
       [para([r('MP8: ', sz=1000, b=True, color=INK),
              r('14 de las 18 fases del flujo único segmentadas; restan nómina, clima, '
                'desvinculación y datos (F15 a F18).', sz=1000)], line=105),
        para([r('MP21: ', sz=1000, b=True, color=INK),
              r('los cinco niveles completos, de sus 4 procesos a sus 56 tareas.', sz=1000)],
             line=105, space_before=4)])
    fuente(s, 'Figura 2. Apertura del macroproceso MP8 en su manual de procesos. Fuente: '
              'elaboración propia.')
    return s


# ------------------------------------------------------------------ OE1 4: ajuste metodológico
def s_ajuste(num):
    s = Slide(L_BODY, label='OE1 ajuste metodologico')
    header(s, oe1(4), num)
    tb(s, 'Bajada', MX, BODY_T - 0.04, CW, 0.46,
       [para([r('Ajuste metodológico. ', sz=1100, b=True, color=INK),
              r('La jerarquía se acordó con el dueño del proceso en lugar de derivarse de un '
                'flujo único de información, como establece la sección 2.3.1 del informe.',
                sz=1100)], line=104)], anchor='ctr')
    bw = (CW - 3 * 0.40) / 4
    xs = [MX + k * (bw + 0.40) for k in range(4)]
    # fila planificada
    head(s, MX, 2.72, CW, 'Planificado (informe, sección 2.3.1) · aplicado en MP8', sz=1000,
         color=TX_SOFT)
    y1, h1 = 2.99, 0.44
    caja(s, 'Plan 1', xs[0], y1, bw, h1, 'Flujo único', 'de información (SIPOC)',
         fill=BLUE_BG, color=TX, line=RULE)
    caja(s, 'Plan 2', xs[1], y1, 2 * bw + 0.40, h1, 'Segmentación en cinco niveles',
         'derivada del flujo levantado', fill=BLUE_BG, color=TX, line=RULE)
    caja(s, 'Plan 3', xs[3], y1, bw, h1, 'Validación', 'con los dueños de proceso',
         fill=BLUE_BG, color=TX, line=RULE)
    for k in (0, 2):
        flecha(s, 'rightArrow', xs[k + 1] - 0.33 if k == 0 else xs[3] - 0.33, y1 + 0.11,
               0.26, 0.22, fill=RULE)
    # fila aplicada
    head(s, MX, 3.55, CW, 'Aplicado en MP21', sz=1000)
    y2, h2 = 3.82, 0.52
    caja(s, 'Apl 1', xs[0], y2, bw, h2, 'Estructura acordada', 'con el dueño del proceso',
         fill=INK)
    caja(s, 'Apl 2', xs[1], y2, bw, h2, 'Contraste APQC', 'PCF 8.0 por subproceso',
         fill=BLUE_VP, color=INK)
    caja(s, 'Apl 3', xs[2], y2, bw, h2, 'Actividades y tareas', 'levantadas y propuestas',
         fill=BLUE_VP, color=INK)
    caja(s, 'Apl 4', xs[3], y2, bw, h2, 'Validación formal', 'pendiente (A1.6)',
         fill=WHITE, color=INK, line=INK, dash='dash', i2=True)
    for k in range(3):
        flecha(s, 'rightArrow', xs[k + 1] - 0.33, y2 + 0.15, 0.26, 0.22, fill=BLUE_S)
    # por qué
    yb = 4.52
    wl = 5.30
    head(s, MX, yb, wl, 'Por qué el ajuste', sz=1100)
    leyenda(s, MX, yb + 0.34, wl, [
        (1, [r('Sin cadena única: ', sz=1000, b=True, color=INK),
             r('MP21 agrupa servicios paralelos (aseo, alimentación, seguridad, transporte, '
               'entre otros) que siguen un mismo ciclo de gestión.', sz=1000)], 2),
        (2, [r('El ciclo es el «cómo» común; ', sz=1000, b=True, color=INK),
             r('cada servicio es el «qué» y entra en actividades y tareas.', sz=1000)], 2),
        (3, [r('Trazabilidad conservada: ', sz=1000, b=True, color=INK),
             r('cada elemento declara su origen (levantado, propuesto o pendiente) y cada '
               'subproceso su equivalencia APQC.', sz=1000)], 2),
    ], sz=1000, gap=0.05, line_h=0.19)
    # tabla APQC
    xr = MX + wl + 0.30
    wr = MX + CW - xr
    head(s, xr, yb, wr, 'Equivalencia APQC PCF 8.0', sz=1100)
    eq = [('A.1 Evaluación de criticidad', '10.1.1.3 · 4.2.1.6'),
          ('B.1 Definición de estándares', '4.2.1.2'),
          ('B.2 Bases técnicas', '4.2.1.1 · 4.2.1.4'),
          ('B.3 Estimación presupuestaria', '10.1.2.2 · 4.2.1.5'),
          ('C.1 Aseguramiento del servicio', '10.1.3 · 10.1.4 · 4.2.3.4'),
          ('C.2 Control y mejora', '4.2.5.2 · 4.2.5.3'),
          ('D.1 y D.2 Evaluación', '4.2.5.2 · 4.2.1.6')]
    rows = [[cell(a, sz=850), cell(b, sz=850, color=INK)] for a, b in eq]
    t = apa([('Subproceso', 'l'), ('Código APQC', 'l')], rows, sz=850)
    s.add(table(s.nid(), 'Tabla APQC', xr, yb + 0.32, [wr * 0.56, wr * 0.44], t,
                [0.22] + [0.205] * len(eq)))
    fuente(s, 'Fuente: elaboración propia sobre la sección 2.3.1 del informe y el manual de '
              'procesos MP21.')
    return s


# ------------------------------------------------------------------ OE1 5: apertura MP21
def s_apertura21(num):
    s = Slide(L_BODY, label='OE1 apertura MP21')
    header(s, oe1(5), num)
    bajada(s, 'Apertura del macroproceso MP21: 4 procesos y 8 subprocesos.')
    crop = (250, 135, 1030, 415)
    w = 6.05
    box = imagen(s, 'Captura apertura MP21', 'mp21v2_apertura', MX, 2.62, w=w)
    f = mapper(box, crop)
    marca(s, f, (255, 140, 443, 279), 1, place='tl')
    marca(s, f, (449, 140, 637, 412), 2, place='bl')
    marca(s, f, (643, 140, 831, 340), 3, place='tr')
    marca(s, f, (837, 140, 1025, 340), 4, place='br')
    xr = MX + w + 0.28
    wr = MX + CW - xr
    head(s, xr, 2.62, wr, 'Qué hace cada proceso', sz=1100)
    leyenda(s, xr, 2.98, wr, [
        (1, [r('A decide ', sz=950, b=True, color=INK),
             r('qué servicio se necesita, según su criticidad.', sz=950)], 2),
        (2, [r('B fija ', sz=950, b=True, color=INK),
             r('el estándar, las bases técnicas y el presupuesto que se entregan a MP19.',
               sz=950)], 3),
        (3, [r('C asegura y controla ', sz=950, b=True, color=INK),
             r('la ejecución del servicio contratado.', sz=950)], 2),
        (4, [r('D evalúa ', sz=950, b=True, color=INK),
             r('las brechas y retroalimenta a A y B.', sz=950)], 2),
    ], sz=950, gap=0.06, line_h=0.18)
    yc = 2.62 + box[3] + 0.22
    callout(s, MX, yc, CW, 0.62,
            [r('Un ciclo de gestión, no la vida de un contrato: ', sz=1000, b=True, color=INK),
             r('los cuatro procesos se repiten para cada servicio no operacional (aseo, '
               'alimentación, seguridad, transporte, entre otros) y en cada planta.', sz=1000)])
    tb(s, 'Nota nombres', MX, yc + 0.70, CW, 0.28,
       [para([r('Los nombres de C.2 y D.2 son propuestos y se validan con el dueño del proceso.',
                sz=900, i=True, color=TX_SOFT)])], anchor='ctr')
    fuente(s, 'Figura 5. Apertura del macroproceso Gestión de Servicios Generales en su manual '
              'de procesos. Fuente: elaboración propia.')
    return s


# ------------------------------------------------------------------ OE1 6: ciclo
def s_ciclo(num):
    s = Slide(L_BODY, label='OE1 ciclo MP21')
    header(s, oe1(6), num)
    bajada(s, 'El ciclo se cierra en D y su límite con Compras y contratos (MP19) es explícito.')
    bw, bh = 2.20, 0.74
    xa, xb = MX + 0.05, MX + 2.95
    ya, ym, yc = 2.86, 4.08, 5.24
    caja(s, 'Proceso A', xa, ya, bw, bh, 'A · Gestión estratégica',
         'Prioriza qué servicio se necesita')
    caja(s, 'Proceso B', xb, ya, bw, bh, 'B · Planificación',
         'Estándar, bases técnicas y presupuesto')
    caja(s, 'MP19', xb, ym, bw, 0.66, 'MP19 · Compras y contratos',
         'Licita y adjudica (fuera del alcance)', fill=WHITE, color=INK, line=INK,
         dash='dash', sz1=980, sz2=850, i2=True)
    caja(s, 'Proceso C', xb, yc, bw, bh, 'C · Administración de servicios',
         'Asegura y controla la ejecución', sz1=1000)
    caja(s, 'Proceso D', xa, yc, bw, bh, 'D · Evaluación del servicio',
         'Identifica brechas y define planes')
    flecha(s, 'rightArrow', xa + bw + 0.12, ya + 0.25, xb - xa - bw - 0.24, 0.26)
    flecha(s, 'downArrow', xb + bw / 2 - 0.13, ya + bh + 0.06, 0.26, ym - ya - bh - 0.12)
    flecha(s, 'downArrow', xb + bw / 2 - 0.13, ym + 0.66 + 0.06, 0.26, yc - ym - 0.66 - 0.12)
    flecha(s, 'leftArrow', xa + bw + 0.12, yc + 0.25, xb - xa - bw - 0.24, 0.26)
    flecha(s, 'upArrow', xa + 0.30, ya + bh + 0.06, 0.26, yc - ya - bh - 0.12, fill=INK)
    tb(s, 'Retro', xa + 0.64, ya + bh + 0.22, 1.60, 0.90,
       [para([r('Retroalimentación (D.2.2): ', sz=900, b=True, color=INK),
              r('actualiza la criticidad (A.1) y los estándares (B.1).', sz=900)], line=104)],
       anchor='ctr')
    # el contrato dentro del ciclo
    x2 = MX + 5.45
    w2 = MX + CW - x2
    head(s, x2, 2.62, w2, 'El contrato dentro del ciclo', sz=1050)
    hitos = [('Entrega de bases y presupuesto a Compras', 'B.3.3 Formalización del '
              'requerimiento', True),
             ('Recepción del contrato adjudicado', 'C.1.1 Inicio del servicio', True),
             ('Control laboral y previsional del contratista', 'C.1.3 Control de cumplimiento '
              'laboral', False),
             ('Multas y sanciones', 'C.1.5 Aplicación de sanciones', False),
             ('Auditoría del servicio', 'C.2.1 Control y auditoría QA QC', False),
             ('Renovar, modificar o terminar', 'C.2.3 Continuidad del contrato', True)]
    rows = [[cell(a, sz=850, b=m, color=INK if m else TX), cell(c, sz=850, color=INK)]
            for a, c, m in hitos]
    t = apa([('Hito del contrato', 'l'), ('Dónde ocurre', 'l')], rows, sz=850)
    s.add(table(s.nid(), 'Tabla hitos', x2, 2.94, [w2 * 0.52, w2 * 0.48], t,
                [0.24] + [0.42] * len(hitos)))
    tb(s, 'Nota hitos', x2, 2.94 + 0.24 + 0.42 * len(hitos) + 0.06, w2, 0.40,
       [para([r('En negrita, los hitos que son interfaz con MP19: puntos de riesgo para el OE2.',
                sz=850, i=True, color=TX_SOFT)], line=104)])
    fuente(s, 'Figura 6. Ciclo de gestión del macroproceso MP21. Fuente: elaboración propia sobre '
              'el manual de procesos MP21.')
    return s


# ------------------------------------------------------------------ OE1 7: SIPOC MP21
def s_sipoc21(num):
    s = Slide(L_BODY, label='OE1 SIPOC MP21')
    header(s, oe1(7), num)
    bajada(s, 'SIPOC del macroproceso MP21: qué recibe, qué entrega y con quién se relaciona.')
    crop = (45, 112, 1245, 700)
    w = 6.05
    box = imagen(s, 'Captura SIPOC MP21', 'mp21v2_sipoc', MX, 2.62, w=w)
    f = mapper(box, crop)
    marca(s, f, (62, 118, 1216, 156), 1, place='tl')
    marca(s, f, (64, 364, 442, 427), 2, place='bl')
    marca(s, f, (852, 212, 1194, 259), 3, place='tr')
    marca(s, f, (852, 381, 1194, 427), 4, place='br')
    marca(s, f, (56, 596, 1232, 692), 5, place='bl')
    tb(s, 'Alcance', MX, 2.62 + box[3] + 0.10, w, 0.50,
       [para([r('Alcance: ', sz=950, b=True, color=INK),
              r('cinco plantas del Negocio de Celulosa (Arauco, Constitución, Licancel, Nueva '
                'Aldea y Valdivia), extensible a Forestal.', sz=950)], line=104)])
    xr = MX + w + 0.28
    wr = MX + CW - xr
    head(s, xr, 2.62, wr, 'Qué muestra', sz=1100)
    leyenda(s, xr, 2.98, wr, [
        (1, [r('Objetivo: ', sz=950, b=True, color=INK),
             r('servicios de soporte no operacional, que no afectan la producción.', sz=950)], 3),
        (2, [r('Entrada desde MP19: ', sz=950, b=True, color=INK),
             r('el contrato adjudicado.', sz=950)], 2),
        (3, [r('Salida hacia MP19: ', sz=950, b=True, color=INK),
             r('bases técnicas y presupuesto.', sz=950)], 2),
        (4, [r('Estándares actualizados: ', sz=950, b=True, color=INK),
             r('el ciclo se retroalimenta.', sz=950)], 2),
        (5, [r('Cinco macroprocesos', sz=950, b=True, color=INK),
             r(' de soporte y transversales.', sz=950)], 2),
    ], sz=950, gap=0.05, line_h=0.18)
    fuente(s, 'Figura 7. SIPOC del macroproceso Gestión de Servicios Generales (manual de procesos '
              'MP21). En naranjo, contenido propuesto.')
    return s


# ------------------------------------------------------------------ OE1 8 y 9: ejemplos
def _ejemplo(num, k, label, titulo, fl, ta, crop_t, marks_f, marks_t, ley, msg, fig):
    s = Slide(L_BODY, label=label)
    header(s, oe1(k), num)
    bajada(s, titulo)
    cropf = (60, 115, 1258, 660)
    wf = 4.70
    boxf = imagen(s, 'Captura flujograma', fl, MX, 2.60, w=wf)
    f = mapper(boxf, cropf)
    for n, ref, place in marks_f:
        marca(s, f, ref, n, place=place, size=0.22)
    xt = MX + wf + 0.22
    wt = MX + CW - xt
    boxt = imagen(s, 'Captura tareas', ta, xt, 2.60, w=wt)
    gm = mapper(boxt, crop_t)
    for n, ref, place in marks_t:
        marca(s, gm, ref, n, place=place, size=0.22)
    leyenda(s, xt, 2.60 + boxt[3] + 0.10, wt, ley, sz=920, gap=0.02, line_h=0.175)
    yc = 2.60 + boxf[3] + 0.22
    callout(s, MX, yc, wf, 0.95, msg, sz=980)
    fuente(s, fig)
    return s


def s_ej_c1(num):
    ley = [
        (1, [r('Cumplimiento laboral ', sz=920, b=True, color=INK),
             r('del contratista.', sz=920)], 1),
        (2, [r('Desarrollo de proveedores: ', sz=920, b=True, color=INK),
             r('KPI y mejora.', sz=920)], 1),
        (3, [r('Sanción ', sz=920, b=True, color=INK),
             r('cuando ocurre el incumplimiento.', sz=920)], 1),
        (4, [r('Naranjo: ', sz=920, b=True, color=INK), r('tarea propuesta.', sz=920)], 1),
        (5, [r('Normal: ', sz=920, b=True, color=INK), r('tarea levantada.', sz=920)], 1),
    ]
    msg = [r('Lo distintivo de MP21: ', sz=980, b=True, color=INK),
           r('el cumplimiento laboral del contratista se controla durante la ejecución (C.1.3) y '
             'la sanción se aplica cuando ocurre el incumplimiento (C.1.5).', sz=980)]
    return _ejemplo(
        num, 8, 'OE1 ejemplo C.1',
        'Ejemplo del manual MP21: subproceso C.1 Aseguramiento del servicio.',
        'mp21v2_c1_flujograma', 'mp21v2_c1_tareas', (30, 135, 1220, 650),
        [(1, (645, 238, 747, 321), 'tl'), (2, (810, 238, 914, 321), 'tr'),
         (3, (644, 435, 914, 527), 'bl')],
        [(4, (372, 178, 842, 274), 'tr'), (5, (372, 470, 842, 497), 'tr')],
        ley, msg,
        'Figura 8. Flujograma y tabla de actividades y tareas del subproceso C.1 (manual MP21).')


def s_ej_c2(num):
    ley = [
        (1, [r('Auditoría QA/QC: ', sz=920, b=True, color=INK),
             r('control del servicio.', sz=920)], 1),
        (2, [r('Gris: ', sz=920, b=True, color=INK),
             r('quién audita, por confirmar.', sz=920)], 1),
        (3, [r('Término como decisión: ', sz=920, b=True, color=INK),
             r('renovar o terminar.', sz=920)], 1),
        (4, [r('Tres responsables pendientes', sz=920, b=True, color=INK),
             r(' (C.2.1).', sz=920)], 1),
        (5, [r('Decide la Subgerencia; ', sz=920, b=True, color=INK),
             r('ejecuta Compras.', sz=920)], 1),
    ]
    msg = [r('La continuidad del contrato se decide dentro del ciclo: ', sz=980, b=True,
             color=INK),
           r('C.2.3 evalúa al contratista y define renovar, modificar o terminar; la decisión '
             'se ejecuta a través de Gestión de compras y contratos.', sz=980)]
    return _ejemplo(
        num, 9, 'OE1 ejemplo C.2',
        'Ejemplo del manual MP21: subproceso C.2 Control y mejora del servicio.',
        'mp21v2_c2_flujograma', 'mp21v2_c2_tareas', (30, 135, 1220, 500),
        [(1, (269, 331, 373, 415), 'tl'), (2, (1004, 164, 1256, 206), 'tr'),
         (3, (546, 331, 936, 506), 'br')],
        [(4, (843, 177, 1214, 274), 'tr'), (5, (372, 424, 1214, 459), 'bl')],
        ley, msg,
        'Figura 9. Flujograma y tabla de actividades y tareas del subproceso C.2 (manual MP21).')


# ------------------------------------------------------------------ OE1 10: estado
def s_estado(num):
    s = Slide(L_BODY, label='OE1 estado del contenido')
    header(s, oe1(10), num)
    bajada(s, 'Estado del contenido de los manuales: tareas levantadas en terreno y propuestas por validar.')
    # barras de MP21
    wl = 4.30
    head(s, MX, 2.62, wl, 'Tareas de MP21 por proceso', sz=1100)
    names = {'A': 'A Gestión estratégica', 'B': 'B Planificación',
             'C': 'C Administración', 'D': 'D Evaluación'}
    x_bar = MX + 1.72
    w_max = 2.10
    vmax = max(d['tasks'] for d in S21['out'].values())
    y = 3.00
    rh = 0.40
    for p, d in S21['out'].items():
        tb(s, 'Barra rotulo ' + p, MX, y, 1.68, rh - 0.06, [para([r(names[p], sz=950)])],
           anchor='ctr')
        wlv = w_max * d['task_lev'] / vmax
        wp = w_max * d['task_prop'] / vmax
        if wlv > 0:
            rect(s, 'Barra lev ' + p, x_bar, y + 0.06, wlv, rh - 0.18, fill=INK)
        if wp > 0:
            rect(s, 'Barra prop ' + p, x_bar + wlv, y + 0.06, wp, rh - 0.18, fill=BLUE_P)
        tb(s, 'Barra valor ' + p, x_bar + wlv + wp + 0.05, y, 0.45, rh - 0.06,
           [para([r(str(d['tasks']), sz=950, b=True, color=INK)])], anchor='ctr')
        y += rh
    rect(s, 'Ley lev', x_bar, y + 0.10, 0.20, 0.14, fill=INK)
    tb(s, 'Ley lev t', x_bar + 0.25, y + 0.03, 0.95, 0.28, [para([r('Levantadas', sz=900)])],
       anchor='ctr')
    rect(s, 'Ley prop', x_bar + 1.15, y + 0.10, 0.20, 0.14, fill=BLUE_P)
    tb(s, 'Ley prop t', x_bar + 1.40, y + 0.03, 1.00, 0.28, [para([r('Propuestas', sz=900)])],
       anchor='ctr')
    # tabla comparativa
    xr = MX + wl + 0.30
    wr = MX + CW - xr
    head(s, xr, 2.62, wr, 'MP8 frente a MP21', sz=1100)

    def v(n, tot):
        return '%d (%s %%)' % (n, pct(n, tot))
    rows = [
        [cell('Tareas', b=True, color=INK), cell(str(T8['task']), algn='ctr'),
         cell(str(T21['task']), algn='ctr')],
        [cell('Tareas levantadas'), cell(v(T8['task_lev'], T8['task']), algn='ctr'),
         cell(v(T21['task_lev'], T21['task']), algn='ctr', b=True, color=INK)],
        [cell('Tareas propuestas'), cell(v(T8['task_prop'], T8['task']), algn='ctr'),
         cell(v(T21['task_prop'], T21['task']), algn='ctr', b=True, color=INK)],
        [cell('Actividades levantadas'), cell('%d de %d' % (ACT8[0], T8['act']), algn='ctr'),
         cell('%d de %d' % (ACT21[0], T21['act']), algn='ctr')],
        [cell('Con responsable'), cell(v(T8['resp_ok'], T8['task']), algn='ctr', b=True,
                                       color=INK),
         cell(v(T21['resp_ok'], T21['task']), algn='ctr')],
    ]
    t = apa([('Indicador', 'l'), ('MP8', 'ctr'), ('MP21', 'ctr')], rows, sz=950)
    s.add(table(s.nid(), 'Tabla comparativa', xr, 2.98, [wr * 0.40, wr * 0.30, wr * 0.30], t,
                [0.28] + [0.30] * 5))
    # mensajes
    ym = 5.02
    wm = (CW - 0.20) / 2
    callout(s, MX, ym, wm, 0.78,
            [r('MP8: ', sz=980, b=True, color=INK),
             r('el contenido está levantado en terreno; falta identificar %d responsables.'
               % T8['resp_pend'], sz=980)])
    callout(s, MX + wm + 0.20, ym, wm, 0.78,
            [r('MP21: ', sz=980, b=True, color=INK),
             r('estructura y responsables definidos; falta validar %d tareas propuestas con '
               'el dueño.' % T21['task_prop'], sz=980)])
    tb(s, 'CAP', MX, 5.92, CW, 0.56,
       [para([r('CAP: ', sz=980, b=True, color=INK),
              r('los 4 procesos de MP21 están documentados en los cinco niveles; el indicador '
                'se declara al cierre de la validación formal (A1.6), como establece la sección '
                '2.3.1.', sz=980)], line=104)])
    fuente(s, 'Fuente: elaboración propia a partir de los manuales MP8 y MP21. Levantado: '
              'texto normal; propuesto: texto naranjo.')
    return s


# ------------------------------------------------------------------ OE2 1
def s_oe2a(num):
    s = Slide(L_BODY, label='OE2 1')
    header(s, 'Desarrollo y resultados del OE2 (1 de 2)', num)
    tot = T8['task'] + T21['task']
    tb(s, 'Estado', MX, BODY_T, CW, 0.56,
       [para([r('Actividades realizadas. ', sz=1150, b=True, color=INK),
              r('OE2 se inicia en la semana 12 según el plan; a la fecha no se han realizado '
                'talleres. El resultado de este periodo es el universo de ', sz=1150),
              r('%d tareas' % tot, sz=1150, b=True, color=INK),
              r(' que se evaluará, definido por OE1.', sz=1150)], line=106)])
    head(s, MX, 2.88, 5.60, 'Tareas por proceso (universo de OE2)', sz=1100)
    n8 = {'A': 'A Necesidad y dotación', 'B': 'B Valoración de cargo',
          'C': 'C Atracción y selección', 'D': 'D Desarrollo organizacional',
          'E': 'E Compensaciones'}
    n21 = {'A': 'A Gestión estratégica', 'B': 'B Planificación',
           'C': 'C Administración de servicios', 'D': 'D Evaluación del servicio'}
    x_lab, w_lab = MX + 0.55, 2.05
    x_bar = MX + 2.68
    w_max = 2.40
    vmax = max(d['tasks'] for d in S8['out'].values())
    y = 3.20
    rh = 0.30
    for grp, data, names in (('MP8', S8['out'], n8), ('MP21', S21['out'], n21)):
        y0 = y
        for p, d in data.items():
            tb(s, 'Rot %s %s' % (grp, p), x_lab, y, w_lab, rh, [para([r(names[p], sz=900)])],
               anchor='ctr')
            wl = w_max * d['task_lev'] / vmax
            wp = w_max * d['task_prop'] / vmax
            if wl > 0:
                rect(s, 'Lev %s %s' % (grp, p), x_bar, y + 0.06, wl, rh - 0.12, fill=INK)
            if wp > 0:
                rect(s, 'Prop %s %s' % (grp, p), x_bar + wl, y + 0.06, wp, rh - 0.12,
                     fill=BLUE_P)
            tb(s, 'Val %s %s' % (grp, p), x_bar + wl + wp + 0.05, y, 0.45, rh,
               [para([r(str(d['tasks']), sz=900, b=True, color=INK)])], anchor='ctr')
            y += rh
        tb(s, 'Grupo ' + grp, MX, y0, 0.52, y - y0,
           [para([r(grp, sz=1000, b=True, color=BLUE_LT)])], anchor='ctr')
        rect(s, 'Llave ' + grp, MX + 0.50, y0 + 0.04, 0.02, y - y0 - 0.08, fill=RULE)
        y += 0.08
    rect(s, 'Ley lev', x_bar, y + 0.06, 0.20, 0.14, fill=INK)
    tb(s, 'Ley lev t', x_bar + 0.25, y, 0.95, 0.26, [para([r('Levantadas', sz=900)])],
       anchor='ctr')
    rect(s, 'Ley prop', x_bar + 1.15, y + 0.06, 0.20, 0.14, fill=BLUE_P)
    tb(s, 'Ley prop t', x_bar + 1.40, y, 1.60, 0.26,
       [para([r('Propuestas por validar', sz=900)])], anchor='ctr')
    xr = MX + 5.95
    wr = MX + CW - xr
    s.add(shape(s.nid(), 'Regla vertical', xr - 0.22, 2.94, 0.014, 3.50, fill=RULE, line=None))
    head(s, xr, 2.88, wr, 'Condiciones para iniciar', sz=1100)
    conds = [
        ('1', 'Validar %d tareas propuestas (%d de MP8 y %d de MP21).'
         % (T8['task_prop'] + T21['task_prop'], T8['task_prop'], T21['task_prop'])),
        ('2', 'Identificar %d responsables pendientes (%d y %d).'
         % (T8['resp_pend'] + T21['resp_pend'], T8['resp_pend'], T21['resp_pend'])),
        ('3', 'Incorporar F15 a F18 al manual MP8.'),
        ('4', 'Formalizar con su dueño la estructura de MP21.'),
    ]
    y = 3.24
    for n, t in conds:
        tb(s, 'Cond num ' + n, xr, y, 0.28, 0.44, [para([r(n, sz=1200, b=True, color=BLUE_LT)])])
        tb(s, 'Cond ' + n, xr + 0.30, y, wr - 0.30, 0.66, [para([r(t, sz=1000)], line=105)])
        y += 0.74
    fuente(s, 'Fuente: elaboración propia a partir de los manuales MP8 y MP21. Una tarea no '
              'descrita no puede evaluarse: estas condiciones fijan la cobertura del TCER.')
    return s


# ------------------------------------------------------------------ OE2 2
def s_oe2b(num):
    s = Slide(L_BODY, label='OE2 2')
    header(s, 'Desarrollo y resultados del OE2 (2 de 2)', num)
    tb(s, 'Bajada', MX, BODY_T, CW, 0.56,
       [para([r('La matriz de riesgos se construye sobre la tabla de actividades y tareas del '
                'manual: cada fila es una tarea con su código y responsable, a la que se agregan '
                'los siete atributos de COSO II.', sz=1150)], line=106)])
    y0 = 2.95
    wcol = 3.90
    head(s, MX, y0, wcol, 'Desde el manual de procesos (OE1)', sz=1100)
    regla(s, y0 + 0.34, x=MX, w=wcol, color=INK, h=0.018)
    y = y0 + 0.42
    for t in ['Código MGO', 'Subproceso y actividad', 'Tarea', 'Responsable']:
        tb(s, 'Origen ' + t, MX, y, wcol, 0.27, [para([r(t, sz=1050)])], anchor='ctr')
        y += 0.29
    flecha(s, 'rightArrow', MX + wcol + 0.18, y0 + 0.95, 0.55, 0.36)
    x2 = MX + wcol + 0.92
    w2 = MX + CW - x2
    head(s, x2, y0, w2, 'Atributos COSO II por tarea (OE2)', sz=1100)
    regla(s, y0 + 0.34, x=x2, w=w2, color=INK, h=0.018)
    attrs = ['Evento de riesgo', 'Categoría', 'Probabilidad', 'Impacto', 'Riesgo inherente',
             'Controles existentes', 'Riesgo residual']
    for k, t in enumerate(attrs):
        cx = x2 + (0 if k < 4 else w2 / 2)
        cy = y0 + 0.42 + 0.29 * (k if k < 4 else k - 4)
        tb(s, 'Atributo ' + t, cx, cy, w2 / 2, 0.27, [para([r(t, sz=1050)])], anchor='ctr')
    ye = 4.66
    head(s, MX, ye, CW, 'Ejemplo de filas con datos de los manuales MP8 y MP21', sz=1100)
    col = [1.05, 3.45, 1.80, 2.80]
    por = 'Se determina en los talleres A2.1 y A2.2'
    rows = [
        [cell('MP8 C.1.1', b=True, color=INK),
         cell('Verificar disponibilidad de la vacante con el área de Personas'),
         cell('Área Solicitante'), cell(por, i=True, color=TX_SOFT)],
        [cell('MP21 C.1.3', b=True, color=INK),
         cell('Revisar certificados de cumplimiento laboral y previsional del contratista'),
         cell('Jefe de Servicios Generales zonal'), cell(por, i=True, color=TX_SOFT)],
    ]
    t = apa([('Código', 'l'), ('Tarea', 'l'), ('Responsable', 'l'), ('Atributos COSO II', 'l')],
            rows, sz=1000)
    s.add(table(s.nid(), 'Tabla ejemplo', MX, ye + 0.36, col, t, [0.30, 0.50, 0.50]))
    fuente(s, 'Fuente: elaboración propia. Atributos según la sección 2.3.2 del informe; tareas y '
              'responsables tomados de los manuales MP8 y MP21.')
    return s


# ------------------------------------------------------------------ Gantt
GANTT = [
    ('A1.1–A1.2 Diagnóstico y entrevistas', (1, 4), (1, 4, 'c'), 'Completada'),
    ('A1.3.1 Flujo único MP8', (3, 5), (3, 5, 'c'), 'Completada'),
    ('A1.3.2 Segmentación MP8', (6, 7), (6, 10, 'e'), '14 de 18 fases'),
    ('A1.4.1 Estructura MP21 con el dueño*', (6, 9), (5, 10, 'c'), 'Completada'),
    ('A1.4.2 Segmentación MP21', (10, 11), (10, 10, 'c'), 'Completada'),
    ('A1.5–A1.6 Manuales y validación', (11, 13), (10, 10, 'e'), 'En curso'),
    ('A2.1–A2.2 Identificación y valoración', (12, 16), None, 'Planificada'),
    ('A2.3–A2.4 Matriz y validación', (16, 18), None, 'Planificada'),
    ('A3.1–A3.2 Planes y gobernanza', (17, 19), None, 'Planificada'),
    ('A3.3–A3.4 Integración y entrega', (19, 20), None, 'Planificada'),
    ('A-G Documentación del TFG', (1, 20), (1, 10, 'e'), 'En curso'),
]


def s_gantt(num):
    s = Slide(L_BODY, label='Gantt')
    header(s, 'Carta Gantt · Nivel de cumplimiento', num)
    sem = g.SEMANA
    tb(s, 'Bajada', MX, BODY_T - 0.06, CW, 0.34,
       [para([r('Horizonte de 20 semanas, agosto 2026 a enero 2027. Esta presentación '
                'corresponde a la ', sz=1100), r('semana %d' % sem, sz=1100, b=True, color=INK),
              r('.', sz=1100)])], anchor='ctr')
    wa, wk, we = 2.62, 0.262, 1.24
    col = [wa] + [wk] * 20 + [we]
    hdr = [cell('Actividad', sz=950, b=True, color=WHITE, fill=INK)]
    for k in range(1, 21):
        hdr.append(cell(str(k), sz=850, b=True, color=WHITE, fill=INK, algn='ctr', marL=0,
                        marR=0))
    hdr.append(cell('Estado', sz=950, b=True, color=WHITE, fill=INK))
    rows = [hdr]
    for name, plan, real, estado in GANTT:
        row = [cell(name, sz=900, color=TX, lnB=(RULE, 0.5))]
        for k in range(1, 21):
            fill = None
            if plan[0] <= k <= plan[1]:
                fill = BLUE_VP
            if real and real[0] <= k <= real[1]:
                fill = INK if real[2] == 'c' else BLUE_LT
            row.append(cell('', sz=600, fill=fill, lnB=(RULE, 0.5), marL=0, marR=0))
        bold = estado.startswith('Completada')
        row.append(cell(estado, sz=900, b=bold, color=INK if bold else TX, lnB=(RULE, 0.5)))
        rows.append(row)
    hit = [cell('Hitos A-P1 · A-P2 · A-PF', sz=900, b=True, color=INK, lnB=(INK, 1.0))]
    for k in range(1, 21):
        fill = INK if k in (5, sem) else (BLUE_VP if k == 20 else None)
        hit.append(cell('', sz=600, fill=fill, lnB=(INK, 1.0), marL=0, marR=0))
    hit.append(cell('A-P2 en S%d' % sem, sz=900, color=TX, lnB=(INK, 1.0)))
    rows.append(hit)
    rh, ty = 0.25, 2.52
    s.add(table(s.nid(), 'Tabla Gantt', MX, ty, col, rows, [0.28] + [rh] * (len(rows) - 1)))
    xl = MX + wa + wk * sem
    th = 0.28 + rh * (len(rows) - 1)
    rect(s, 'Semana actual', xl - 0.012, ty, 0.024, th + 0.10, fill='C0504D')
    tb(s, 'Semana actual rotulo', xl - 0.60, ty + th + 0.10, 1.20, 0.22,
       [para([r('Semana %d' % sem, sz=850, b=True, color='C0504D')], algn='ctr')],
       anchor='ctr')
    ly = ty + th + 0.38
    lx = MX
    for fill, text, w in ((INK, 'Completado', 1.05), (BLUE_LT, 'Ejecutado o en curso', 1.70),
                          (BLUE_VP, 'Planificado', 1.05)):
        rect(s, 'Ley ' + text, lx, ly + 0.05, 0.22, 0.14, fill=fill)
        tb(s, 'Ley t ' + text, lx + 0.27, ly, w, 0.24, [para([r(text, sz=900)])], anchor='ctr')
        lx += w + 0.40
    tb(s, 'Cumplimiento', MX, ly + 0.24, 8.40, 0.56,
       [para([r('OE1: ', sz=1000, b=True, color=INK),
              r('de 8 actividades, 5 completadas, 2 en curso y 1 no iniciada (validación formal, '
                'A1.6). Pendiente principal: fases F15 a F18 de MP8.', sz=1000)], line=105),
        para([r('* Ajuste metodológico: en MP21 la estructura se acordó con el dueño del proceso '
                'en lugar de derivarse de un flujo único (lámina 9).', sz=850, i=True,
                color=TX_SOFT)], line=105, space_before=2)])
    return s


# ------------------------------------------------------------------ armado
def make(out):
    seq = [s_problema1, g.s_problema2, g.s_objetivos, s_oe1_niveles, g.s_oe1b, g.s_oe1c,
           s_ajuste, s_apertura21, s_ciclo, s_sipoc21, s_ej_c1, s_ej_c2, s_estado,
           s_oe2a, s_oe2b, s_gantt]
    TITLES.update({
        6: (TAG1, 'Resultado: MP8 y MP21 modelados en cinco niveles'),
        7: (TAG1, 'Resultado: MP8 documentado hasta la tarea'),
        8: (TAG1, 'Resultado: cómo se articula el macroproceso MP8'),
        9: (TAG1, 'Resultado: estructura de MP21 acordada con su dueño'),
        10: (TAG1, 'Resultado: arquitectura del macroproceso MP21'),
        11: (TAG1, 'Resultado: un ciclo de gestión único en MP21'),
        12: (TAG1, 'Resultado: cómo se articula el macroproceso MP21'),
        13: (TAG1, 'Resultado: MP21 controla al contratista en ejecución'),
        14: (TAG1, 'Resultado: control y continuidad del servicio'),
        15: (TAG1, 'Resultado: MP8 levantado; MP21 por validar'),
        16: (TAG2, 'Resultado: universo de 248 tareas para el OE2'),
        17: (TAG2, 'Resultado: estructura de la matriz de riesgos'),
    })
    slides = [g.s_portada(), s_contenidos(2)]
    for k, fn in enumerate(seq, 3):
        slides.append(fn(k))
    slides.append(g.s_cierre())
    build(g.BASE, out, slides, title='Avance 2 TFG - ' + AUTOR, author=AUTOR, media=g.MEDIA)
    return slides


if __name__ == '__main__':
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, 'out', 'Avance_2_v3.pptx')
    sl = make(out)
    print('OK', out, len(sl), 'láminas')
