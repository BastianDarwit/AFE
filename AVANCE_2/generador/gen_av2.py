# -*- coding: utf-8 -*-
"""Avance 2 del TFG · genera la presentación sobre la plantilla oficial.

Estructura de la plantilla «Avance 2 AFE 2017»:
  portada · título · problema (2) · marco teórico (2) · OE1 (1-2) · OE2 (2)
  · carta Gantt con nivel de cumplimiento · cierre.
Los datos del avance provienen de los manuales MP8 v1 y MP21 v1 (stats_*.json).
"""
import json
import os
import struct
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

from pptxlib import (Slide, shape, textbox, table, para, run, picture, esc,  # noqa: E402
                     REL_IMAGE, build, WHITE)

BASE = os.path.join(ROOT, 'plantilla_av2.pptx')
CAP = os.path.join(ROOT, 'capturas')

L_COVER = 'slideLayout1.xml'
L_BODY = 'slideLayout10.xml'
L_PHOTO = 'slideLayout21.xml'
L_CLOSE = 'slideLayout7.xml'

# --- paleta institucional (azules UFRO y neutros) ---
INK = '17375E'
BLUE = '1F497D'
BLUE_M = '2E5F92'
BLUE_LT = '4F81BD'
BLUE_S = '7FA6D0'
BLUE_P = 'A9C4E2'
BLUE_VP = 'D2E0EF'
BLUE_BG = 'EDF2F8'
RULE = 'B8C6D9'
TX = '3A3A3A'
TX_SOFT = '6B6B6B'
TITLE_GRAY = '595959'

TFG = ('Propuesta de un sistema de gestión de riesgos para los macroprocesos de '
       'Gestión de Personas (MP8) y Servicios Generales (MP21) de Celulosa Arauco y '
       'Constitución S.A.')
AUTOR = 'Bastián Darwit Vera'
GUIA = 'Carlos Muñoz P.'
COGUIA = 'Manuel Echiburú Fuenzalida'
FECHA = '14/10/2026'
SEMANA = 10

MX = 0.45          # margen izquierdo del cuerpo
CW = 9.10          # ancho útil
BODY_T = 2.22      # inicio del cuerpo
BODY_B = 6.50      # fin del cuerpo (sobre el número de lámina)

S8 = json.load(open(os.path.join(ROOT, 'stats_mp8.json')))
S21 = json.load(open(os.path.join(ROOT, 'stats_mp21.json')))
T8 = S8['tot']
T21 = S21['tot']

MEDIA = {}


def pct(a, b):
    return ('%.1f' % (100.0 * a / b)).replace('.', ',')


def png_size(path):
    d = open(path, 'rb').read(24)
    return struct.unpack('>II', d[16:24])


# ------------------------------------------------------------------ helpers
def r(text, sz=1200, b=False, i=False, color=TX):
    return run(text, sz=sz, b=b, i=i, color=color)


def tb(s, name, x, y, w, h, paras, anchor='t', ins=(0.04, 0.02, 0.04, 0.02)):
    return s.add(textbox(s.nid(), name, x, y, w, h, paras, anchor=anchor, ins=ins,
                         autofit=False))


def rect(s, name, x, y, w, h, fill=None, line=None, line_w=1.0, geom='rect',
         paras=None, anchor='ctr', ins=(0.04, 0.02, 0.04, 0.02), dash=None):
    return s.add(shape(s.nid(), name, x, y, w, h, geom=geom, fill=fill, line=line,
                       line_w=line_w, paras=paras, anchor=anchor, ins=ins,
                       autofit=False, dash=dash))


def regla(s, y, x=MX, w=CW, color=RULE, h=0.014):
    rect(s, 'Regla', x, y, w, h, fill=color)


def header(s, titulo, num):
    """Elementos comunes de las láminas de cuerpo, como en el Avance 1."""
    tb(s, 'Titulo TFG', 1.64, 0.18, 6.96, 0.98,
       [para([r(TFG, sz=1100, b=True, color=INK)], algn='r', line=112)],
       anchor='ctr')
    tb(s, 'Estudiante', 6.28, 1.13, 2.21, 0.24,
       [para([r('| ', sz=900, color='7F7F7F'), r('Estudiante: ', sz=900, b=True, color=INK),
              r(AUTOR, sz=900, color=INK), r(' |', sz=900, color='7F7F7F')], algn='ctr')],
       anchor='ctr', ins=(0.02, 0.01, 0.02, 0.01))
    tb(s, 'Titulo lamina', MX, 1.46, CW, 0.56,
       [para([r(titulo, sz=2200, b=True, color=TITLE_GRAY)])], anchor='ctr')
    tb(s, 'Numero', 8.92, 6.62, 0.62, 0.44,
       [para([r(str(num), sz=2000, b=True, color=INK)], algn='r')], anchor='ctr')


def fuente(s, text, y=6.62, w=8.30):
    tb(s, 'Fuente', MX, y, w, 0.28,
       [para([r(text, sz=850, i=True, color=TX_SOFT)])], anchor='ctr')


def head(s, x, y, w, text, sz=1300, color=BLUE):
    tb(s, 'Encabezado', x, y, w, 0.30, [para([r(text, sz=sz, b=True, color=color)])],
       anchor='ctr')


def imagen(s, name, fname, x, y, w=None, h=None, line=True):
    """Inserta una captura y devuelve (x, y, w, h)."""
    path = os.path.join(CAP, fname + '.png')
    pw, ph = png_size(path)
    if w is None:
        w = h * pw / ph
    if h is None:
        h = w * ph / pw
    media_name = 'av2_%s.png' % fname
    MEDIA[media_name] = open(path, 'rb').read()
    rid = 'rId%d' % (10 + len(s.rels))
    s.rels.append((rid, REL_IMAGE, '../media/' + media_name))
    s.add(picture(s.nid(), name, x, y, w, h, rid))
    if line:
        rect(s, 'Borde captura', x, y, w, h, fill=None, line=RULE, line_w=0.75)
    return x, y, w, h


def mapper(box, crop):
    """Convierte píxeles de referencia (1280 x 720) a pulgadas de la lámina."""
    x, y, w, h = box
    x0, y0, x1, y1 = crop

    def f(px, py):
        return x + (px - x0) / (x1 - x0) * w, y + (py - y0) / (y1 - y0) * h
    return f


def marca(s, f, ref, n, place='tl', size=0.24):
    """Señalización: rectángulo de contorno sobre la captura y número de referencia."""
    x0, y0, x1, y1 = ref
    ax, ay = f(x0, y0)
    bx, by = f(x1, y1)
    rect(s, 'Senal %d' % n, ax, ay, bx - ax, by - ay, fill=None, line=INK, line_w=2.0)
    if place == 'tl':
        cx, cy = ax - size * 0.55, ay - size * 0.55
    elif place == 'tr':
        cx, cy = bx - size * 0.45, ay - size * 0.55
    elif place == 'bl':
        cx, cy = ax - size * 0.55, by - size * 0.45
    else:
        cx, cy = bx - size * 0.45, by - size * 0.45
    rect(s, 'Marca %d' % n, cx, cy, size, size, fill=INK, line=WHITE, line_w=1.0,
         geom='ellipse',
         paras=[para([r(str(n), sz=1000, b=True, color=WHITE)], algn='ctr')],
         ins=(0, 0, 0, 0))


def leyenda(s, x, y, w, items, sz=1050, gap=0.06, line_h=0.205):
    """Leyenda numerada de señalizaciones. Devuelve la y final."""
    cur = y
    for n, text, lines in items:
        rect(s, 'Ley marca %d' % n, x, cur + 0.02, 0.22, 0.22, fill=INK, geom='ellipse',
             paras=[para([r(str(n), sz=900, b=True, color=WHITE)], algn='ctr')],
             ins=(0, 0, 0, 0))
        h = lines * line_h + 0.04
        tb(s, 'Ley texto %d' % n, x + 0.30, cur, w - 0.30, h,
           [para(text if isinstance(text, list) else [r(text, sz=sz)], line=105)])
        cur += h + gap
    return cur


def cell(t='', sz=1050, b=False, color=TX, fill=None, algn='l', anchor='ctr', i=False,
         lnT=None, lnB=None, runs=None, marL=0.05, marR=0.05):
    d = dict(t=t, sz=sz, b=b, color=color, fill=fill, algn=algn, anchor=anchor, i=i,
             lnT=lnT, lnB=lnB, marL=marL, marR=marR, marT=0.02, marB=0.02)
    if runs:
        d['runs'] = runs
    return d


def apa(header_cells, rows, sz=1050, hsz=None):
    """Tabla con reglas horizontales: superior, bajo encabezado e inferior."""
    hsz = hsz or sz
    out = [[cell(h, sz=hsz, b=True, color=INK, algn=a, lnT=(INK, 1.25), lnB=(INK, 0.75))
            for h, a in header_cells]]
    for k, row in enumerate(rows):
        last = k == len(rows) - 1
        cur = []
        for c in row:
            c = dict(c)
            c['lnB'] = (INK, 1.25) if last else (RULE, 0.5)
            cur.append(c)
        out.append(cur)
    return out


# ------------------------------------------------------------------ 1 portada
def s_portada():
    s = Slide(L_COVER, label='Portada')
    tb(s, 'Antetitulo', 0.80, 3.18, 8.40, 0.36,
       [para([r('Avance 2 · Trabajo Final de Grado', sz=1500, b=True, color=TITLE_GRAY)],
             algn='ctr')], anchor='ctr')
    tb(s, 'Titulo', 0.80, 3.58, 8.40, 0.98,
       [para([r(TFG, sz=1600, b=True, color=INK)], algn='ctr', line=108)], anchor='ctr')
    tb(s, 'Programa', 0.80, 4.62, 8.40, 0.46,
       [para([r('Magíster en Sistemas de Gestión Integral de la Calidad', sz=1150, b=True,
                color=TX)], algn='ctr'),
        para([r('Facultad de Ingeniería y Ciencias · Universidad de La Frontera', sz=1000,
                color=TX_SOFT)], algn='ctr')], anchor='ctr')
    tb(s, 'Autoria', 2.30, 5.14, 5.40, 0.92,
       [para([r('| ', sz=1150, color='7F7F7F'), r('Estudiante: ', sz=1150, b=True, color=INK),
              r(AUTOR, sz=1150, color=INK), r(' |', sz=1150, color='7F7F7F')], algn='ctr'),
        para([r('Profesor guía: ', sz=1100, b=True, color=INK), r(GUIA, sz=1100, color=INK)],
             algn='ctr', space_before=2),
        para([r('Profesor co-guía: ', sz=1100, b=True, color=INK), r(COGUIA, sz=1100, color=INK)],
             algn='ctr', space_before=2)], anchor='ctr')
    tb(s, 'Fecha', 3.50, 6.38, 3.00, 0.30,
       [para([r('Fecha: ', sz=1100, b=True, color=INK), r(FECHA, sz=1100, color=INK)],
             algn='ctr')], anchor='ctr')
    return s


# ------------------------------------------------------------------ 2 título
def s_titulo():
    s = Slide(L_PHOTO, label='Título del trabajo')
    # Fotografía institucional de la plantilla (image4.jpeg)
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
              r('el nivel corporativo que diseña ambos macroprocesos para el conjunto de las '
                'plantas.', sz=1050)], line=105)])
    fuente(s, 'Fuente: elaboración propia sobre el mapa de macroprocesos de los manuales MP8 y '
              'MP21 (v1) y la sección 1.5 del informe.')
    return s


# ------------------------------------------------------------------ 4 problema 2
def s_problema2(num):
    s = Slide(L_BODY, label='Problema 2')
    header(s, 'Problema (2 de 2) · Descripción y relevancia', num)
    head(s, MX, BODY_T, CW, 'Descripción del problema')
    tb(s, 'Problema', MX, BODY_T + 0.32, CW, 0.70,
       [para([r('MP8 y MP21 operan sin un sistema de gestión de riesgos estructurado y sin '
                'una arquitectura de procesos formalizada bajo el MGO, lo que limita la '
                'trazabilidad de responsabilidades y la gestión de sus riesgos operacionales.',
                sz=1200)], line=108)])
    items = [('1', 'Sin arquitectura formalizada', 'en los niveles del MGO, que sirva de base '
              'para analizar los riesgos.'),
             ('2', 'Sin evaluación estructurada', 'de los riesgos operacionales de las tareas.'),
             ('3', 'Sin un sistema que integre', 'riesgos, planes de mitigación, responsables y '
              'monitoreo.')]
    y = 3.26
    for n, t1, t2 in items:
        tb(s, 'Carencia num %s' % n, MX, y, 0.36, 0.30,
           [para([r(n, sz=1300, b=True, color=BLUE_LT)])], anchor='ctr')
        tb(s, 'Carencia %s' % n, MX + 0.40, y, CW - 0.40, 0.30,
           [para([r(t1 + ' ', sz=1150, b=True, color=INK), r(t2, sz=1150)])], anchor='ctr')
        y += 0.36
    regla(s, 4.42)
    head(s, MX, 4.54, CW, 'Relevancia para la gestión de la calidad')
    xr = MX + 4.62
    tb(s, 'Relevancia 1', MX, 4.88, 4.40, 1.62,
       [para([r('La arquitectura de procesos es la condición previa para gestionar riesgos: '
                'ordena actividades dispersas y les asigna responsables trazables '
                '(Davenport, 1993; Teixeira et al., 2024).', sz=1100)], line=108),
        para([r('Un sistema sustentado en esa arquitectura puede replicarse en las cinco '
                'plantas y homogeneizar la gestión.', sz=1100)], line=108, space_before=6)])
    tb(s, 'Relevancia 2', xr, 4.88, MX + CW - xr, 1.62,
       [para([r('Evidencia externa, no proyectable al caso: ', sz=1100, b=True, color=INK),
              r('la adopción de ERM se asocia a una prima de valoración cercana al 20 % en '
                'aseguradoras (Hoyt y Liebenberg, 2011) y su madurez a una mayor Q de Tobin, '
                'del orden de 25 % (Farrell y Gallagher, 2015).', sz=1100)], line=108)])
    s.add(shape(s.nid(), 'Regla vertical', MX + 4.50, 4.92, 0.014, 1.50, fill=RULE, line=None))
    fuente(s, 'Fuente: secciones 1.2 y 1.3 del informe de avance.')
    return s


# ------------------------------------------------------------------ 5 marco 1
def s_marco1(num):
    s = Slide(L_BODY, label='Marco teórico 1')
    header(s, 'Marco teórico (1 de 2) · Qué dice la teoría', num)
    tb(s, 'Bajada', MX, BODY_T, CW, 0.44,
       [para([r('Gestión por procesos y gestión de riesgos operan como disciplinas '
                'complementarias: primero se ordena el proceso y luego se gestionan sus riesgos.',
                sz=1150)])], anchor='ctr')
    col = [2.10, 4.70, 2.30]
    rows = [
        [cell('Gestión por procesos', b=True, color=INK),
         cell('Ordena las actividades en flujos que cruzan las áreas y aporta trazabilidad y '
              'control.'),
         cell('Davenport (1993); Teixeira et al. (2024)', color=TX_SOFT, sz=950)],
        [cell('Marco de referencia', b=True, color=INK),
         cell('El APQC PCF clasifica procesos con un lenguaje común y permite contrastar lo '
              'que debería existir con lo documentado.'),
         cell('APQC (2023)', color=TX_SOFT, sz=950)],
        [cell('Madurez operativa', b=True, color=INK),
         cell('Describe la progresión desde procesos informales hasta procesos definidos y '
              'gestionados con datos.'),
         cell('SEI (2010); Szelągowski y Berniak-Woźny (2022)', color=TX_SOFT, sz=950)],
        [cell('ERM y COSO II', b=True, color=INK),
         cell('Ciclo de identificación de eventos, evaluación y respuesta al riesgo, aplicable '
              'en el nivel de tarea.'),
         cell('COSO (2004, 2017); Anton y Nucu (2020)', color=TX_SOFT, sz=950)],
        [cell('Sistema de gestión de riesgos', b=True, color=INK),
         cell('Integra la valoración con el tratamiento, la gobernanza y el seguimiento; se '
              'distingue de un catálogo.'),
         cell('COSO (2017)', color=TX_SOFT, sz=950)],
    ]
    t = apa([('Concepto', 'l'), ('Aporte al trabajo', 'l'), ('Fuente', 'l')], rows)
    s.add(table(s.nid(), 'Tabla marco', MX, 2.72, col, t, [0.36] + [0.66] * 5))
    fuente(s, 'Fuente: sección 1.4 del informe de avance. Referencias completas en normas APA en '
              'el informe.')
    return s


# ------------------------------------------------------------------ 6 marco 2
def s_marco2(num):
    s = Slide(L_BODY, label='Marco teórico 2')
    header(s, 'Marco teórico (2 de 2) · Qué se ha hecho en la práctica', num)
    col = [2.15, 2.55, 4.40]
    rows = [
        [cell('Hoyt y Liebenberg (2011)', b=True, color=INK),
         cell('Aseguradoras de EE. UU.'),
         cell('Prima de valoración cercana al 20 % (Q de Tobin) en adoptantes de ERM frente a '
              'no adoptantes.')],
        [cell('Farrell y Gallagher (2015)', b=True, color=INK),
         cell('Empresas de distintos sectores'),
         cell('Mayor valor en empresas con madurez ERM avanzada, del orden de 25 %.')],
        [cell('Yang et al. (2018)', b=True, color=INK),
         cell('Pequeñas y medianas empresas'),
         cell('Las prácticas de ERM se asocian a mejor desempeño mediante la ventaja '
              'competitiva.')],
        [cell('Tubis y Werbińska-W. (2021)', b=True, color=INK),
         cell('Procesos logísticos'),
         cell('Modelo de madurez de la gestión de riesgos por proceso.')],
        [cell('Soner y Karadag (2026)', b=True, color=INK),
         cell('Organizaciones de servicios'),
         cell('Índice de madurez ERM construido sobre COSO 2017.')],
        [cell('Jagrič et al. (2026)', b=True, color=INK),
         cell('Empresas medianas y grandes (n = 181)'),
         cell('Asociación positiva y significativa entre la calidad de la gestión de riesgos y '
              'el desempeño.')],
    ]
    t = apa([('Estudio', 'l'), ('Contexto', 'l'), ('Hallazgo', 'l')], rows)
    s.add(table(s.nid(), 'Tabla practica', MX, BODY_T, col, t, [0.36] + [0.50] * 6))
    regla(s, 5.72)
    tb(s, 'Brecha', MX, 5.80, CW, 0.70,
       [para([r('Brecha. ', sz=1150, b=True, color=INK),
              r('La evidencia se concentra en el nivel corporativo. En la revisión realizada no '
                'se identificaron estudios que documenten COSO II en el nivel de tarea de '
                'procesos de apoyo del sector forestal.', sz=1150)], line=108)])
    fuente(s, 'Fuente: sección 1.4.4 del informe de avance.')
    return s


# ------------------------------------------------------------------ 7 OE1 1
def s_oe1a(num):
    s = Slide(L_BODY, label='OE1 1')
    header(s, 'Desarrollo y resultados del OE1 (1 de 4)', num)
    tb(s, 'Actividades', MX, BODY_T, CW, 0.70,
       [para([r('Actividades realizadas. ', sz=1150, b=True, color=INK),
              r('A1.3.2 segmentación jerárquica de MP8 desde el flujo único del Avance 1, y '
                'A1.5 consolidación en manuales con el formato del MGO aprobado por la '
                'organización: MP8 v1 (58 láminas) y esqueleto MP21 v1 (37 láminas).',
                sz=1150)], line=106)])
    crop = (150, 135, 1128, 548)
    w = 5.55
    box = imagen(s, 'Captura apertura MP8', 'mp8_apertura', MX, 3.02, w=w)
    f = mapper(box, crop)
    marca(s, f, (158, 143, 346, 222), 1, place='tl')
    marca(s, f, (158, 224, 346, 286), 2, place='bl')
    marca(s, f, (546, 143, 734, 545), 3, place='tr')
    ly = 3.02 + box[3] + 0.10
    leyenda(s, MX, ly, w, [
        (1, [r('Proceso', sz=1000, b=True, color=INK), r(' (A a E), con su ficha y SIPOC.', sz=1000)], 1),
        (2, [r('Subproceso', sz=1000, b=True, color=INK),
             r(' (A.1), con flujograma, actividades y tareas.', sz=1000)], 1),
        (3, [r('Atracción y selección', sz=1000, b=True, color=INK),
             r(': cinco subprocesos (fases F3 a F10 del flujo).', sz=1000)], 1),
    ], sz=1000, gap=0.03, line_h=0.19)
    # tabla de niveles
    xr = MX + w + 0.30
    wr = MX + CW - xr
    head(s, xr, 3.02, wr, 'Elementos por nivel MGO', sz=1150)
    col = [1.27, 0.88, 0.88]
    rows = [
        [cell('Macroproceso', b=True, color=INK), cell('1', algn='ctr'), cell('1', algn='ctr')],
        [cell('Procesos', b=True, color=INK), cell(str(T8['proc']), algn='ctr'),
         cell(str(T21['proc']), algn='ctr')],
        [cell('Subprocesos', b=True, color=INK), cell(str(T8['sub']), algn='ctr'),
         cell(str(T21['sub']), algn='ctr')],
        [cell('Actividades', b=True, color=INK), cell(str(T8['act']), algn='ctr'),
         cell('Por levantar', algn='ctr', color=TX_SOFT, i=True, sz=950)],
        [cell('Tareas', b=True, color=INK), cell(str(T8['task']), algn='ctr', b=True, color=INK),
         cell('Por levantar', algn='ctr', color=TX_SOFT, i=True, sz=950)],
    ]
    t = apa([('Nivel', 'l'), ('MP8', 'ctr'), ('MP21', 'ctr')], rows)
    s.add(table(s.nid(), 'Tabla niveles', xr, 3.38, col, t, [0.32] + [0.30] * 5))
    tb(s, 'Cobertura', xr, 5.32, wr, 1.18,
       [para([r('MP8: ', sz=1000, b=True, color=INK),
              r('14 de las 18 fases del flujo único segmentadas; restan nómina, clima, '
                'desvinculación y datos (F15 a F18).', sz=1000)], line=105),
        para([r('MP21: ', sz=1000, b=True, color=INK),
              r('estructura de procesos y subprocesos; inicia el levantamiento del ciclo del '
                'contrato.', sz=1000)], line=105, space_before=4)])
    fuente(s, 'Figura 2. Apertura del macroproceso MP8 en el manual de procesos v1. Fuente: '
              'elaboración propia.')
    return s


# ------------------------------------------------------------------ 8 OE1 2
def s_oe1b(num):
    s = Slide(L_BODY, label='OE1 2')
    header(s, 'Desarrollo y resultados del OE1 (2 de 4)', num)
    tb(s, 'Bajada', MX, BODY_T - 0.04, CW, 0.32,
       [para([r('Ejemplo del manual MP8: subproceso C.1 Requisición y apertura de vacante.',
                sz=1100, color=INK, b=True)])], anchor='ctr')
    # captura del flujograma
    cropf = (60, 55, 1258, 660)
    wf = 4.70
    boxf = imagen(s, 'Captura flujograma C.1', 'mp8_c1_flujograma', MX, 2.60, w=wf)
    f = mapper(boxf, cropf)
    marca(s, f, (266, 331, 316, 354), 1, place='tl', size=0.22)
    marca(s, f, (596, 334, 700, 424), 2, place='tr', size=0.22)
    marca(s, f, (1000, 132, 1252, 632), 3, place='tl', size=0.22)
    # captura de la tabla de tareas
    cropt = (100, 135, 1220, 462)
    xt = MX + wf + 0.22
    wt = MX + CW - xt
    boxt = imagen(s, 'Captura tareas C.1', 'mp8_c1_tareas', xt, 2.60, w=wt)
    g = mapper(boxt, cropt)
    marca(s, g, (840, 176, 1090, 202), 4, place='tr', size=0.22)
    marca(s, g, (838, 278, 915, 300), 5, place='tr', size=0.22)
    ly = 2.60 + boxt[3] + 0.10
    leyenda(s, xt, ly, wt, [
        (1, [r('Código MGO', sz=950, b=True, color=INK), r(' de la actividad (C.1.1).', sz=950)], 1),
        (2, [r('Decisión', sz=950, b=True, color=INK), r(' del flujo: proceso masivo o individual.', sz=950)], 1),
        (3, [r('Soporte, herramientas, KPI y dueño', sz=950, b=True, color=INK),
             r(' del subproceso.', sz=950)], 1),
        (4, [r('Tarea con responsable', sz=950, b=True, color=INK), r(' identificado.', sz=950)], 1),
        (5, [r('Responsable pendiente', sz=950, b=True, color=INK), r(' de levantar.', sz=950)], 1),
    ], sz=950, gap=0.02, line_h=0.18)
    # tabla de estado del contenido
    yt = 2.60 + boxf[3] + 0.16
    head(s, MX, yt, wf, 'Estado del contenido de MP8 (tareas)', sz=1100)
    tot = T8['task']
    col = [2.55, 0.85, 1.30]
    rows = [
        [cell('Levantadas en terreno'), cell(str(T8['task_lev']), algn='ctr'),
         cell(pct(T8['task_lev'], tot) + ' %', algn='ctr')],
        [cell('Propuestas, por validar'), cell(str(T8['task_prop']), algn='ctr'),
         cell(pct(T8['task_prop'], tot) + ' %', algn='ctr')],
        [cell('Con responsable identificado'), cell(str(T8['resp_ok']), algn='ctr'),
         cell(pct(T8['resp_ok'], tot) + ' %', algn='ctr')],
    ]
    t = apa([('Tareas (n = %d)' % tot, 'l'), ('N.º', 'ctr'), ('%', 'ctr')], rows, sz=1000)
    s.add(table(s.nid(), 'Tabla estado', MX, yt + 0.32, col, t, [0.28] + [0.27] * 3))
    fuente(s, 'Figura 3. Flujograma y tabla de actividades y tareas del subproceso C.1 (manual '
              'MP8 v1). En el manual, el texto naranjo identifica propuestas por validar.')
    return s



# ------------------------------------------------------------------ objetivos
def s_objetivos(num):
    s = Slide(L_BODY, label='Objetivos')
    header(s, 'Objetivos', num)
    head(s, MX, BODY_T, CW, 'Objetivo general')
    tb(s, 'OG', MX, BODY_T + 0.34, CW, 0.66,
       [para([r('Estructurar un sistema de gestión de riesgos según COSO II para los '
                'macroprocesos de Gestión de Personas (MP8) y Servicios Generales (MP21) de '
                'Celulosa Arauco y Constitución S.A.', sz=1400, color=INK)], line=108)])
    regla(s, 3.34)
    head(s, MX, 3.48, CW, 'Objetivos específicos')
    oes = [('OE1', 'Modelar', 'la arquitectura de procesos de MP8 y MP21, desde sus procesos '
            'hasta sus tareas, según el marco del Modelo de Gestión Operacional.', 'CAP'),
           ('OE2', 'Evaluar', 'los riesgos operacionales asociados a las tareas modeladas de '
            'ambos macroprocesos según COSO II.', 'TCER'),
           ('OE3', 'Elaborar', 'el sistema de gestión de riesgos, integrando el catálogo de '
            'riesgos y sus planes de mitigación según COSO II.', 'CRPM')]
    y = 3.88
    for code, verb, text, ind in oes:
        tb(s, 'OE cod ' + code, MX, y, 0.62, 0.32, [para([r(code, sz=1300, b=True, color=BLUE_LT)])])
        tb(s, 'OE txt ' + code, MX + 0.66, y, 6.90, 0.62,
           [para([r(verb + ' ', sz=1250, b=True, color=INK), r(text, sz=1250)], line=106)])
        tb(s, 'OE ind ' + code, MX + 7.70, y, 1.40, 0.32,
           [para([r('Indicador: ', sz=950, color=TX_SOFT), r(ind, sz=950, b=True, color=INK)])])
        y += 0.80
    tb(s, 'Nota', MX, 6.28, CW, 0.30,
       [para([r('Objetivos sin modificaciones. Este avance reporta el OE1 y la preparación del '
                'OE2.', sz=1000, i=True, color=TX_SOFT)])], anchor='ctr')
    return s


# ------------------------------------------------------------------ OE1 3: SIPOC MP8
def s_oe1c(num):
    s = Slide(L_BODY, label='OE1 3')
    header(s, 'Desarrollo y resultados del OE1 (3 de 4)', num)
    tb(s, 'Bajada', MX, BODY_T - 0.04, CW, 0.32,
       [para([r('SIPOC del macroproceso MP8: cómo se articulan sus cinco procesos.', sz=1100,
                b=True, color=INK)])], anchor='ctr')
    crop = (45, 50, 1245, 712)
    w = 6.10
    box = imagen(s, 'Captura SIPOC MP8', 'mp8_sipoc', MX, 2.62, w=w)
    f = mapper(box, crop)
    marca(s, f, (64, 211, 442, 435), 1, place='tl')
    marca(s, f, (480, 230, 803, 518), 2, place='tr')
    marca(s, f, (851, 211, 1194, 496), 3, place='tr')
    marca(s, f, (54, 595, 1232, 691), 4, place='bl')
    xr = MX + w + 0.30
    wr = MX + CW - xr
    head(s, xr, 2.62, wr, 'Qué muestra', sz=1150)
    leyenda(s, xr, 3.00, wr, [
        (1, [r('Proveedores y entradas: ', sz=1000, b=True, color=INK),
             r('desde el plan estratégico hasta la normativa laboral.', sz=1000)], 3),
        (2, [r('Cinco procesos', sz=1000, b=True, color=INK),
             r(' encadenados, de la dotación a la compensación.', sz=1000)], 3),
        (3, [r('Salidas y clientes: ', sz=1000, b=True, color=INK),
             r('personas contratadas, habilitadas y retribuidas.', sz=1000)], 3),
        (4, [r('Macroprocesos de soporte', sz=1000, b=True, color=INK),
             r(' con que interactúa MP8.', sz=1000)], 2),
    ], sz=1000, gap=0.10, line_h=0.19)
    fuente(s, 'Figura 4. SIPOC del macroproceso Gestión de Personas (manual MP8 v1). Fuente: '
              'elaboración propia.')
    return s


# ------------------------------------------------------------------ OE1 4: MP21
def s_oe1d(num):
    s = Slide(L_BODY, label='OE1 4')
    header(s, 'Desarrollo y resultados del OE1 (4 de 4)', num)
    tb(s, 'Bajada', MX, BODY_T - 0.04, CW, 0.32,
       [para([r('MP21: estructura definida y levantamiento del ciclo del contrato en curso.',
                sz=1100, b=True, color=INK)])], anchor='ctr')
    crop = (45, 50, 1245, 700)
    w = 6.10
    box = imagen(s, 'Captura ficha MP21', 'mp21_ficha', MX, 2.62, w=w)
    f = mapper(box, crop)
    marca(s, f, (63, 140, 1222, 220), 1, place='tl')
    marca(s, f, (63, 456, 617, 671), 2, place='bl')
    marca(s, f, (63, 228, 1214, 448), 3, place='tr')
    marca(s, f, (646, 456, 1214, 671), 4, place='br')
    xr = MX + w + 0.30
    wr = MX + CW - xr
    head(s, xr, 2.62, wr, 'Estado del manual MP21', sz=1150)
    leyenda(s, xr, 3.00, wr, [
        (1, [r('Objetivo definido', sz=1000, b=True, color=INK),
             r('; propósito propuesto por validar.', sz=1000)], 2),
        (2, [r('Dueño identificado: ', sz=1000, b=True, color=INK),
             r('Subgerencia de Administración y jefes zonales.', sz=1000)], 3),
        (3, [r('Entradas y salidas', sz=1000, b=True, color=INK),
             r(' pendientes de levantamiento.', sz=1000)], 2),
        (4, [r('KPI', sz=1000, b=True, color=INK),
             r(' pendientes de levantamiento.', sz=1000)], 1),
    ], sz=1000, gap=0.10, line_h=0.19)
    regla(s, 5.22, x=xr, w=wr)
    tb(s, 'Estructura', xr, 5.30, wr, 1.30,
       [para([r('%d procesos y %d subprocesos' % (T21['proc'], T21['sub']), sz=1000, b=True,
                color=INK),
              r(' según el mapa de procesos vigente: gestión estratégica, planificación, '
                'administración y término de contrato.', sz=1000)], line=105)])
    fuente(s, 'Figura 5. Ficha resumen de Gestión de Servicios Generales (manual MP21 v1). En '
              'gris, contenido pendiente de levantamiento.')
    return s


# ------------------------------------------------------------------ 9 OE2 1
def s_oe2a(num):
    s = Slide(L_BODY, label='OE2 1')
    header(s, 'Desarrollo y resultados del OE2 (1 de 2)', num)
    tb(s, 'Estado', MX, BODY_T, CW, 0.56,
       [para([r('Actividades realizadas. ', sz=1150, b=True, color=INK),
              r('OE2 se inicia en la semana 12 según el plan; a la fecha no se han realizado '
                'talleres de evaluación. El resultado de este periodo es el universo de tareas '
                'que se evaluará, definido por OE1.', sz=1150)], line=106)])
    head(s, MX, 2.92, 5.60, 'Tareas de MP8 por proceso (universo de OE2)', sz=1150)
    procs = S8['out']
    names = {'A': 'A Necesidad y dotación', 'B': 'B Valoración de cargo',
             'C': 'C Atracción y selección', 'D': 'D Desarrollo organizacional',
             'E': 'E Compensaciones'}
    x_lab, w_lab = MX, 2.05
    x_bar = MX + 2.12
    w_max = 2.85
    vmax = max(d['tasks'] for d in procs.values())
    y = 3.32
    rh = 0.40
    for p, d in procs.items():
        tb(s, 'Barra rotulo ' + p, x_lab, y, w_lab, rh - 0.06,
           [para([r(names[p], sz=1000, color=TX)])], anchor='ctr')
        wl = w_max * d['task_lev'] / vmax
        wp = w_max * d['task_prop'] / vmax
        if wl > 0:
            rect(s, 'Barra lev ' + p, x_bar, y + 0.06, wl, rh - 0.18, fill=INK)
        if wp > 0:
            rect(s, 'Barra prop ' + p, x_bar + wl, y + 0.06, wp, rh - 0.18, fill=BLUE_P)
        tb(s, 'Barra valor ' + p, x_bar + wl + wp + 0.05, y, 0.55, rh - 0.06,
           [para([r(str(d['tasks']), sz=1000, b=True, color=INK)])], anchor='ctr')
        y += rh
    # leyenda del gráfico
    rect(s, 'Ley lev', x_bar, y + 0.10, 0.20, 0.14, fill=INK)
    tb(s, 'Ley lev t', x_bar + 0.25, y + 0.03, 1.20, 0.28, [para([r('Levantadas', sz=950)])],
       anchor='ctr')
    rect(s, 'Ley prop', x_bar + 1.40, y + 0.10, 0.20, 0.14, fill=BLUE_P)
    tb(s, 'Ley prop t', x_bar + 1.65, y + 0.03, 1.60, 0.28,
       [para([r('Propuestas por validar', sz=950)])], anchor='ctr')
    # condiciones de inicio
    xr = MX + 5.95
    wr = MX + CW - xr
    s.add(shape(s.nid(), 'Regla vertical', xr - 0.22, 2.98, 0.014, 3.40, fill=RULE, line=None))
    head(s, xr, 2.92, wr, 'Condiciones para iniciar', sz=1150)
    conds = [
        ('1', 'Validar las %d tareas propuestas.' % T8['task_prop']),
        ('2', 'Identificar %d responsables pendientes.' % T8['resp_pend']),
        ('3', 'Incorporar F15 a F18 al manual MP8.'),
        ('4', 'Levantar actividades y tareas de MP21.'),
    ]
    y = 3.32
    for n, t in conds:
        tb(s, 'Cond num ' + n, xr, y, 0.28, 0.44, [para([r(n, sz=1200, b=True, color=BLUE_LT)])])
        tb(s, 'Cond ' + n, xr + 0.30, y, wr - 0.30, 0.56, [para([r(t, sz=1050)], line=105)])
        y += 0.66
    fuente(s, 'Fuente: elaboración propia a partir del manual MP8 v1. Una tarea no descrita no '
              'puede evaluarse: estas condiciones fijan la cobertura alcanzable del TCER.')
    return s


# ------------------------------------------------------------------ 10 OE2 2
def s_oe2b(num):
    s = Slide(L_BODY, label='OE2 2')
    header(s, 'Desarrollo y resultados del OE2 (2 de 2)', num)
    tb(s, 'Bajada', MX, BODY_T, CW, 0.56,
       [para([r('La matriz de riesgos se construye sobre la tabla de actividades y tareas del '
                'manual: cada fila es una tarea con su código y responsable, a la que se agregan '
                'los siete atributos de COSO II.', sz=1150)], line=106)])
    # dos columnas: origen y atributos
    y0 = 2.95
    wcol = 3.90
    head(s, MX, y0, wcol, 'Desde el manual de procesos (OE1)', sz=1100)
    regla(s, y0 + 0.34, x=MX, w=wcol, color=INK, h=0.018)
    src = ['Código MGO', 'Subproceso y actividad', 'Tarea', 'Responsable']
    y = y0 + 0.42
    for t in src:
        tb(s, 'Origen ' + t, MX, y, wcol, 0.27, [para([r(t, sz=1050)])], anchor='ctr')
        y += 0.29
    s.add(shape(s.nid(), 'Flecha', MX + wcol + 0.18, y0 + 0.95, 0.55, 0.36, geom='rightArrow',
                fill=BLUE_S, line=None, adj=[('1', 50000), ('2', 50000)]))
    x2 = MX + wcol + 0.92
    w2 = MX + CW - x2
    head(s, x2, y0, w2, 'Atributos COSO II por tarea (OE2)', sz=1100)
    regla(s, y0 + 0.34, x=x2, w=w2, color=INK, h=0.018)
    attrs = ['Evento de riesgo', 'Categoría', 'Probabilidad', 'Impacto', 'Riesgo inherente',
             'Controles existentes', 'Riesgo residual']
    half = 4
    for k, t in enumerate(attrs):
        cx = x2 + (0 if k < half else w2 / 2)
        cy = y0 + 0.42 + 0.29 * (k if k < half else k - half)
        tb(s, 'Atributo ' + t, cx, cy, w2 / 2, 0.27, [para([r(t, sz=1050)])], anchor='ctr')
    # ejemplo con datos reales del manual
    ye = 4.72
    head(s, MX, ye, CW, 'Ejemplo de fila con datos del manual MP8 (subproceso C.1)', sz=1100)
    col = [0.62, 3.55, 1.70, 3.23]
    por = 'Se determina en los talleres A2.1 y A2.2'
    rows = [
        [cell('C.1.1', b=True, color=INK), cell('Verificar disponibilidad de la vacante con el área '
                                               'de Personas'),
         cell('Área Solicitante'), cell(por, i=True, color=TX_SOFT)],
        [cell('C.1.1', b=True, color=INK), cell('Priorizar la necesidad según urgencia en reunión '
                                               'de inicio'),
         cell('Pendiente', i=True, color=TX_SOFT), cell(por, i=True, color=TX_SOFT)],
    ]
    t = apa([('Código', 'l'), ('Tarea', 'l'), ('Responsable', 'l'), ('Atributos COSO II', 'l')],
            rows, sz=1000)
    s.add(table(s.nid(), 'Tabla ejemplo', MX, ye + 0.36, col, t, [0.30, 0.46, 0.46]))
    fuente(s, 'Fuente: elaboración propia. Estructura de atributos según la sección 2.3.2 del '
              'informe; tareas y responsables tomados del manual MP8 v1.')
    return s


# ------------------------------------------------------------------ 11 Gantt
GANTT = [
    # (actividad, plan (ini, fin), real (ini, fin, tipo) | None, estado)
    ('A1.1–A1.2 Diagnóstico y entrevistas', (1, 4), (1, 4, 'c'), 'Completada'),
    ('A1.3.1 Flujo único MP8', (3, 5), (3, 5, 'c'), 'Completada'),
    ('A1.3.2 Segmentación MP8', (6, 7), (6, 10, 'e'), '14 de 18 fases'),
    ('A1.4.1 Flujo único MP21', (6, 9), (5, 10, 'e'), 'En curso, atrasada'),
    ('A1.4.2 Segmentación MP21', (10, 11), None, 'No iniciada'),
    ('A1.5–A1.6 Manuales y validación', (11, 13), (10, 10, 'e'), 'Iniciada (v1)'),
    ('A2.1–A2.2 Identificación y valoración', (12, 16), None, 'Planificada'),
    ('A2.3–A2.4 Matriz y validación', (16, 18), None, 'Planificada'),
    ('A3.1–A3.2 Planes y gobernanza', (17, 19), None, 'Planificada'),
    ('A3.3–A3.4 Integración y entrega', (19, 20), None, 'Planificada'),
    ('A-G Documentación del TFG', (1, 20), (1, 10, 'e'), 'En curso'),
]


def s_gantt(num):
    s = Slide(L_BODY, label='Gantt')
    header(s, 'Carta Gantt · Nivel de cumplimiento', num)
    tb(s, 'Bajada', MX, BODY_T - 0.06, CW, 0.34,
       [para([r('Horizonte de 20 semanas, agosto 2026 a enero 2027. Esta presentación '
                'corresponde a la ', sz=1100), r('semana %d' % SEMANA, sz=1100, b=True, color=INK),
              r('.', sz=1100)])], anchor='ctr')
    wa = 2.62
    wk = 0.262
    we = 1.24
    col = [wa] + [wk] * 20 + [we]
    hdr = [cell('Actividad', sz=950, b=True, color=WHITE, fill=INK)]
    for k in range(1, 21):
        hdr.append(cell(str(k), sz=850, b=True, color=WHITE, fill=INK, algn='ctr', marL=0, marR=0))
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
        bold = estado in ('Completada',)
        row.append(cell(estado, sz=900, b=bold, color=INK if bold else TX, lnB=(RULE, 0.5)))
        rows.append(row)
    # hitos
    hit = [cell('Hitos A-P1 · A-P2 · A-PF', sz=900, b=True, color=INK, lnB=(INK, 1.0))]
    for k in range(1, 21):
        fill = None
        if k in (5, SEMANA):
            fill = INK
        elif k == 20:
            fill = BLUE_VP
        hit.append(cell('', sz=600, fill=fill, lnB=(INK, 1.0), marL=0, marR=0))
    hit.append(cell('A-P2 en S%d' % SEMANA, sz=900, color=TX, lnB=(INK, 1.0)))
    rows.append(hit)
    rh = 0.27
    ty = 2.56
    s.add(table(s.nid(), 'Tabla Gantt', MX, ty, col, rows, [0.28] + [rh] * (len(rows) - 1)))
    # línea de la semana actual
    xl = MX + wa + wk * SEMANA
    th = 0.28 + rh * (len(rows) - 1)
    rect(s, 'Semana actual', xl - 0.012, ty, 0.024, th + 0.10, fill='C0504D')
    tb(s, 'Semana actual rotulo', xl - 0.60, ty + th + 0.10, 1.20, 0.22,
       [para([r('Semana %d' % SEMANA, sz=850, b=True, color='C0504D')], algn='ctr')],
       anchor='ctr')
    # leyenda
    ly = ty + th + 0.38
    lx = MX
    for fill, text, w in ((INK, 'Completado', 1.05), (BLUE_LT, 'Ejecutado o en curso', 1.70),
                          (BLUE_VP, 'Planificado', 1.05)):
        rect(s, 'Ley ' + text, lx, ly + 0.05, 0.22, 0.14, fill=fill)
        tb(s, 'Ley t ' + text, lx + 0.27, ly, w, 0.24, [para([r(text, sz=900)])], anchor='ctr')
        lx += w + 0.40
    tb(s, 'Cumplimiento', MX, ly + 0.30, CW, 0.40,
       [para([r('OE1: ', sz=1000, b=True, color=INK),
              r('de 8 actividades, 3 completadas, 3 en curso y 2 no iniciadas. Atraso principal: '
                'levantamiento de MP21.', sz=1000)], line=105)])
    return s


# ------------------------------------------------------------------ cierre
def s_cierre():
    s = Slide(L_CLOSE, label='Cierre')
    tb(s, 'Gracias', 0.80, 3.30, 8.40, 0.62,
       [para([r('Gracias por su atención', sz=2600, b=True, color=INK)], algn='ctr')],
       anchor='ctr')
    tb(s, 'Preguntas', 0.80, 3.98, 8.40, 0.36,
       [para([r('Preguntas y comentarios', sz=1400, color=TX)], algn='ctr')], anchor='ctr')
    tb(s, 'Autoria', 0.80, 6.36, 8.40, 0.56,
       [para([r(AUTOR, sz=1200, b=True, color=INK)], algn='ctr'),
        para([r('Magíster en Sistemas de Gestión Integral de la Calidad · Universidad de La '
                'Frontera', sz=950, color=TX_SOFT)], algn='ctr', space_before=2)],
       anchor='ctr')
    return s


# ------------------------------------------------------------------ anexos
def s_anexo(num, titulo, fname, rot, nota):
    s = Slide(L_BODY, label=titulo)
    header(s, titulo, num)
    pw, ph = png_size(os.path.join(CAP, fname + '.png'))
    maxw, maxh = CW, 4.05
    w = min(maxw, maxh * pw / ph)
    x = MX + (CW - w) / 2
    box = imagen(s, 'Captura ' + fname, fname, x, BODY_T, w=w)
    tb(s, 'Rotulo figura', MX, BODY_T + box[3] + 0.06, CW, 0.26,
       [para([r(rot + ' ', sz=900, b=True, color=INK), r(nota, sz=900, i=True, color=TX_SOFT)])],
       anchor='ctr')
    return s


def make(out):
    slides = [s_portada(), s_titulo()]
    seq = [s_problema1, s_problema2, s_objetivos, s_oe1a, s_oe1b, s_oe1c, s_oe1d, s_oe2a,
           s_oe2b, s_gantt]
    for k, fn in enumerate(seq, 3):
        slides.append(fn(k))
    slides.append(s_cierre())
    build(BASE, out, slides, title='Avance 2 TFG - ' + AUTOR, author=AUTOR, media=MEDIA)
    return slides


if __name__ == '__main__':
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, 'Avance_2_TFG_Bastian_Darwit_v1.pptx')
    sl = make(out)
    print('OK', out, len(sl), 'láminas')
