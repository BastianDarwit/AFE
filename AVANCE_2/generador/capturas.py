# -*- coding: utf-8 -*-
"""Genera las capturas de los manuales MP8 y MP21 usadas en el Avance 2.

Recorte en píxeles de referencia de la lámina a 1280 x 720 (13,33 x 7,5 in).
"""
import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

# nombre: (manual, lámina, (x0, y0, x1, y1))
CAPTURAS = {
    'mp8_mapa_macroprocesos': ('mp8.pptx', 1, (110, 175, 1170, 642)),
    'mp8_apertura': ('mp8.pptx', 2, (150, 135, 1128, 548)),
    'mp8_sipoc': ('mp8.pptx', 4, (45, 50, 1245, 712)),
    'mp8_c1_flujograma': ('mp8.pptx', 27, (60, 55, 1258, 660)),
    'mp8_c1_tareas': ('mp8.pptx', 28, (100, 135, 1220, 462)),
    'mp21_apertura': ('mp21.pptx', 2, (250, 135, 1030, 462)),
    'mp21_ficha': ('mp21.pptx', 3, (45, 50, 1245, 700)),
}


def main(outdir):
    os.makedirs(outdir, exist_ok=True)
    jobs = []
    by_manual = {}
    for name, (man, n, crop) in CAPTURAS.items():
        by_manual.setdefault(man, set()).add(n)
    for man, ns in by_manual.items():
        hdir = os.path.join(outdir, 'html_' + man.split('.')[0])
        subprocess.check_call([sys.executable, os.path.join(HERE, 'render.py'),
                               os.path.join(ROOT, man), hdir,
                               ','.join(str(n) for n in sorted(ns)), '--fontk', '0.88'])
    for name, (man, n, (x0, y0, x1, y1)) in CAPTURAS.items():
        hdir = os.path.join(outdir, 'html_' + man.split('.')[0])
        jobs.append(dict(html=os.path.join(hdir, 'slide%02d.html' % n),
                         png=os.path.join(outdir, name + '.png'),
                         clip=dict(x=x0, y=y0, width=x1 - x0, height=y1 - y0)))
    jp = os.path.join(outdir, 'jobs.json')
    json.dump(jobs, open(jp, 'w'))
    subprocess.check_call(['node', os.path.join(HERE, 'clip.js'), jp, '2'])


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, 'capturas'))
