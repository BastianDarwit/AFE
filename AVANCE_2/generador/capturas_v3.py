# -*- coding: utf-8 -*-
"""Capturas del manual MP21 v2 para el Avance 2 v3 (no reemplaza las del v1).

Recorte en píxeles de referencia de la lámina a 1280 x 720.
"""
import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

MAN = 'mp21_v2.pptx'
# nombre: (lámina, (x0, y0, x1, y1))
CAPTURAS = {
    'mp21v2_apertura': (2, (250, 135, 1030, 415)),
    'mp21v2_sipoc': (4, (45, 112, 1245, 700)),
    'mp21v2_c1_flujograma': (21, (60, 115, 1258, 660)),
    'mp21v2_c1_tareas': (22, (30, 135, 1220, 650)),
    'mp21v2_c2_flujograma': (23, (60, 115, 1258, 660)),
    'mp21v2_c2_tareas': (24, (30, 135, 1220, 500)),
}


def main(outdir):
    hdir = os.path.join(outdir, 'html_mp21v2')
    ns = sorted({n for n, _ in CAPTURAS.values()})
    subprocess.check_call([sys.executable, os.path.join(HERE, 'render.py'),
                           os.path.join(ROOT, MAN), hdir, ','.join(map(str, ns)),
                           '--fontk', '0.88'])
    jobs = []
    for name, (n, (x0, y0, x1, y1)) in CAPTURAS.items():
        jobs.append(dict(html=os.path.join(hdir, 'slide%02d.html' % n),
                         png=os.path.join(outdir, name + '.png'),
                         clip=dict(x=x0, y=y0, width=x1 - x0, height=y1 - y0)))
    jp = os.path.join(outdir, 'jobs_v3.json')
    json.dump(jobs, open(jp, 'w'))
    subprocess.check_call(['node', os.path.join(HERE, 'clip.js'), jp, '2'])


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, 'capturas'))
