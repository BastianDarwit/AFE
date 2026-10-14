# -*- coding: utf-8 -*-
"""Estadísticas de un manual MGO: procesos, subprocesos, actividades y tareas."""
import re
import sys
import json
import zipfile
from collections import OrderedDict
from xml.etree import ElementTree as ET

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
ORANGE = 'EA7600'
GRAY = '9F968D'


def runs_info(tc):
    txt = ''
    cols = set()
    ital = False
    for r in tc.iter(A + 'r'):
        t = r.findtext(A + 't') or ''
        if not t.strip():
            txt += t
            continue
        txt += t
        rp = r.find(A + 'rPr')
        if rp is not None:
            sf = rp.find(A + 'solidFill')
            if sf is not None and len(sf):
                cols.add(sf[0].get('val'))
            if rp.get('i') == '1':
                ital = True
    return txt.strip(), cols, ital


def slide_title(root):
    for sp in root.iter(P + 'sp'):
        t = ''.join(x.text or '' for x in sp.iter(A + 't'))
        if '|' in t and ('Actividades y tareas' in t or 'Flujograma' in t or 'Ficha' in t
                         or 'SIPOC' in t or 'Apertura' in t):
            return t
    return ''


def main(path):
    z = zipfile.ZipFile(path)
    names = sorted([n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml$', n)],
                   key=lambda n: int(re.findall(r'\d+', n)[0]))
    procs = OrderedDict()
    tasks = []
    flows = {}
    for n in names:
        num = int(re.findall(r'\d+', n)[0])
        root = ET.fromstring(z.read(n))
        title = slide_title(root)
        m = re.match(r'\s*([A-Z])\.(\d+) (Actividades y tareas|Flujograma subproceso)\|\s*(.*)', title)
        if m:
            p, sp, kind, name = m.group(1), m.group(2), m.group(3), m.group(4).strip()
            procs.setdefault(p, OrderedDict()).setdefault(p + '.' + sp, name)
            if kind.startswith('Flujograma'):
                codes = set()
                for sp_ in root.iter(P + 'sp'):
                    t = ''.join(x.text or '' for x in sp_.iter(A + 't')).strip()
                    if re.fullmatch(r'%s\.%s\.\d+' % (p, sp), t):
                        codes.add(t)
                flows[p + '.' + sp] = dict(slide=num, activities=sorted(codes, key=lambda c: int(c.split('.')[-1])))
                continue
            # tabla de actividades y tareas
            for tbl in root.iter(A + 'tbl'):
                rows = tbl.findall(A + 'tr')
                last_act = None
                for tr in rows[1:]:
                    tcs = tr.findall(A + 'tc')
                    if len(tcs) < 3:
                        continue
                    act, acols, _ = runs_info(tcs[0])
                    if tcs[0].get('vMerge') == '1' or not act:
                        act = last_act
                    last_act = act
                    task, tcols, _ = runs_info(tcs[1])
                    resp, rcols, rital = runs_info(tcs[2])
                    if not task:
                        continue
                    tasks.append(dict(slide=num, sub=p + '.' + sp, act=act, task=task,
                                      task_prop=ORANGE in tcols, resp=resp,
                                      resp_pend=(resp.lower().startswith('pendiente') or GRAY in rcols),
                                      resp_prop=ORANGE in rcols))
    out = OrderedDict()
    tot = dict(proc=0, sub=0, act=0, task=0, task_lev=0, task_prop=0, resp_ok=0, resp_pend=0)
    for p, subs in procs.items():
        acts = sum(len(flows.get(s, {}).get('activities', [])) for s in subs)
        ts = [t for t in tasks if t['sub'].startswith(p + '.')]
        d = dict(subs=len(subs), acts=acts, tasks=len(ts),
                 task_lev=sum(1 for t in ts if not t['task_prop']),
                 task_prop=sum(1 for t in ts if t['task_prop']),
                 resp_ok=sum(1 for t in ts if not t['resp_pend']),
                 resp_pend=sum(1 for t in ts if t['resp_pend']),
                 subnames=list(subs.items()))
        out[p] = d
        tot['proc'] += 1
        tot['sub'] += d['subs']
        tot['act'] += acts
        tot['task'] += d['tasks']
        tot['task_lev'] += d['task_lev']
        tot['task_prop'] += d['task_prop']
        tot['resp_ok'] += d['resp_ok']
        tot['resp_pend'] += d['resp_pend']
    return out, tot, tasks, flows


if __name__ == '__main__':
    out, tot, tasks, flows = main(sys.argv[1])
    for p, d in out.items():
        print(p, {k: v for k, v in d.items() if k != 'subnames'})
        for s, nm in d['subnames']:
            print('    ', s, nm, '| act:', len(flows.get(s, {}).get('activities', [])),
                  '| tareas:', sum(1 for t in tasks if t['sub'] == s), '| flujo lam.', flows.get(s, {}).get('slide'))
    print('TOTAL', tot)
    if len(sys.argv) > 2:
        json.dump(dict(out=out, tot=tot, tasks=tasks, flows=flows), open(sys.argv[2], 'w'), ensure_ascii=False, indent=1)
