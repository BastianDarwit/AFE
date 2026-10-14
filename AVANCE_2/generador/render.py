# -*- coding: utf-8 -*-
"""Renderizador aproximado PPTX -> HTML (para capturas y control visual).

Uso: python3 render.py archivo.pptx salida_dir [n1,n2,...] [--scale 1.0] [--fontk 1.0]
Genera salida_dir/slideNN.html; luego shot.js las convierte a PNG.
"""
import base64
import math
import os
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
A = '{%s}' % NS['a']
P = '{%s}' % NS['p']
R = '{%s}' % NS['r']
EMU = 914400.0
PRST = {'black': '000000', 'white': 'FFFFFF', 'red': 'FF0000', 'gray': '808080',
        'blue': '0000FF', 'green': '008000', 'yellow': 'FFFF00'}


def esc(t):
    return (t or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


class Pkg:
    def __init__(self, path):
        self.z = zipfile.ZipFile(path)
        self.names = set(self.z.namelist())
        self.cache = {}

    def xml(self, name):
        if name not in self.cache:
            self.cache[name] = ET.fromstring(self.z.read(name))
        return self.cache[name]

    def rels(self, name):
        d, b = os.path.split(name)
        rn = d + '/_rels/' + b + '.rels'
        out = {}
        if rn in self.names:
            for r in self.xml(rn):
                tgt = r.get('Target')
                if r.get('TargetMode') == 'External':
                    full = tgt
                else:
                    full = os.path.normpath(os.path.join(d, tgt)).replace('\\', '/')
                out[r.get('Id')] = (r.get('Type').split('/')[-1], full)
        return out


# ------------------------------------------------------------------ color
def hex2rgb(h):
    return [int(h[i:i + 2], 16) / 255.0 for i in (0, 2, 4)]


def rgb2hsl(r, g, b):
    mx, mn = max(r, g, b), min(r, g, b)
    l = (mx + mn) / 2
    if mx == mn:
        return 0.0, 0.0, l
    d = mx - mn
    s = d / (2 - mx - mn) if l > 0.5 else d / (mx + mn)
    if mx == r:
        h = (g - b) / d + (6 if g < b else 0)
    elif mx == g:
        h = (b - r) / d + 2
    else:
        h = (r - g) / d + 4
    return h / 6, s, l


def hsl2rgb(h, s, l):
    if s == 0:
        return l, l, l

    def f(p, q, t):
        t %= 1
        if t < 1 / 6:
            return p + (q - p) * 6 * t
        if t < 1 / 2:
            return q
        if t < 2 / 3:
            return p + (q - p) * (2 / 3 - t) * 6
        return p
    q = l * (1 + s) if l < 0.5 else l + s - l * s
    p = 2 * l - q
    return f(p, q, h + 1 / 3), f(p, q, h), f(p, q, h - 1 / 3)


class Ctx:
    def __init__(self, pkg, theme, clrmap):
        self.pkg = pkg
        self.scheme = {}
        self.fonts = {'major': 'Calibri', 'minor': 'Calibri'}
        cs = theme.find('.//' + A + 'clrScheme')
        for c in cs:
            tag = c.tag.split('}')[1]
            sub = c[0]
            if sub.tag == A + 'srgbClr':
                self.scheme[tag] = sub.get('val')
            elif sub.tag == A + 'sysClr':
                self.scheme[tag] = sub.get('lastClr', '000000')
        self.clrmap = clrmap
        fs = theme.find('.//' + A + 'fontScheme')
        if fs is not None:
            mj = fs.find(A + 'majorFont/' + A + 'latin')
            mn = fs.find(A + 'minorFont/' + A + 'latin')
            if mj is not None:
                self.fonts['major'] = mj.get('typeface')
            if mn is not None:
                self.fonts['minor'] = mn.get('typeface')
        self.fmt = theme.find('.//' + A + 'fmtScheme')

    def color(self, el, phclr=None):
        """el: un elemento de color (srgbClr, schemeClr...). Devuelve css."""
        if el is None:
            return None
        tag = el.tag.split('}')[1]
        alpha = 1.0
        if tag == 'srgbClr':
            base = el.get('val')
        elif tag == 'schemeClr':
            v = el.get('val')
            if v == 'phClr':
                base = phclr or '000000'
            else:
                v = self.clrmap.get(v, v)
                base = self.scheme.get(v, '000000')
        elif tag == 'prstClr':
            base = PRST.get(el.get('val'), '000000')
        elif tag == 'sysClr':
            base = el.get('lastClr', '000000')
        elif tag == 'scrgbClr':
            base = '%02X%02X%02X' % tuple(int(float(el.get(k)) / 100000 * 255) for k in 'rgb')
        else:
            return None
        r, g, b = hex2rgb(base)
        for m in el:
            mt = m.tag.split('}')[1]
            val = int(m.get('val', '0')) / 100000.0
            if mt in ('lumMod', 'lumOff'):
                h, s, l = rgb2hsl(r, g, b)
                l = l * val if mt == 'lumMod' else min(1, l + val)
                r, g, b = hsl2rgb(h, s, l)
            elif mt == 'tint':
                r, g, b = [c + (1 - c) * (1 - val) for c in (r, g, b)]
            elif mt == 'shade':
                r, g, b = [c * val for c in (r, g, b)]
            elif mt == 'alpha':
                alpha = val
        r, g, b = [max(0, min(255, int(round(c * 255)))) for c in (r, g, b)]
        if alpha < 1:
            return 'rgba(%d,%d,%d,%.3f)' % (r, g, b, alpha)
        return '#%02X%02X%02X' % (r, g, b)

    def hexof(self, el):
        c = self.color(el)
        return c


def first_color(el):
    if el is None:
        return None
    for c in el:
        if c.tag.split('}')[1] in ('srgbClr', 'schemeClr', 'prstClr', 'sysClr', 'scrgbClr'):
            return c
    return None


# ------------------------------------------------------------------ geometry
def preset_path(prst, w, h, av):
    """Devuelve un 'd' de SVG para geometrías frecuentes."""
    def adj(name, default):
        return av.get(name, default)
    if prst in ('rect', 'flowChartProcess', 'snip1Rect', 'flowChartPredefinedProcess', 'frame'):
        return 'M0,0H%.2fV%.2fH0Z' % (w, h)
    if prst in ('roundRect', 'flowChartAlternateProcess', 'round2SameRect'):
        r = min(w, h) * adj('adj', 16667) / 100000.0
        if prst == 'flowChartAlternateProcess':
            r = min(w, h) * 0.1667
        return ('M%.2f,0H%.2fA%.2f,%.2f 0 0 1 %.2f,%.2fV%.2fA%.2f,%.2f 0 0 1 %.2f,%.2fH%.2f'
                'A%.2f,%.2f 0 0 1 0,%.2fV%.2fA%.2f,%.2f 0 0 1 %.2f,0Z') % (
            r, w - r, r, r, w, r, h - r, r, r, w - r, h, r, r, r, h - r, r, r, r, r)
    if prst in ('ellipse', 'flowChartConnector', 'donut'):
        return 'M0,%.2fA%.2f,%.2f 0 1 0 %.2f,%.2fA%.2f,%.2f 0 1 0 0,%.2fZ' % (
            h / 2, w / 2, h / 2, w, h / 2, w / 2, h / 2, h / 2)
    if prst in ('diamond', 'flowChartDecision'):
        return 'M%.2f,0L%.2f,%.2fL%.2f,%.2fL0,%.2fZ' % (w / 2, w, h / 2, w / 2, h, h / 2)
    if prst in ('triangle', 'isosTriangle'):
        return 'M%.2f,0L%.2f,%.2fL0,%.2fZ' % (w / 2, w, h, h)
    if prst == 'rtTriangle':
        return 'M0,0L%.2f,%.2fL0,%.2fZ' % (w, h, h)
    if prst == 'homePlate':
        a = min(w, h) * adj('adj', 50000) / 100000.0
        return 'M0,0H%.2fL%.2f,%.2fL%.2f,%.2fH0Z' % (w - a, w, h / 2, w - a, h)
    if prst == 'chevron':
        a = min(w, h) * adj('adj', 50000) / 100000.0
        return 'M0,0H%.2fL%.2f,%.2fL%.2f,%.2fH0L%.2f,%.2fZ' % (w - a, w, h / 2, w - a, h, a, h / 2)
    if prst in ('rightArrow', 'leftArrow', 'downArrow', 'upArrow'):
        a1 = adj('adj1', 50000) / 100000.0
        a2 = adj('adj2', 50000) / 100000.0
        if prst in ('rightArrow', 'leftArrow'):
            th = h * a1
            hl = min(w, h * a2 * 1.0)
            y0, y1 = (h - th) / 2, (h + th) / 2
            d = 'M0,%.2fH%.2fV0L%.2f,%.2fL%.2f,%.2fV%.2fH0Z' % (y0, w - hl, w, h / 2, w - hl, h, y1)
            if prst == 'leftArrow':
                d = 'M%.2f,%.2fH%.2fV0L0,%.2fL%.2f,%.2fV%.2fH%.2fZ' % (w, y0, hl, h / 2, hl, h, y1, w)
            return d
        th = w * a1
        hl = min(h, w * a2)
        x0, x1 = (w - th) / 2, (w + th) / 2
        if prst == 'downArrow':
            return 'M%.2f,0H%.2fV%.2fH%.2fL%.2f,%.2fL0,%.2fH%.2fZ' % (x0, x1, h - hl, w, w / 2, h, h - hl, x0)
        return 'M%.2f,%.2fH%.2fV0' % (x0, h, x1) if False else 'M%.2f,%.2fV%.2fH0L%.2f,0L%.2f,%.2fH%.2fV%.2fZ' % (x0, h, hl, w / 2, w, hl, x1, h)
    if prst == 'flowChartTerminator':
        r = min(w / 2, h / 2)
        return 'M%.2f,0H%.2fA%.2f,%.2f 0 0 1 %.2f,%.2fH%.2fA%.2f,%.2f 0 0 1 %.2f,0Z' % (
            r, w - r, r, h / 2, w - r, h, r, r, h / 2, r)
    if prst in ('flowChartDocument',):
        return 'M0,0H%.2fV%.2fC%.2f,%.2f %.2f,%.2f 0,%.2fZ' % (w, h * 0.83, w * 0.6, h * 0.6, w * 0.4, h * 1.1, h * 0.9)
    if prst in ('parallelogram', 'flowChartInputOutput'):
        a = w * 0.2
        return 'M%.2f,0H%.2fL%.2f,%.2fH0Z' % (a, w, w - a, h)
    if prst == 'hexagon':
        a = w * 0.25
        return 'M%.2f,0H%.2fL%.2f,%.2fL%.2f,%.2fH%.2fL0,%.2fZ' % (a, w - a, w, h / 2, w - a, h, a, h / 2)
    if prst == 'plus':
        a = min(w, h) * 0.25
        return ('M%.2f,0H%.2fV%.2fH%.2fV%.2fH%.2fV%.2fH%.2fV%.2fH0V%.2fH%.2fZ' %
                (a, w - a, a, w, h - a, w - a, h, a, h - a, a, a))
    if prst in ('line', 'straightConnector1'):
        return 'M0,0L%.2f,%.2f' % (w, h)
    return 'M0,0H%.2fV%.2fH0Z' % (w, h)


def cust_path(cg, w, h):
    out = []
    for path in cg.findall('.//' + A + 'path'):
        pw = float(path.get('w', '0') or 0) or w
        ph = float(path.get('h', '0') or 0) or h
        sx = w / pw if pw else 1
        sy = h / ph if ph else 1
        for c in path:
            t = c.tag.split('}')[1]
            pts = [(float(pt.get('x')) * sx, float(pt.get('y')) * sy) for pt in c.findall(A + 'pt')]
            if t == 'moveTo':
                out.append('M%.2f,%.2f' % pts[0])
            elif t == 'lnTo':
                out.append('L%.2f,%.2f' % pts[0])
            elif t == 'cubicBezTo':
                out.append('C' + ' '.join('%.2f,%.2f' % p for p in pts))
            elif t == 'quadBezTo':
                out.append('Q' + ' '.join('%.2f,%.2f' % p for p in pts))
            elif t == 'close':
                out.append('Z')
    return ' '.join(out)


# ------------------------------------------------------------------ renderer
class Renderer:
    def __init__(self, path, scale=96.0, fontk=1.0):
        self.pkg = Pkg(path)
        self.scale = scale  # px por pulgada
        self.fontk = fontk
        pres = self.pkg.xml('ppt/presentation.xml')
        sz = pres.find(P + 'sldSz')
        self.W = int(sz.get('cx'))
        self.H = int(sz.get('cy'))
        prels = self.pkg.rels('ppt/presentation.xml')
        self.slides = []
        for sid in pres.find(P + 'sldIdLst'):
            rid = sid.get(R + 'id')
            self.slides.append((prels[rid][1], sid.get('show') == '0'))

    def px(self, emu):
        return emu / EMU * self.scale

    # --- herencia de texto ---
    def master_styles(self, master):
        ts = master.find(P + 'txStyles')
        return ts

    def render_slide(self, idx):
        sname, hidden = self.slides[idx]
        pkg = self.pkg
        srels = pkg.rels(sname)
        lay = [v[1] for v in srels.values() if v[0] == 'slideLayout'][0]
        lrels = pkg.rels(lay)
        mas = [v[1] for v in lrels.values() if v[0] == 'slideMaster'][0]
        mrels = pkg.rels(mas)
        theme = [v[1] for v in mrels.values() if v[0] == 'theme'][0]
        master = pkg.xml(mas)
        cm = master.find(P + 'clrMap')
        clrmap = dict(cm.attrib)
        slide = pkg.xml(sname)
        layout = pkg.xml(lay)
        for src in (slide, layout):
            ov = src.find(P + 'clrMapOvr/' + A + 'overrideClrMapping')
            if ov is not None:
                clrmap = dict(ov.attrib)
                break
        self.ctx = Ctx(pkg, pkg.xml(theme), clrmap)
        self.master = master
        self.layout = layout
        self.txstyles = master.find(P + 'txStyles')
        html = []
        W, H = self.px(self.W), self.px(self.H)
        bg = self.background([(slide, sname), (layout, lay), (master, mas)])
        html.append('<div class="slide" style="width:%.1fpx;height:%.1fpx;%s">' % (W, H, bg))
        # master y layout: formas no marcadoras
        show_master = layout.get('showMasterSp', '1') != '0' and slide.get('showMasterSp', '1') != '0'
        if show_master:
            html.append(self.tree(master.find(P + 'cSld/' + P + 'spTree'), mas, skip_ph=True))
        html.append(self.tree(layout.find(P + 'cSld/' + P + 'spTree'), lay, skip_ph=True))
        self.ph_sources = [(layout, lay), (master, mas)]
        html.append(self.tree(slide.find(P + 'cSld/' + P + 'spTree'), sname, skip_ph=False))
        html.append('</div>')
        return '\n'.join(html), W, H

    def background(self, srcs):
        for el, name in srcs:
            bg = el.find(P + 'cSld/' + P + 'bg')
            if bg is None:
                continue
            bp = bg.find(P + 'bgPr')
            if bp is not None:
                sf = bp.find(A + 'solidFill')
                if sf is not None:
                    return 'background:%s;' % self.ctx.color(first_color(sf))
                bf = bp.find(A + 'blipFill')
                if bf is not None:
                    uri = self.blip_uri(bf, name)
                    if uri:
                        return "background:url('%s');background-size:100%% 100%%;" % uri
                gf = bp.find(A + 'gradFill')
                if gf is not None:
                    gs = gf.findall('.//' + A + 'gs')
                    if gs:
                        return 'background:%s;' % self.ctx.color(first_color(gs[0]))
            ref = bg.find(P + 'bgRef')
            if ref is not None:
                c = self.ctx.color(first_color(ref))
                return 'background:%s;' % (c or '#fff')
        return 'background:#fff;'

    def blip_uri(self, bf, partname):
        blip = bf.find(A + 'blip')
        if blip is None:
            return None
        rid = blip.get(R + 'embed')
        rels = self.pkg.rels(partname)
        if rid not in rels:
            return None
        t = rels[rid][1]
        if t not in self.pkg.names:
            return None
        data = self.pkg.z.read(t)
        ext = t.rsplit('.', 1)[-1].lower()
        mime = {'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg', 'gif': 'image/gif',
                'svg': 'image/svg+xml', 'emf': None, 'wmf': None}.get(ext, 'image/png')
        if mime is None:
            return None
        return 'data:%s;base64,%s' % (mime, base64.b64encode(data).decode())

    # --- recorrido ---
    def tree(self, sptree, part, skip_ph, xf=None):
        out = []
        if sptree is None:
            return ''
        for el in sptree:
            tag = el.tag.split('}')[1]
            if tag == 'sp':
                ph = el.find(P + 'nvSpPr/' + P + 'nvPr/' + P + 'ph')
                if ph is not None and skip_ph:
                    continue
                if ph is not None and ph.get('type') in ('sldNum', 'dt', 'ftr'):
                    continue
                out.append(self.shape(el, part, xf, ph))
            elif tag == 'cxnSp':
                out.append(self.shape(el, part, xf, None, conn=True))
            elif tag == 'grpSp':
                out.append(self.group(el, part, skip_ph, xf))
            elif tag == 'pic':
                out.append(self.pic(el, part, xf))
            elif tag == 'graphicFrame':
                out.append(self.frame(el, part, xf))
            elif tag == 'AlternateContent':
                fb = el.find('{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback')
                if fb is not None:
                    out.append(self.tree(fb, part, skip_ph, xf))
        return '\n'.join(o for o in out if o)

    def xfrm_of(self, el, xfpath):
        x = el.find(xfpath)
        if x is None:
            return None
        off = x.find(A + 'off')
        ext = x.find(A + 'ext')
        if off is None or ext is None:
            return None
        return dict(x=int(off.get('x')), y=int(off.get('y')), w=int(ext.get('cx')), h=int(ext.get('cy')),
                    rot=int(x.get('rot', '0')) / 60000.0, fh=x.get('flipH') == '1', fv=x.get('flipV') == '1',
                    node=x)

    def apply(self, xf, b):
        """Aplica la transformación de grupo xf (función) a la caja b."""
        if xf is None:
            return b
        return xf(b)

    def group(self, el, part, skip_ph, xf):
        g = self.xfrm_of(el, P + 'grpSpPr/' + A + 'xfrm')
        if g is None:
            return self.tree(el, part, skip_ph, xf)
        node = g['node']
        cho = node.find(A + 'chOff')
        che = node.find(A + 'chExt')
        cx, cy = int(cho.get('x')), int(cho.get('y'))
        cw, ch = int(che.get('cx')) or 1, int(che.get('cy')) or 1
        sx, sy = g['w'] / cw, g['h'] / ch
        parent = xf

        def f(b, g=g, cx=cx, cy=cy, sx=sx, sy=sy, parent=parent):
            nb = dict(b)
            nb['x'] = g['x'] + (b['x'] - cx) * sx
            nb['y'] = g['y'] + (b['y'] - cy) * sy
            nb['w'] = b['w'] * sx
            nb['h'] = b['h'] * sy
            nb['rot'] = b.get('rot', 0) + g['rot']
            return parent(nb) if parent else nb
        return self.tree(el, part, skip_ph, f)

    def ph_lookup(self, ph):
        """Busca el marcador equivalente en layout/master para heredar posición y estilo."""
        typ = ph.get('type', 'body')
        idx = ph.get('idx')
        found = []
        for src, name in self.ph_sources:
            for sp in src.iter(P + 'sp'):
                p2 = sp.find(P + 'nvSpPr/' + P + 'nvPr/' + P + 'ph')
                if p2 is None:
                    continue
                t2 = p2.get('type', 'body')
                if (idx is not None and p2.get('idx') == idx) or (t2 == typ and (idx is None or typ != 'body')):
                    found.append(sp)
                    break
        return found

    def shape(self, el, part, xf, ph, conn=False):
        sppr = el.find(P + 'spPr')
        b = self.xfrm_of(el, P + 'spPr/' + A + 'xfrm')
        inherited = []
        if ph is not None:
            inherited = self.ph_lookup(ph)
            if b is None:
                for sp in inherited:
                    b = self.xfrm_of(sp, P + 'spPr/' + A + 'xfrm')
                    if b:
                        break
        if b is None:
            return ''
        b = self.apply(xf, b)
        X, Y, Wd, Ht = self.px(b['x']), self.px(b['y']), self.px(b['w']), self.px(b['h'])
        style = el.find(P + 'style')
        # geometría
        geom = sppr.find(A + 'prstGeom') if sppr is not None else None
        cgeom = sppr.find(A + 'custGeom') if sppr is not None else None
        prst = geom.get('prst') if geom is not None else ('cust' if cgeom is not None else 'rect')
        av = {}
        if geom is not None:
            for gd in geom.findall('.//' + A + 'gd'):
                m = re.match(r'val (-?\d+)', gd.get('fmla', ''))
                if m:
                    av[gd.get('name')] = int(m.group(1))
        # relleno
        fill = self.fill_of(sppr, style, part)
        line = self.line_of(sppr, style, conn or prst in ('line', 'straightConnector1') or prst.startswith('bentConnector') or prst.startswith('curvedConnector'))
        parts = []
        tr = []
        if b.get('rot'):
            tr.append('rotate(%.2fdeg)' % b['rot'])
        if b.get('fh'):
            tr.append('scaleX(-1)')
        if b.get('fv'):
            tr.append('scaleY(-1)')
        trs = 'transform:%s;' % ' '.join(tr) if tr else ''
        box = 'left:%.2fpx;top:%.2fpx;width:%.2fpx;height:%.2fpx;' % (X, Y, Wd, Ht)
        svg = ''
        if fill or line:
            if prst.startswith('bentConnector'):
                adj1 = av.get('adj1', 50000) / 100000.0
                if prst == 'bentConnector2':
                    d = 'M0,0H%.2fV%.2f' % (Wd, Ht)
                else:
                    d = 'M0,0H%.2fV%.2fH%.2f' % (Wd * adj1, Ht, Wd)
            elif prst.startswith('curvedConnector'):
                d = 'M0,0C%.2f,0 %.2f,%.2f %.2f,%.2f' % (Wd / 2, Wd / 2, Ht, Wd, Ht)
            elif cgeom is not None:
                d = cust_path(cgeom, Wd, Ht)
            else:
                d = preset_path(prst, Wd, Ht, av)
            open_path = prst in ('line', 'straightConnector1') or prst.startswith('bentConnector') or prst.startswith('curvedConnector') or conn
            markers = ''
            mk = ''
            if line:
                lw, lc, dash, head, tail = line
                if head or tail:
                    mid = 'm%d' % id(el)
                    ms = max(3.0, lw * 3.0)
                    markers = ('<defs><marker id="%s_e" markerWidth="%.1f" markerHeight="%.1f" refX="%.1f" refY="%.1f" orient="auto" markerUnits="userSpaceOnUse">'
                               '<path d="M0,0L%.1f,%.1fL0,%.1fZ" fill="%s"/></marker>'
                               '<marker id="%s_s" markerWidth="%.1f" markerHeight="%.1f" refX="0" refY="%.1f" orient="auto" markerUnits="userSpaceOnUse">'
                               '<path d="M%.1f,0L0,%.1fL%.1f,%.1fZ" fill="%s"/></marker></defs>') % (
                        mid, ms, ms, ms, ms / 2, ms, ms / 2, ms, lc, mid, ms, ms, ms / 2, ms, ms / 2, ms, ms, lc)
                    if tail:
                        mk += ' marker-end="url(#%s_e)"' % mid
                    if head:
                        mk += ' marker-start="url(#%s_s)"' % mid
            stroke = ''
            if line:
                lw, lc, dash, head, tail = line
                stroke = ' stroke="%s" stroke-width="%.2f"%s' % (lc, lw, (' stroke-dasharray="%s"' % dash) if dash else '')
            else:
                stroke = ' stroke="none"'
            fl = fill if (fill and not open_path) else 'none'
            svg = ('<svg class="g" width="%.2f" height="%.2f" style="overflow:visible">%s'
                   '<path d="%s" fill="%s"%s%s/></svg>') % (max(Wd, 1), max(Ht, 1), markers, d, fl, stroke, mk)
        nm = el.find('.//' + P + 'cNvPr').get('name', '')
        if 'slideLayout' in part or 'slideMaster' in part:
            nm = '_' + nm
        parts.append('<div class="sp" data-name="%s" style="%s%s">%s' % (
            esc(nm), box, trs, svg))
        tb = el.find(P + 'txBody')
        if tb is not None:
            parts.append(self.text(tb, Wd, Ht, inherited, ph, style))
        parts.append('</div>')
        return ''.join(parts)

    def fill_of(self, sppr, style, part):
        if sppr is not None:
            for c in sppr:
                t = c.tag.split('}')[1]
                if t == 'noFill':
                    return None
                if t == 'solidFill':
                    return self.ctx.color(first_color(c))
                if t == 'gradFill':
                    gs = c.findall('.//' + A + 'gs')
                    if gs:
                        return self.ctx.color(first_color(gs[len(gs) // 2]))
                if t == 'blipFill':
                    return None
                if t == 'grpFill':
                    return None
        if style is not None:
            fr = style.find(A + 'fillRef')
            if fr is not None and fr.get('idx', '0') != '0':
                return self.ctx.color(first_color(fr))
        return None

    def line_of(self, sppr, style, is_line):
        ln = sppr.find(A + 'ln') if sppr is not None else None
        color = None
        width = None
        dash = None
        head = tail = False
        if style is not None:
            lr = style.find(A + 'lnRef')
            if lr is not None and lr.get('idx', '0') != '0':
                color = self.ctx.color(first_color(lr))
                width = 0.75
        if ln is not None:
            if ln.find(A + 'noFill') is not None:
                return None
            sf = ln.find(A + 'solidFill')
            if sf is not None:
                color = self.ctx.color(first_color(sf))
            if ln.get('w'):
                width = int(ln.get('w')) / 12700.0
            pd = ln.find(A + 'prstDash')
            if pd is not None and pd.get('val') not in ('solid', None):
                dash = '4,3' if 'dash' in pd.get('val').lower() else '1,2'
            he = ln.find(A + 'headEnd')
            te = ln.find(A + 'tailEnd')
            head = he is not None and he.get('type', 'none') != 'none'
            tail = te is not None and te.get('type', 'none') != 'none'
        if color is None:
            return None
        if width is None:
            width = 0.75
        return (max(0.5, width * self.scale / 72.0), color, dash, head, tail)

    # --- texto ---
    def lvl_defaults(self, ph, inherited, lvl):
        """Propiedades por nivel: master txStyles -> marcadores heredados."""
        d = {'sz': 1800, 'b': False, 'i': False, 'color': None, 'algn': 'l', 'font': None,
             'bullet': None, 'marL': 0, 'indent': 0, 'lnSpc': None, 'spcBef': 0, 'spcAft': 0}
        chain = []
        if self.txstyles is not None:
            if ph is None:
                st = self.txstyles.find(P + 'otherStyle')
            else:
                t = ph.get('type', 'body')
                st = self.txstyles.find(P + ('titleStyle' if t in ('title', 'ctrTitle') else 'bodyStyle'))
                if t in ('sldNum', 'dt', 'ftr'):
                    st = self.txstyles.find(P + 'otherStyle')
            if st is not None:
                chain.append(st)
        for sp in reversed(inherited):
            ls = sp.find(P + 'txBody/' + A + 'lstStyle')
            if ls is not None:
                chain.append(ls)
        for st in chain:
            lp = st.find(A + 'lvl%dpPr' % (lvl + 1))
            if lp is not None:
                self.merge_ppr(d, lp)
        return d

    def merge_ppr(self, d, pp):
        if pp.get('algn'):
            d['algn'] = pp.get('algn')
        if pp.get('marL'):
            d['marL'] = int(pp.get('marL'))
        if pp.get('indent'):
            d['indent'] = int(pp.get('indent'))
        ls = pp.find(A + 'lnSpc/' + A + 'spcPct')
        if ls is not None:
            d['lnSpc'] = int(ls.get('val')) / 100000.0
        for k, tagn in (('spcBef', 'spcBef'), ('spcAft', 'spcAft')):
            sp = pp.find(A + tagn + '/' + A + 'spcPts')
            if sp is not None:
                d[k] = int(sp.get('val')) / 100.0
        if pp.find(A + 'buNone') is not None:
            d['bullet'] = None
        bc = pp.find(A + 'buChar')
        if bc is not None:
            d['bullet'] = bc.get('char')
        if pp.find(A + 'buAutoNum') is not None:
            d['bullet'] = '1.'
        dr = pp.find(A + 'defRPr')
        if dr is not None:
            self.merge_rpr(d, dr)

    def merge_rpr(self, d, rp):
        if rp.get('sz'):
            d['sz'] = int(rp.get('sz'))
        if rp.get('b') is not None:
            d['b'] = rp.get('b') == '1'
        if rp.get('i') is not None:
            d['i'] = rp.get('i') == '1'
        if rp.get('u') is not None:
            d['u'] = rp.get('u') not in ('none',)
        if rp.get('cap') is not None:
            d['cap'] = rp.get('cap')
        sf = rp.find(A + 'solidFill')
        if sf is not None:
            d['color'] = self.ctx.color(first_color(sf))
        lt = rp.find(A + 'latin')
        if lt is not None:
            d['font'] = lt.get('typeface')

    def text(self, tb, Wd, Ht, inherited, ph, style):
        bp = tb.find(A + 'bodyPr')
        ins = [91440, 45720, 91440, 45720]
        anchor = 't'
        wrap = True
        vert = None
        fontscale = 1.0
        lnred = 0.0
        bps = [sp.find(P + 'txBody/' + A + 'bodyPr') for sp in reversed(inherited)] + [bp]
        for b in bps:
            if b is None:
                continue
            for i, k in enumerate(('lIns', 'tIns', 'rIns', 'bIns')):
                if b.get(k) is not None:
                    ins[i] = int(b.get(k))
            if b.get('anchor'):
                anchor = b.get('anchor')
            if b.get('wrap'):
                wrap = b.get('wrap') != 'none'
            if b.get('vert'):
                vert = b.get('vert')
            na = b.find(A + 'normAutofit')
            if na is not None:
                fontscale = int(na.get('fontScale', '100000')) / 100000.0
                lnred = int(na.get('lnSpcReduction', '0')) / 100000.0
        default_color = None
        if style is not None:
            fr = style.find(A + 'fontRef')
            if fr is not None:
                default_color = self.ctx.color(first_color(fr))
        jmap = {'l': 'flex-start', 'ctr': 'center', 'r': 'flex-end', 'just': 'flex-start', 'dist': 'flex-start'}
        amap = {'t': 'flex-start', 'ctr': 'center', 'b': 'flex-end'}
        paras = []
        for p in tb.findall(A + 'p'):
            pp = p.find(A + 'pPr')
            lvl = int(pp.get('lvl', '0')) if pp is not None else 0
            d = self.lvl_defaults(ph, inherited, lvl)
            ls0 = tb.find(A + 'lstStyle')
            if ls0 is not None:
                lp = ls0.find(A + 'lvl%dpPr' % (lvl + 1))
                if lp is not None:
                    self.merge_ppr(d, lp)
            if pp is not None:
                self.merge_ppr(d, pp)
            runs = []
            maxsz = 0
            for r in p:
                t = r.tag.split('}')[1]
                if t in ('r', 'fld'):
                    rd = dict(d)
                    rp = r.find(A + 'rPr')
                    if rp is not None:
                        self.merge_rpr(rd, rp)
                    txt = r.findtext(A + 't') or ''
                    sz = rd['sz'] * fontscale * self.fontk
                    maxsz = max(maxsz, sz)
                    col = rd['color'] or default_color or self.ctx.color(ET.fromstring('<a:schemeClr xmlns:a="%s" val="tx1"/>' % NS['a']))
                    css = 'font-size:%.2fpx;color:%s;' % (sz / 100.0 * self.scale / 72.0, col)
                    if rd['b']:
                        css += 'font-weight:bold;'
                    if rd['i']:
                        css += 'font-style:italic;'
                    if rd.get('u'):
                        css += 'text-decoration:underline;'
                    if rd.get('cap') == 'all':
                        txt = txt.upper()
                    runs.append('<span style="%s">%s</span>' % (css, esc(txt)))
                elif t == 'br':
                    runs.append('<br/>')
            if maxsz == 0:
                er = p.find(A + 'endParaRPr')
                ed = dict(d)
                if er is not None:
                    self.merge_rpr(ed, er)
                maxsz = ed['sz'] * fontscale * self.fontk
            lh = (d['lnSpc'] or 1.0) * (1 - lnred) * 1.2
            pcss = 'line-height:%.3f;margin-top:%.2fpx;margin-bottom:%.2fpx;' % (
                lh, d['spcBef'] * self.scale / 72.0, d['spcAft'] * self.scale / 72.0)
            pcss += 'padding-left:%.2fpx;text-indent:%.2fpx;' % (self.px(d['marL'] + 0), self.px(d['indent']))
            pcss += 'text-align:%s;font-size:%.2fpx;' % ({'l': 'left', 'ctr': 'center', 'r': 'right', 'just': 'justify'}.get(d['algn'], 'left'), maxsz / 100.0 * self.scale / 72.0)
            bullet = ''
            if d['bullet'] and runs:
                bch = d['bullet']
                if bch in ('q', 'Ø', '§', 'ü', 'n', 'l'):
                    bch = '▪'
                bullet = '<span style="display:inline-block;width:%.2fpx;text-indent:0">%s</span>' % (
                    max(self.px(-d['indent']), 8), esc(bch))
            content = ''.join(runs) if runs else '&#8203;'
            paras.append('<p style="%s">%s%s</p>' % (pcss, bullet, content))
        css = ('position:absolute;left:0;top:0;width:%.2fpx;height:%.2fpx;display:flex;flex-direction:column;'
               'justify-content:%s;padding:%.2fpx %.2fpx %.2fpx %.2fpx;box-sizing:border-box;%s') % (
            Wd, Ht, amap.get(anchor, 'flex-start'), self.px(ins[1]), self.px(ins[2]), self.px(ins[3]), self.px(ins[0]),
            '' if wrap else 'white-space:nowrap;')
        if vert in ('vert270',):
            css += 'writing-mode:vertical-rl;transform:rotate(180deg);'
        elif vert in ('vert', 'eaVert'):
            css += 'writing-mode:vertical-rl;'
        return '<div class="tx" style="%s">%s</div>' % (css, ''.join(paras))

    def pic(self, el, part, xf):
        b = self.xfrm_of(el, P + 'spPr/' + A + 'xfrm')
        if b is None:
            return ''
        b = self.apply(xf, b)
        bf = el.find(P + 'blipFill')
        uri = self.blip_uri(bf, part)
        X, Y, Wd, Ht = self.px(b['x']), self.px(b['y']), self.px(b['w']), self.px(b['h'])
        if not uri:
            return '<div class="sp" style="left:%.2fpx;top:%.2fpx;width:%.2fpx;height:%.2fpx;background:#eee"></div>' % (X, Y, Wd, Ht)
        sr = bf.find(A + 'srcRect')
        crop = {k: (int(sr.get(k, '0')) / 100000.0 if sr is not None else 0.0) for k in ('l', 't', 'r', 'b')}
        fw = 1 - crop['l'] - crop['r']
        fhh = 1 - crop['t'] - crop['b']
        iw, ih = Wd / fw, Ht / fhh
        return ('<div class="sp" style="left:%.2fpx;top:%.2fpx;width:%.2fpx;height:%.2fpx;overflow:hidden">'
                '<img src="%s" style="position:absolute;left:%.2fpx;top:%.2fpx;width:%.2fpx;height:%.2fpx"/></div>') % (
            X, Y, Wd, Ht, uri, -iw * crop['l'], -ih * crop['t'], iw, ih)

    def frame(self, el, part, xf):
        b = self.xfrm_of(el, P + 'xfrm')
        if b is None:
            return ''
        b = self.apply(xf, b)
        tbl = el.find('.//' + A + 'tbl')
        if tbl is None:
            return ''
        X, Y = self.px(b['x']), self.px(b['y'])
        sx = b['w'] / max(1, sum(int(g.get('w')) for g in tbl.find(A + 'tblGrid')))
        cols = [self.px(int(g.get('w')) * sx) for g in tbl.find(A + 'tblGrid')]
        th = sum(self.px(int(tr.get('h'))) for tr in tbl.findall(A + 'tr'))
        tname = esc(el.find('.//' + P + 'cNvPr').get('name', ''))
        out = ['<table class="tb" data-h="%.1f" data-name="%s" style="left:%.2fpx;top:%.2fpx;width:%.2fpx">' % (th, tname, X, Y, sum(cols)),
               '<colgroup>' + ''.join('<col style="width:%.2fpx"/>' % c for c in cols) + '</colgroup>']
        for tr in tbl.findall(A + 'tr'):
            out.append('<tr style="height:%.2fpx">' % self.px(int(tr.get('h'))))
            ci = 0
            for tc in tr.findall(A + 'tc'):
                if tc.get('hMerge') == '1' or tc.get('vMerge') == '1':
                    ci += 1
                    continue
                span = int(tc.get('gridSpan', '1'))
                rspan = int(tc.get('rowSpan', '1'))
                tcpr = tc.find(A + 'tcPr')
                css = ''
                ins = [91440, 45720, 91440, 45720]
                anchor = 't'
                if tcpr is not None:
                    sf = tcpr.find(A + 'solidFill')
                    if sf is not None:
                        css += 'background:%s;' % self.ctx.color(first_color(sf))
                    for side, tagn in (('left', 'lnL'), ('right', 'lnR'), ('top', 'lnT'), ('bottom', 'lnB')):
                        ln = tcpr.find(A + tagn)
                        if ln is not None and ln.find(A + 'solidFill') is not None:
                            w = int(ln.get('w', '12700')) / 12700.0 * self.scale / 72.0
                            css += 'border-%s:%.2fpx solid %s;' % (side, max(0.5, w), self.ctx.color(first_color(ln.find(A + 'solidFill'))))
                    for i, k in enumerate(('marL', 'marT', 'marR', 'marB')):
                        if tcpr.get(k) is not None:
                            ins[i] = int(tcpr.get(k))
                    anchor = tcpr.get('anchor', 't')
                cw = sum(cols[ci:ci + span])
                txb = tc.find(A + 'txBody')
                inner = ''
                if txb is not None:
                    inner = self.cell_text(txb, ins)
                va = {'t': 'top', 'ctr': 'middle', 'b': 'bottom'}.get(anchor, 'top')
                out.append('<td colspan="%d" rowspan="%d" style="%svertical-align:%s;width:%.2fpx">%s</td>' % (
                    span, rspan, css, va, cw, inner))
                ci += span
            out.append('</tr>')
        out.append('</table>')
        return ''.join(out)

    def cell_text(self, txb, ins):
        paras = []
        for p in txb.findall(A + 'p'):
            d = {'sz': 1800, 'b': False, 'i': False, 'color': '#000', 'algn': 'l', 'font': None,
                 'bullet': None, 'marL': 0, 'indent': 0, 'lnSpc': None, 'spcBef': 0, 'spcAft': 0}
            pp = p.find(A + 'pPr')
            if pp is not None:
                self.merge_ppr(d, pp)
            runs = []
            maxsz = 0
            for r in p:
                t = r.tag.split('}')[1]
                if t in ('r', 'fld'):
                    rd = dict(d)
                    rp = r.find(A + 'rPr')
                    if rp is not None:
                        self.merge_rpr(rd, rp)
                    sz = rd['sz'] * self.fontk
                    maxsz = max(maxsz, sz)
                    css = 'font-size:%.2fpx;color:%s;' % (sz / 100.0 * self.scale / 72.0, rd['color'] or '#000')
                    if rd['b']:
                        css += 'font-weight:bold;'
                    if rd['i']:
                        css += 'font-style:italic;'
                    runs.append('<span style="%s">%s</span>' % (css, esc(r.findtext(A + 't'))))
                elif t == 'br':
                    runs.append('<br/>')
            if not maxsz:
                er = p.find(A + 'endParaRPr')
                if er is not None and er.get('sz'):
                    maxsz = int(er.get('sz')) * self.fontk
                else:
                    maxsz = d['sz'] * self.fontk
            lh = (d['lnSpc'] or 1.0) * 1.2
            paras.append('<p style="line-height:%.3f;text-align:%s;font-size:%.2fpx;margin:%.2fpx 0 %.2fpx 0">%s</p>' % (
                lh, {'l': 'left', 'ctr': 'center', 'r': 'right'}.get(d['algn'], 'left'),
                maxsz / 100.0 * self.scale / 72.0, d['spcBef'] * self.scale / 72.0, d['spcAft'] * self.scale / 72.0,
                ''.join(runs) or '&#8203;'))
        return '<div style="padding:%.2fpx %.2fpx %.2fpx %.2fpx">%s</div>' % (
            self.px(ins[1]), self.px(ins[2]), self.px(ins[3]), self.px(ins[0]), ''.join(paras))


CSS = """
*{margin:0;padding:0}
body{background:#888}
.slide{position:relative;overflow:hidden;font-family:'Carlito','Calibri','Noto Sans',sans-serif}
.sp{position:absolute;transform-origin:center center}
.sp svg.g{position:absolute;left:0;top:0}
.tx{overflow:visible;word-wrap:break-word}
.tx p{margin:0}
table.tb{position:absolute;border-collapse:collapse;table-layout:fixed}
table.tb td{overflow:hidden;word-wrap:break-word}
table.tb p{margin:0}
"""


def main():
    src, outdir = sys.argv[1], sys.argv[2]
    which = None
    scale = 96.0
    fontk = 1.0
    args = sys.argv[3:]
    i = 0
    while i < len(args):
        if args[i] == '--scale':
            scale = float(args[i + 1]); i += 2; continue
        if args[i] == '--fontk':
            fontk = float(args[i + 1]); i += 2; continue
        which = [int(x) for x in args[i].split(',')]
        i += 1
    os.makedirs(outdir, exist_ok=True)
    rd = Renderer(src, scale=scale, fontk=fontk)
    idxs = which or list(range(1, len(rd.slides) + 1))
    for n in idxs:
        body, W, H = rd.render_slide(n - 1)
        with open(os.path.join(outdir, 'slide%02d.html' % n), 'w') as f:
            f.write('<!doctype html><html><head><meta charset="utf-8"><style>%s</style></head><body>%s</body></html>' % (CSS, body))
    print('rendered', len(idxs), 'slides', '%dx%d' % (W, H))


if __name__ == '__main__':
    main()
