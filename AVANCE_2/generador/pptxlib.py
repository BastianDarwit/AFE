"""Constructor de PPTX por OOXML directo (sin dependencias externas).

Reutiliza los patrones, temas, layouts y medios de un .pptx base para conservar
la identidad institucional, y reemplaza por completo el conjunto de laminas.
"""
import re
import shutil
import zipfile

EMU = 914400
SLIDE_W = 10.0
SLIDE_H = 7.5

# --- Paleta institucional (tomada del tema del pptx base) ---
BLUE = '1F497D'        # dk2 - azul institucional
BLUE_LT = '4F81BD'     # accent1
BLUE_PALE = 'DCE6F1'   # relleno claro
BLUE_PALE2 = 'EAF0F8'
RED = 'C0504D'         # accent2
RED_PALE = 'F7E4E3'
GREEN = '77933C'       # accent3 oscurecido
GREEN_PALE = 'EAF1DD'
TEAL = '31859B'        # accent5 oscurecido
TEAL_PALE = 'DEEAEF'
ORANGE = 'E36C0A'      # accent6 oscurecido
ORANGE_PALE = 'FDE9D9'
GRAY_TX = '404040'
GRAY_MID = '808080'
GRAY_LT = 'F2F2F2'
GRAY_BD = 'BFBFBF'
WHITE = 'FFFFFF'
FONT = 'Calibri'

LAYOUT_COVER = 'slideLayout1.xml'    # master1 - fondo portada
LAYOUT_BODY = 'slideLayout10.xml'    # master2 - fondo contenido
LAYOUT_SECTION = 'slideLayout38.xml'  # master4 - fondo seccion (oscuro)
LAYOUT_BLANK = 'slideLayout7.xml'    # master1 - fondo portada, sin marcadores


def e(v):
    """Pulgadas a EMU."""
    return int(round(v * EMU))


def esc(t):
    return (t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            .replace('"', '&quot;'))


# ---------------------------------------------------------------- texto
def run(text, sz=1100, b=False, i=False, color=GRAY_TX, font=FONT, spc=None,
        caps=False, u=False):
    """Un <a:r>. sz en centesimas de punto (1100 = 11 pt)."""
    at = ' lang="es-CL" sz="%d" dirty="0"' % sz
    if b:
        at += ' b="1"'
    if i:
        at += ' i="1"'
    if u:
        at += ' u="sng"'
    if spc is not None:
        at += ' spc="%d"' % spc
    if caps:
        at += ' cap="all"'
    return ('<a:r><a:rPr%s><a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
            '<a:latin typeface="%s"/><a:cs typeface="%s"/></a:rPr>'
            '<a:t>%s</a:t></a:r>' % (at, color, font, font, esc(text)))


def para(runs, algn='l', space_before=0, space_after=0, line=100, bullet=None,
         indent=None, marL=None):
    """Un <a:p>. space_* en puntos; line en % (100 = simple)."""
    if isinstance(runs, str):
        runs = [run(runs)]
    props = ' algn="%s"' % algn
    if marL is not None:
        props += ' marL="%d"' % int(marL * EMU)
    if indent is not None:
        props += ' indent="%d"' % int(indent * EMU)
    inner = ''
    if line != 100:
        inner += '<a:lnSpc><a:spcPct val="%d"/></a:lnSpc>' % (line * 1000)
    if space_before:
        inner += '<a:spcBef><a:spcPts val="%d"/></a:spcBef>' % int(space_before * 100)
    if space_after:
        inner += '<a:spcAft><a:spcPts val="%d"/></a:spcAft>' % int(space_after * 100)
    if bullet is None:
        inner += '<a:buNone/>'
    else:
        inner += ('<a:buClr><a:srgbClr val="%s"/></a:buClr>'
                  '<a:buFont typeface="Arial" pitchFamily="34" charset="0"/>'
                  '<a:buChar char="%s"/>' % (bullet[1], bullet[0]))
    return '<a:pPr%s>%s</a:pPr>%s' % (props, inner, ''.join(runs)) and \
           '<a:p><a:pPr%s>%s</a:pPr>%s</a:p>' % (props, inner, ''.join(runs))


def body_pr(anchor='t', wrap=True, ins=(0.06, 0.05, 0.06, 0.05), autofit=True,
            anchor_ctr=False):
    """<a:bodyPr>. ins = (left, top, right, bottom) en pulgadas."""
    l, t, r, b = ins
    at = ' wrap="%s" lIns="%d" tIns="%d" rIns="%d" bIns="%d" anchor="%s"' % (
        'square' if wrap else 'none', e(l), e(t), e(r), e(b), anchor)
    if anchor_ctr:
        at += ' anchorCtr="1"'
    return '<a:bodyPr%s>%s</a:bodyPr>' % (at, '<a:normAutofit/>' if autofit else '')


# ---------------------------------------------------------------- formas
class Shape:
    """Registro de una forma, con su geometria para poder validarla."""

    def __init__(self, xml, x, y, w, h, kind='sp', name='', text='',
                 text_specs=None, container=False):
        self.xml = xml
        self.x, self.y, self.w, self.h = x, y, w, h
        self.kind = kind
        self.name = name
        self.text = text
        self.text_specs = text_specs or []
        self.container = container


def _xfrm(x, y, w, h, rot=None, flip=None):
    at = ''
    if rot:
        at += ' rot="%d"' % int(rot * 60000)
    if flip:
        at += ' flipH="1"' if flip == 'h' else ' flipV="1"'
    return ('<a:xfrm%s><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            % (at, e(x), e(y), e(w), e(h)))


def _fill(color, alpha=None):
    if color is None:
        return '<a:noFill/>'
    if alpha is not None:
        return ('<a:solidFill><a:srgbClr val="%s"><a:alpha val="%d"/>'
                '</a:srgbClr></a:solidFill>' % (color, int(alpha * 1000)))
    return '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>' % color


def _line(color, width=1.0, dash=None):
    if color is None:
        return '<a:ln><a:noFill/></a:ln>'
    d = '<a:prstDash val="%s"/>' % dash if dash else ''
    return ('<a:ln w="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill>%s'
            '</a:ln>' % (int(width * 12700), color, d))


def shape(sid, name, x, y, w, h, geom='rect', fill=BLUE_PALE, line=None,
          line_w=1.0, dash=None, paras=None, anchor='t',
          ins=(0.10, 0.07, 0.10, 0.07), adj=None, rot=None, shadow=False,
          container=False, autofit=True):
    """Forma con geometria preestablecida y texto opcional."""
    paras = paras or []
    av = ''
    if adj:
        av = ''.join('<a:gd name="adj%s" fmla="val %d"/>'
                     % (k, v) for k, v in adj)
    txt = ''
    if paras:
        txt = ('<p:txBody>%s<a:lstStyle/>%s</p:txBody>'
               % (body_pr(anchor=anchor, ins=ins, autofit=autofit), ''.join(paras)))
    else:
        txt = ('<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="es-CL"/>'
               '</a:p></p:txBody>')
    sh = ''
    if shadow:
        sh = ('<a:effectLst><a:outerShdw blurRad="38100" dist="25400" dir="5400000"'
              ' algn="t" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="18000"/>'
              '</a:srgbClr></a:outerShdw></a:effectLst>')
    xml = ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/><p:nvPr/>'
           '</p:nvSpPr><p:spPr>%s<a:prstGeom prst="%s"><a:avLst>%s</a:avLst>'
           '</a:prstGeom>%s%s%s</p:spPr>%s</p:sp>'
           % (sid, esc(name), _xfrm(x, y, w, h, rot), geom, av,
              _fill(fill), _line(line, line_w, dash), sh, txt))
    return Shape(xml, x, y, w, h, 'sp', name, container=container)


def textbox(sid, name, x, y, w, h, paras, anchor='t', wrap=True,
            ins=(0.05, 0.03, 0.05, 0.03), autofit=True):
    """Cuadro de texto sin relleno ni borde."""
    xml = ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/>'
           '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>%s'
           '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
           '<p:txBody>%s<a:lstStyle/>%s</p:txBody></p:sp>'
           % (sid, esc(name), _xfrm(x, y, w, h),
              body_pr(anchor=anchor, wrap=wrap, ins=ins, autofit=autofit),
              ''.join(paras)))
    return Shape(xml, x, y, w, h, 'sp', name)


def _cell_line(tag, spec):
    """<a:lnT>/<a:lnB>. spec = None (sin linea) o (color, ancho_pt)."""
    if spec is None:
        return '<a:%s><a:noFill/></a:%s>' % (tag, tag)
    color, w = spec
    return ('<a:%s w="%d" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>'
            '<a:srgbClr val="%s"/></a:solidFill><a:prstDash val="solid"/>'
            '</a:%s>' % (tag, int(w * 12700), color, tag))


def table(sid, name, x, y, col_w, rows, row_h, first_row=False):
    """Tabla nativa. col_w en pulgadas; rows = [[celda,...],...].

    Cada celda: dict(t=texto|runs=list, sz, b, i, color, fill, algn, anchor,
                     lnT=(color,pt)|None, lnB=(color,pt)|None)
    Sin relleno y sin lineas verticales por defecto, conforme a normas APA.
    """
    grid = ''.join('<a:gridCol w="%d"/>' % e(cw) for cw in col_w)
    trs = []
    for ri, r in enumerate(rows):
        tcs = []
        for cell in r:
            if cell is None:
                cell = {}
            runs = cell.get('runs')
            if runs is None:
                runs = [run(cell.get('t', ''), sz=cell.get('sz', 1100),
                            b=cell.get('b', False),
                            color=cell.get('color', GRAY_TX),
                            i=cell.get('i', False))]
            p = para(runs, algn=cell.get('algn', 'l'), line=cell.get('line', 100))
            tcpr = ('<a:tcPr marL="%d" marR="%d" marT="%d" marB="%d" anchor="%s">'
                    % (e(cell.get('marL', 0.06)), e(cell.get('marR', 0.06)),
                       e(cell.get('marT', 0.03)), e(cell.get('marB', 0.03)),
                       cell.get('anchor', 't')))
            tcpr += _cell_line('lnL', cell.get('lnL'))
            tcpr += _cell_line('lnR', cell.get('lnR'))
            tcpr += _cell_line('lnT', cell.get('lnT'))
            tcpr += _cell_line('lnB', cell.get('lnB'))
            tcpr += _fill(cell.get('fill'))
            tcpr += '</a:tcPr>'
            tcs.append('<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>%s</a:txBody>%s</a:tc>'
                       % (p, tcpr))
        trs.append('<a:tr h="%d">%s</a:tr>' % (e(row_h[ri]), ''.join(tcs)))
    total_w = sum(col_w)
    total_h = sum(row_h)
    xml = ('<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="%d" name="%s"/>'
           '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/>'
           '</p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
           '<p:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></p:xfrm>'
           '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/'
           'drawingml/2006/table"><a:tbl><a:tblPr firstRow="%d" bandRow="0">'
           '<a:tableStyleId>{5940675A-B579-460E-94D1-54222C63F5DA}</a:tableStyleId>'
           '</a:tblPr><a:tblGrid>%s</a:tblGrid>%s</a:tbl></a:graphicData>'
           '</a:graphic></p:graphicFrame>'
           % (sid, esc(name), e(x), e(y), e(total_w), e(total_h),
              1 if first_row else 0, grid, ''.join(trs)))
    return Shape(xml, x, y, total_w, total_h, 'tbl', name)


def picture(sid, name, x, y, w, h, rid, crop=None):
    src = ''
    if crop:
        l, t, r, b = crop
        src = ('<a:srcRect l="%d" t="%d" r="%d" b="%d"/>'
               % (int(l * 1000), int(t * 1000), int(r * 1000), int(b * 1000)))
    xml = ('<p:pic><p:nvPicPr><p:cNvPr id="%d" name="%s"/><p:cNvPicPr>'
           '<a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
           '<p:blipFill><a:blip r:embed="%s"/>%s<a:stretch><a:fillRect/>'
           '</a:stretch></p:blipFill><p:spPr>%s<a:prstGeom prst="rect"><a:avLst/>'
           '</a:prstGeom></p:spPr></p:pic>'
           % (sid, esc(name), rid, src, _xfrm(x, y, w, h)))
    return Shape(xml, x, y, w, h, 'pic', name)


# ---------------------------------------------------------------- lamina
class Slide:
    def __init__(self, layout, title=None, hidden=False, notes='', label=''):
        self.layout = layout
        self.shapes = []
        self.hidden = hidden
        self.notes = notes
        self.label = label or title or ''
        self.title = title
        self._sid = 1
        self.rels = []   # (rId, type, target)

    def nid(self):
        self._sid += 1
        return self._sid

    def add(self, sh):
        self.shapes.append(sh)
        return sh

    def xml(self):
        show = '' if not self.hidden else ' show="0"'
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
                ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships" xmlns:p="http://schemas.openxmlformats.org/'
                'presentationml/2006/main"%s><p:cSld><p:spTree><p:nvGrpSpPr>'
                '<p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
                '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
                '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
                '%s</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/>'
                '</p:clrMapOvr></p:sld>'
                % (show, ''.join(s.xml for s in self.shapes)))

    def notes_xml(self):
        paras = []
        for block in self.notes.strip().split('\n'):
            block = block.strip()
            paras.append('<a:p><a:pPr/>%s</a:p>'
                         % (run(block, sz=1200, color='000000') if block else ''))
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                '<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/'
                'main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships" xmlns:p="http://schemas.openxmlformats.org/'
                'presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr>'
                '<p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
                '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
                '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
                '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Marcador de imagen"/>'
                '<p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/>'
                '</p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr>'
                '<p:spPr/></p:sp>'
                '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Marcador de notas"/>'
                '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>'
                '<p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>'
                '<p:txBody><a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/>%s'
                '</p:txBody></p:sp>'
                '<p:sp><p:nvSpPr><p:cNvPr id="4" name="Marcador de numero"/>'
                '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>'
                '<p:ph type="sldNum" sz="quarter" idx="10"/></p:nvPr></p:nvSpPr>'
                '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p>'
                '<a:fld id="{5A8FB45B-0388-4BF1-AE79-E6A2536ED20E}" type="slidenum">'
                '<a:rPr lang="es-CL" smtClean="0"/><a:t>1</a:t></a:fld>'
                '<a:endParaRPr lang="es-CL"/></a:p></p:txBody></p:sp>'
                '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/>'
                '</p:clrMapOvr></p:notes>' % ''.join(paras))


# ---------------------------------------------------------------- empaquetado
SLIDE_CT = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
NOTES_CT = ('application/vnd.openxmlformats-officedocument.presentationml.'
            'notesSlide+xml')
REL_SLIDE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
REL_LAYOUT = ('http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
              'slideLayout')
REL_NOTES = ('http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
             'notesSlide')
REL_NOTESMASTER = ('http://schemas.openxmlformats.org/officeDocument/2006/'
                   'relationships/notesMaster')
REL_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'


def build(base_pptx, out_pptx, slides, title='', author='', media=None):
    """Escribe out_pptx reutilizando patrones/temas/medios de base_pptx."""
    zin = zipfile.ZipFile(base_pptx)
    keep = []
    for n in zin.namelist():
        if re.match(r'ppt/slides/', n) or re.match(r'ppt/notesSlides/', n):
            continue
        if n in ('[Content_Types].xml', 'ppt/presentation.xml',
                 'ppt/_rels/presentation.xml.rels', 'docProps/app.xml',
                 'docProps/core.xml', 'docProps/thumbnail.jpeg'):
            continue
        keep.append(n)

    # --- Content_Types ---
    ct = zin.read('[Content_Types].xml').decode('utf-8')
    ct = re.sub(r'<Override PartName="/ppt/slides/slide\d+\.xml"[^>]*/>', '', ct)
    ct = re.sub(r'<Override PartName="/ppt/notesSlides/notesSlide\d+\.xml"[^>]*/>',
                '', ct)
    ct = re.sub(r'<Override PartName="/docProps/thumbnail\.jpeg"[^>]*/>', '', ct)
    ct = ct.replace('<Default Extension="jpeg" ContentType="image/jpeg"/>',
                    '<Default Extension="jpeg" ContentType="image/jpeg"/>')
    adds = []
    for i, s in enumerate(slides, 1):
        adds.append('<Override PartName="/ppt/slides/slide%d.xml" ContentType="%s"/>'
                    % (i, SLIDE_CT))
        if s.notes.strip():
            adds.append('<Override PartName="/ppt/notesSlides/notesSlide%d.xml"'
                        ' ContentType="%s"/>' % (i, NOTES_CT))
    if media and 'Extension="png"' not in ct:
        adds.append('<Default Extension="png" ContentType="image/png"/>')
    ct = ct.replace('</Types>', ''.join(adds) + '</Types>')

    # --- presentation.xml + rels ---
    pres = zin.read('ppt/presentation.xml').decode('utf-8')
    prels = zin.read('ppt/_rels/presentation.xml.rels').decode('utf-8')
    prels = re.sub(r'<Relationship[^>]*Target="slides/slide\d+\.xml"[^>]*/>', '',
                   prels)
    slide_rids = []
    new_rels = []
    for i in range(1, len(slides) + 1):
        rid = 'rId%d' % (500 + i)
        slide_rids.append(rid)
        new_rels.append('<Relationship Id="%s" Type="%s" Target="slides/slide%d.xml"/>'
                        % (rid, REL_SLIDE, i))
    prels = prels.replace('</Relationships>', ''.join(new_rels) + '</Relationships>')
    sldid = ''.join('<p:sldId id="%d" r:id="%s"/>' % (256 + i, rid)
                    for i, rid in enumerate(slide_rids))
    pres = re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>', '<p:sldIdLst>%s</p:sldIdLst>'
                  % sldid, pres, flags=re.S)

    zout = zipfile.ZipFile(out_pptx, 'w', zipfile.ZIP_DEFLATED)
    for n in keep:
        data = zin.read(n)
        if n == '_rels/.rels':
            data = re.sub(rb'<Relationship[^>]*Target="docProps/thumbnail\.jpeg"[^>]*/>', b'', data)
        zout.writestr(n, data)
    zout.writestr('[Content_Types].xml', ct)
    zout.writestr('ppt/presentation.xml', pres)
    zout.writestr('ppt/_rels/presentation.xml.rels', prels)
    zout.writestr('docProps/core.xml',
                  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                  '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/'
                  'package/2006/metadata/core-properties" xmlns:dc="http://purl.org/'
                  'dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
                  'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://'
                  'www.w3.org/2001/XMLSchema-instance"><dc:title>%s</dc:title>'
                  '<dc:creator>%s</dc:creator><cp:lastModifiedBy>%s</cp:lastModifiedBy>'
                  '<cp:revision>1</cp:revision></cp:coreProperties>'
                  % (esc(title), esc(author), esc(author)))
    zout.writestr('docProps/app.xml',
                  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                  '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument'
                  '/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats'
                  '.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft '
                  'Office PowerPoint</Application><Slides>%d</Slides>'
                  '<Company></Company></Properties>' % len(slides))

    for mname, data in (media or {}).items():
        zout.writestr('ppt/media/' + mname, data)
    for i, s in enumerate(slides, 1):
        zout.writestr('ppt/slides/slide%d.xml' % i, s.xml())
        rels = ['<Relationship Id="rId1" Type="%s" Target="../slideLayouts/%s"/>'
                % (REL_LAYOUT, s.layout)]
        nxt = 2
        if s.notes.strip():
            rels.append('<Relationship Id="rId%d" Type="%s" '
                        'Target="../notesSlides/notesSlide%d.xml"/>'
                        % (nxt, REL_NOTES, i))
            nxt += 1
        for rid, rtype, target in s.rels:
            rels.append('<Relationship Id="%s" Type="%s" Target="%s"/>'
                        % (rid, rtype, target))
        zout.writestr('ppt/slides/_rels/slide%d.xml.rels' % i,
                      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                      '<Relationships xmlns="http://schemas.openxmlformats.org/'
                      'package/2006/relationships">%s</Relationships>' % ''.join(rels))
        if s.notes.strip():
            zout.writestr('ppt/notesSlides/notesSlide%d.xml' % i, s.notes_xml())
            zout.writestr(
                'ppt/notesSlides/_rels/notesSlide%d.xml.rels' % i,
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
                '2006/relationships"><Relationship Id="rId1" Type="%s" '
                'Target="../notesMasters/notesMaster1.xml"/><Relationship Id="rId2" '
                'Type="%s" Target="../slides/slide%d.xml"/></Relationships>'
                % (REL_NOTESMASTER, REL_SLIDE, i))
    zout.close()
    zin.close()
