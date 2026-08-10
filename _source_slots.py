from docx import Document
from docx.oxml.ns import qn
EMU=914400
for name in ['New Template (BOV) - Short Form.docx']:
    d=Document(name)
    print('===',name)
    for drawing in d.element.body.iter(qn('w:drawing')):
        docPr=next(drawing.iter(qn('wp:docPr')),None)
        nm=(docPr.get('name') if docPr is not None else '') or ''
        rid=None
        for blip in drawing.iter(qn('a:blip')):
            rid=blip.get(qn('r:embed')); break
        if not rid or rid not in d.part.related_parts:
            continue
        size=len(d.part.related_parts[rid].blob)
        descr=' '.join((el.get('descr') or '') for tag in (qn('wp:docPr'), qn('pic:cNvPr')) for el in drawing.iter(tag))
        posV=next(drawing.iter(qn('wp:positionV')),None)
        y=None
        if posV is not None:
            off=posV.find(qn('wp:posOffset'))
            if off is not None and off.text:
                y=round(int(off.text)/EMU,3)
        ext=next(drawing.iter(qn('wp:extent')),None)
        cx=cy=None
        if ext is not None:
            cx=round(int(ext.get('cx'))/EMU,3) if ext.get('cx') else None
            cy=round(int(ext.get('cy'))/EMU,3) if ext.get('cy') else None
        print(nm, 'y=',y,'cx=',cx,'cy=',cy,'size=',size,'descr=',repr(descr[:80]))
