from docx import Document
doc = Document('bov_template.docx')
for ti in [0, 1]:
    t = doc.tables[ti]
    print('=== Table', ti, '===')
    for ri, row in enumerate(t.rows):
        print(ri, [c.text[:50] for c in row.cells])
