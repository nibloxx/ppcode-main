"""Patch bov_template.docx: replace remaining static Dallas/Utah sample text with placeholders."""
from docx import Document
from templatize_bov import set_runs_text, OUTPUT

# (paragraph_index, new_text) — exact replacement for single paragraphs
PARA_MAP = {
    9: "{{executive_summary}}",
    24: "{{regional_analysis}}",
    63: "{{vacancy_rates}}",
    69: "{{lease_rates}}",
    71: "{{construction_activity}}",
    73: "{{market_trends}}",
    77: "{{investment_insights}}",
    80: "{{market_recommendations}}",
    85: "{{market_data_sources}}",
    100: "{{reconciliation_summary}}",
    104: "{{sales_conclusion}}",
}

# Paragraph indices to clear (duplicate bullets / old static blocks)
CLEAR_PARAS = {64, 65, 66, 67, 74, 75, 78, 81, 82, 84, 86, 87, 88, 89, 90, 93, 94, 95, 96, 97, 98}


def fix_template(path: str = OUTPUT):
    doc = Document(path)
    for idx, text in PARA_MAP.items():
        if idx < len(doc.paragraphs):
            set_runs_text(doc.paragraphs[idx], text)
    for idx in CLEAR_PARAS:
        if idx < len(doc.paragraphs):
            set_runs_text(doc.paragraphs[idx], "")
            # Remove embedded drawings that cause blank comp pages
            p = doc.paragraphs[idx]._element
            for drawing in list(p.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing")):
                drawing.getparent().remove(drawing)

    # Update employment table title to current-era framing
    if len(doc.tables) > 9:
        t = doc.tables[9]
        if t.rows:
            set_runs_text(t.rows[0].cells[0].paragraphs[0],
                          "EMPLOYMENT & UNEMPLOYMENT STATISTICS (Recent)")

    doc.save(path)
    print(f"Patched {path}: {len(PARA_MAP)} placeholders, cleared {len(CLEAR_PARAS)} static/blank blocks")


if __name__ == "__main__":
    fix_template(OUTPUT)
    fix_template("bov_prospect_template.docx")
