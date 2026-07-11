"""Compare page breaks and cover image relationships."""
import zipfile
import re
from docx import Document
from docx.oxml.ns import qn

def page_breaks(path):
    z = zipfile.ZipFile(path)
    xml = z.read("word/document.xml").decode("utf-8", "ignore")
    # page breaks
    pb = len(re.findall(r'w:type="page"', xml))
    sect = xml.count("w:sectPr")
    print(f"{path}: pageBreaks={pb} sectPr={sect}")
    # find context around EXECUTIVE SUMMARY
    plain_parts = []
    # look for page break near executive
    idx = xml.find("EXECUTIVE SUMMARY")
    if idx > 0:
        chunk = xml[max(0, idx - 800): idx + 200]
        print("  before EXECUTIVE has page break:", 'w:type="page"' in chunk)
        print("  before EXECUTIVE has lastRenderedPageBreak:", "lastRenderedPageBreak" in chunk)
        # show nearby tags of interest
        tags = re.findall(r"<(w:br[^>]*|w:lastRenderedPageBreak[^/]*|w:p |/w:p|w:txbxContent)", chunk)
        print("  nearby tags:", tags[-15:])

def cover_images(path):
    print("\nCover images:", path)
    z = zipfile.ZipFile(path)
    xml = z.read("word/document.xml").decode("utf-8", "ignore")
    # Find drawings with main_img or near PREPARED / cover
    # Get all a:blip embed rIds near descr main/aerial/subject or empty after our clear
    # Parse docPr name/descr + following blip
    pattern = re.compile(
        r'<wp:docPr[^>]*name="([^"]*)"[^>]*(?:descr="([^"]*)")?[^/]*/>.{0,400}?<a:blip[^>]*r:embed="([^"]+)"',
        re.DOTALL,
    )
    # simpler: all blips with preceding descr within 500 chars
    for m in re.finditer(r'descr="([^"]*)"', xml):
        descr = m.group(1)
        window = xml[m.start(): m.start() + 600]
        embeds = re.findall(r'r:embed="([^"]+)"', window)
        if embeds:
            print(f"  descr={descr!r} -> {embeds[0]}")

    # Check first few media file headers
    for name in sorted(n for n in z.namelist() if n.startswith("word/media/"))[:8]:
        data = z.read(name)[:16]
        kind = "unknown"
        if data[:3] == b"\xff\xd8\xff":
            kind = "jpeg"
        elif data[:8] == b"\x89PNG\r\n\x1a\n":
            kind = "png"
        elif data[:4] == b"GIF8":
            kind = "gif"
        print(f"  {name}: {kind} size={z.getinfo(name).file_size}")

for p in [
    "BOV Report (CLIENT) (1).docx",
    "bov_template.docx",
    r"property_reports\BOV_River_Walk_Retail_Center_2026-07-11_2029.docx",
]:
    page_breaks(p)
cover_images(r"property_reports\BOV_River_Walk_Retail_Center_2026-07-11_2029.docx")
cover_images("bov_template.docx")

# Did collapse_empty remove page-break paragraphs?
print("\nEmpty para / break counts in generated:")
from docx import Document
doc = Document(r"property_reports\BOV_River_Walk_Retail_Center_2026-07-11_2029.docx")
breaks = 0
for p in doc.paragraphs:
    xml = p._element.xml
    if 'w:type="page"' in xml:
        breaks += 1
        print("  page break para text=", repr(p.text[:40]), " before next content?")
print("total page-break paragraphs:", breaks)
