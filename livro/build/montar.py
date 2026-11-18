"""Gera o reference.docx (A5, tipografia de livro) e compila o manuscrito em .docx.
Uso: python3 montar.py  (a partir da pasta livro/build)"""
import glob, os, re, subprocess
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

AQUI = os.path.dirname(os.path.abspath(__file__))
LIVRO = os.path.dirname(AQUI)
MARINHO = RGBColor(0x1B, 0x2A, 0x41)

ref = Document(os.path.join(AQUI, "ref-base.docx"))
for s in ref.sections:
    s.page_width, s.page_height = Cm(14.8), Cm(21.0)
    s.left_margin = s.right_margin = Cm(1.7)
    s.top_margin, s.bottom_margin = Cm(1.8), Cm(1.8)
st = ref.styles
def fmt(nome, fonte, tam, cor=None, negrito=None, antes=0, depois=6, alinh=None, entre=None):
    if nome not in [x.name for x in st]:
        return
    f = next(x for x in st if x.name == nome)
    f.font.name, f.font.size = fonte, Pt(tam)
    rpr = f.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is not None:
        for a in list(rf.attrib):
            if a.endswith("Theme"):
                del rf.attrib[a]
        for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            rf.set(qn(a), fonte)
    if cor: f.font.color.rgb = cor
    if negrito is not None: f.font.bold = negrito
    pf = f.paragraph_format
    pf.space_before, pf.space_after = Pt(antes), Pt(depois)
    if alinh is not None: pf.alignment = alinh
    if entre: pf.line_spacing = entre
for nome in ["Normal", "Body Text", "First Paragraph", "Compact"]:
    fmt(nome, "Times New Roman", 10.5, depois=3, alinh=WD_ALIGN_PARAGRAPH.JUSTIFY, entre=1.1)
fmt("Title", "Times New Roman", 24, MARINHO, True, 0, 12, WD_ALIGN_PARAGRAPH.CENTER)
fmt("Subtitle", "Times New Roman", 13, MARINHO, False, 0, 12, WD_ALIGN_PARAGRAPH.CENTER)
fmt("Author", "Times New Roman", 12, MARINHO, False, 24, 6, WD_ALIGN_PARAGRAPH.CENTER)
fmt("Heading 1", "Times New Roman", 20, MARINHO, True, 0, 12, WD_ALIGN_PARAGRAPH.CENTER)
fmt("Heading 2", "Times New Roman", 16, MARINHO, True, 6, 14, WD_ALIGN_PARAGRAPH.CENTER)
fmt("Heading 3", "Times New Roman", 12, MARINHO, True, 14, 6)
fmt("Block Text", "Times New Roman", 10.5, MARINHO, False, 6, 12, WD_ALIGN_PARAGRAPH.CENTER)
next(x for x in st if x.name == "Heading 1").paragraph_format.page_break_before = True
ref.save(os.path.join(AQUI, "reference.docx"))

partes = [os.path.join(LIVRO, "00-capa.md")] + sorted(glob.glob(os.path.join(LIVRO, "capitulos", "cap-*.md")))
partes = [p for p in partes if os.path.exists(p)]
saida = os.path.join(LIVRO, "Meu-CNPJ-precisa-de-um-CPF-manuscrito.docx")
subprocess.run(["pandoc", *partes, "-f", "markdown-implicit_figures", "-o", saida,
                "--reference-doc", os.path.join(AQUI, "reference.docx"), "--toc", "--toc-depth=2",
                "-M", "toc-title=Sumário", "-M", "lang=pt-BR"], check=True)
total = sum(len(open(p, encoding="utf-8").read().split()) for p in partes)
print(f"{saida}\n{total} palavras ≈ {total/300:.0f} páginas de 300 palavras")
