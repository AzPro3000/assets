#!/usr/bin/env python3
"""doc-informe · Extrae el contenido de un .docx (Word o Google Docs exportado) al formato contenido.txt.

Uso:
    python3 docx_a_contenido.py original.docx contenido.txt

No interpreta ni corrige nada: copia texto, negritas, títulos, listas y tablas.
Deja las líneas @ de apertura con lo que pueda deducir para que Claude las revise.
Avisa de lo que el formato no recoge (imágenes, notas al pie, cursivas, enlaces).
"""
import re, sys
from docx import Document
from docx.oxml.ns import qn

SRC, OUT = sys.argv[1], sys.argv[2]
d = Document(SRC)
avisos = []

def esc(t):
    return t.replace("|", "\\|")

def runs_text(p_el, bold_ok=True):
    """Texto de un párrafo con **negrita**. Los saltos de línea manuales pasan a \\n."""
    parts = []
    for r in p_el.iter(qn("w:r")):
        rpr = r.find(qn("w:rPr"))
        b = rpr is not None and rpr.find(qn("w:b")) is not None and rpr.find(qn("w:b")).get(qn("w:val")) not in ("0", "false")
        if rpr is not None and rpr.find(qn("w:i")) is not None and rpr.find(qn("w:i")).get(qn("w:val")) not in ("0", "false"):
            avisos.append("cursiva")
        txt = ""
        for e in r:
            tag = e.tag.split("}")[1]
            if tag == "t": txt += e.text or ""
            elif tag == "tab": txt += "\t"
            elif tag in ("br", "cr"): txt += "\\n"
            elif tag == "drawing": avisos.append("imagen")
            elif tag in ("footnoteReference", "endnoteReference"): avisos.append("nota al pie")
        parts.append((txt, b and bold_ok and txt.strip() != ""))
    # fusiona tramos contiguos con el mismo estado de negrita
    out, cur, curb = [], "", None
    for t, b in parts:
        if curb is None or b == curb: cur += t; curb = b if curb is None else curb
        else: out.append((cur, curb)); cur, curb = t, b
    if cur: out.append((cur, curb))
    s = ""
    for t, b in out:
        if b:
            lead = t[:len(t) - len(t.lstrip())]; trail = t[len(t.rstrip()):]
            s += lead + "**" + t.strip() + "**" + trail
        else: s += t
    if p_el.find(".//" + qn("w:hyperlink")) is not None: avisos.append("enlace")
    return s

def style_of(p_el):
    ps = p_el.find(qn("w:pPr") + "/" + qn("w:pStyle"))
    sid = ps.get(qn("w:val")) if ps is not None else ""
    try: name = d.styles.get_by_id(sid).name if sid else ""
    except Exception: name = sid
    return (sid + " " + (name or "")).lower()

def is_list(p_el):
    return p_el.find(qn("w:pPr") + "/" + qn("w:numPr")) is not None

lines, titulo, first_h1_seen, meta_done = [], None, False, False
for blk in d.element.body:
    tag = blk.tag.split("}")[1]
    if tag == "sdt":     # índice manual (control de contenido): se sustituye; no es contenido
        avisos.append("índice manual (no se copia; si lo quieres, dilo)"); continue
    if tag == "p":
        txt = runs_text(blk)
        sty = style_of(blk)
        if not txt.strip():
            if blk.find(".//" + qn("w:drawing")) is not None: avisos.append("imagen")
            continue
        plain = txt.replace("**", "").strip()
        if "title" in sty and titulo is None:
            titulo = plain; continue
        if re.search(r"heading ?1\b|heading1\b|título 1|titulo1", sty):
            lines.append("# " + plain); first_h1_seen = True
        elif re.search(r"heading ?[2-6]\b|heading[2-6]\b|título [2-6]", sty):
            lines.append("## " + plain)
        elif is_list(blk) or re.search(r"list|lista|viñeta", sty):
            lines.append("- " + txt.strip())
        else:
            lines.append(txt.strip())
    elif tag == "tbl":
        rows = []
        for tr in blk.iter(qn("w:tr")):
            cells = []
            for tc in tr.iter(qn("w:tc")):
                ps = [runs_text(p) for p in tc.iter(qn("w:p"))]
                cells.append(esc("\\n".join(x.strip() for x in ps if x.strip())))
            rows.append(cells)
        rows = [r for r in rows if any(c.strip() for c in r)]
        if not rows: continue
        ncol = max(len(r) for r in rows)
        if not first_h1_seen and ncol == 2 and not meta_done:
            for k, v in rows:   # ficha inicial: etiqueta | valor
                lines.append("META|%s|%s" % (k.replace("**", ""), v))
            meta_done = True
        else:
            lines.append("TABLE")
            # la primera fila se trata como cabecera; si no lo es, cambiar a "TABLE nohead"
            for r in rows: lines.append("|".join(c.replace("**", "") if rows.index(r) == 0 else c for c in r))
            lines.append("ENDTABLE")

head = ["@titulo|" + (titulo or "REVISAR: título"),
        "@para|", "@fecha|", "@cabecera|", "@pie|"]
open(OUT, "w", encoding="utf-8").write("\n".join(head + lines) + "\n")
print("OK ->", OUT, "| bloques:", len(lines))
for a in sorted(set(avisos)):
    print("AVISO: el original contiene %s; el formato no lo recoge y hay que decírselo al usuario." % a)
