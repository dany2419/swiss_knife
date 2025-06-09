#!/usr/bin/env python3
"""
Swiss Knife: conviértelo todo, sin pagarle a nadie.
Autor: Dani (aka Lúa) – Licencia MIT
"""

import click
from pathlib import Path

# --- PDF → DOCX ---------------------------------------------------------------
@click.group()
def cli():
    """Herramientas de conversión (PDF, DOCX, XLSX, imágenes)."""
    pass


@cli.command()
@click.argument("pdf_in", type=click.Path(exists=True))
@click.argument("docx_out", type=click.Path())
def pdf2docx(pdf_in, docx_out):
    """Convierte PDF_IN a DOCX_OUT."""
    from pdf2docx import Converter
    cv = Converter(pdf_in)
    cv.convert(docx_out, start=0, end=None)
    cv.close()
    click.echo(f"✅ PDF → DOCX completado: {docx_out}")


# --- DOCX → PDF ---------------------------------------------------------------
@cli.command()
@click.argument("docx_in", type=click.Path(exists=True))
@click.argument("pdf_out", type=click.Path())
def docx2pdf(docx_in, pdf_out):
    """Convierte DOCX_IN a PDF_OUT."""
    from docx2pdf import convert
    convert(docx_in, pdf_out)
    click.echo(f"✅ DOCX → PDF listo: {pdf_out}")


# --- PDF tabular → XLSX -------------------------------------------------------
@cli.command()
@click.argument("pdf_in", type=click.Path(exists=True))
@click.argument("xlsx_out", type=click.Path())
def pdf2xlsx(pdf_in, xlsx_out):
    """Extrae tablas de PDF_IN a XLSX_OUT."""
    import tabula
    dfs = tabula.read_pdf(pdf_in, pages="all", lattice=True)
    if dfs:
        with pd.ExcelWriter(xlsx_out) as writer:
            for i, df in enumerate(dfs):
                df.to_excel(writer, sheet_name=f"Tabla_{i+1}", index=False)
        click.echo(f"✅ PDF → XLSX creado: {xlsx_out}")
    else:
        click.echo("⚠️  No se encontraron tablas.")


# --- Imagen → PDF -------------------------------------------------------------
@cli.command()
@click.argument("img_in", type=click.Path(exists=True))
@click.argument("pdf_out", type=click.Path())
def img2pdf(img_in, pdf_out):
    """Empaqueta IMG_IN (JPG/PNG) en un PDF_OUT."""
    import img2pdf
    from PIL import Image
    img_path = Path(img_in)
    pdf_bytes = img2pdf.convert(str(img_path))
    Path(pdf_out).write_bytes(pdf_bytes)
    click.echo(f"✅ Imagen → PDF hecho: {pdf_out}")


# --- Unir PDFs ----------------------------------------------------------------
@cli.command()
@click.argument("pdfs", nargs=-1, type=click.Path(exists=True))
@click.argument("pdf_out", type=click.Path())
def mergepdf(pdfs, pdf_out):
    """Une uno o varios PDFs en PDF_OUT."""
    from PyPDF2 import PdfMerger
    merger = PdfMerger()
    for pdf in pdfs:
        merger.append(pdf)
    merger.write(pdf_out)
    merger.close()
    click.echo(f"✅ PDFs unidos en: {pdf_out}")


# --- Dividir PDF --------------------------------------------------------------
@cli.command()
@click.argument("pdf_in", type=click.Path(exists=True))
@click.argument("start", type=int)
@click.argument("end", type=int)
@click.argument("pdf_out", type=click.Path())
def splitpdf(pdf_in, start, end, pdf_out):
    """Extrae páginas START-END de PDF_IN a PDF_OUT (1-based)."""
    from PyPDF2 import PdfReader, PdfWriter
    reader = PdfReader(pdf_in)
    writer = PdfWriter()
    for p in range(start-1, end):
        writer.add_page(reader.pages[p])
    with open(pdf_out, "wb") as f:
        writer.write(f)
    click.echo(f"✅ Páginas {start}-{end} guardadas en {pdf_out}")


if __name__ == "__main__":
    cli()
