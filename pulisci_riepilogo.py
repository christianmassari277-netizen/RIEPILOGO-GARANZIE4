# -*- coding: utf-8 -*-
"""
pulisci_riepilogo.py
- Trascina un file di testo sull'eseguibile per ottenere RIEPILOGO_PULITO.pdf
- Oppure da terminale: pulisci_riepilogo.exe percorso\al\file.txt
"""

import pandas as pd
import re
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
import sys
import os


def _read_text(path):
    """Legge un file di testo provando più codifiche."""
    for enc in ("latin-1", "utf-8", "cp1252"):
        try:
            with open(path, "r", encoding=enc, errors="ignore") as f:
                return f.readlines()
        except Exception:
            continue
    with open(path, "r") as f:
        return f.readlines()


def processa_file(input_path, output_path):
    lines = _read_text(input_path)

    # Regex: 7 cifre (NUMERO GARANZIA), poi SUFFISSO, JOB e TOTALE JOB
    pattern = re.compile(r"^\s*(\d{7})\s+(\d+)\s+(\d+)\s+(\d+)\s*$")

    records = []
    for line in lines:
        m = pattern.match(line)
        if m:
            numero_garanzia, suffisso, job, totale_job = m.groups()
            records.append([numero_garanzia, suffisso, int(job), int(totale_job)])

    if not records:
        # PDF minimale in caso di nessun match
        doc = SimpleDocTemplate(output_path, pagesize=A4)
        styles = getSampleStyleSheet()
        elements = [
            Paragraph("Nessun dato riconosciuto", styles["Title"]),
            Paragraph("Il file non contiene righe nel formato atteso.", styles["BodyText"])
        ]
        doc.build(elements)
        return

    df = pd.DataFrame(records, columns=["NUMERO GARANZIA", "SUFFISSO", "JOB", "TOTALE JOB"])
    df = df.drop_duplicates(subset=["NUMERO GARANZIA", "SUFFISSO", "JOB"], keep="first")
    df = df.sort_values(by=["NUMERO GARANZIA", "SUFFISSO", "JOB"])
    totale_generale = int(df["TOTALE JOB"].sum())

    # Creazione PDF
    doc = SimpleDocTemplate(output_path, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = []
    elements.append(Paragraph("Riepilogo Garanzie", styles["Title"]))
    elements.append(Spacer(1, 12))

    data = [list(df.columns)] + df.astype(str).values.tolist()
    data.append(["", "", "TOTALE", str(totale_generale)])

    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("TEXTCOLOR", (0,0), (-1,0), colors.black),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.black),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("BACKGROUND", (0,-1), (-1,-1), colors.lightgrey),
        ("FONTNAME", (0,-1), (-1,-1), "Helvetica-Bold"),
    ]))
    elements.append(table)
    doc.build(elements)


def main():
    if len(sys.argv) >= 2:
        input_path = sys.argv[1]
        base, _ = os.path.splitext(input_path)
        output_path = base + "_PULITO.pdf"
        processa_file(input_path, output_path)
    else:
        # Messaggio di aiuto su Windows (MessageBox), altrimenti stampa in console
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(
                0,
                "Trascina un file .txt sopra l'eseguibile per elaborarlo.\n"
                "Oppure lancialo da terminale con: pulisci_riepilogo.exe <percorso file>",
                "pulisci_riepilogo",
                0x40
            )
        except Exception:
            print("Uso: pulisci_riepilogo.exe <percorso file .txt>")


if __name__ == "__main__":
    main()
