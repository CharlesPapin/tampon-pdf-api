from fastapi import FastAPI, Request
import pandas as pd
import fitz  # PyMuPDF
import os

app = FastAPI()

@app.post("/tamponner")
async def tamponner(request: Request):
    data = await request.json()
    excel_path = data["excel_path"]
    pdf_path = data["pdf_path"]
    sheet_name = data.get("sheet_name", "LD TIDE 2025")

    pdf_basename = os.path.splitext(os.path.basename(pdf_path))[0]
    output_pdf_path = os.path.join(os.path.dirname(pdf_path), f"{pdf_basename}_tamponné.pdf")

    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=1, engine="openpyxl")

    with fitz.open(pdf_path) as doc:
        text = doc[0].get_text()
        for line in text.splitlines():
            if "N° de PO" in line:
                po_number = line.split(":")[-1].strip()
                break
        else:
            raise ValueError("Numéro de PO non trouvé dans le PDF.")

    row = df[df["N° de commande"] == po_number].iloc[0]

    code_p1 = row["Code P1 Réduit"]
    code_p2 = row["Code P2 réduit"]
    code_p3 = row["Code P3 réduit"]
    controle_par = row["Qui"]
    code_entite = "131"
    validation_par = "Alban BILLAUD"

    lines = [
        f"N° de PO : {po_number}",
        f"Code entité : {code_entite}",
        f"Code P1 : {code_p1}",
        f"Code P2 : {code_p2}",
        f"Code P3 : {code_p3}",
        f"Contrôlé par : {controle_par}",
        f"Validation par : {validation_par}"
    ]

    doc = fitz.open(pdf_path)
    page = doc[0]

    line_height = 14
    padding = 5
    num_lines = len(lines)
    table_height = num_lines * line_height + 2 * padding
    table_width = 250
    margin_x = 40
    margin_y = 40

    page_height = page.rect.height
    x0 = margin_x
    y0 = page_height - margin_y - table_height
    x1 = x0 + table_width
    y1 = y0 + table_height

    page.draw_rect(fitz.Rect(x0, y0, x1, y1), color=(1, 1, 1), fill=(1, 1, 1))
    page.draw_rect(fitz.Rect(x0, y0, x1, y1), color=(1, 0, 0), width=1)

    for i in range(1, num_lines):
        y = y0 + padding + i * line_height
        page.draw_line(fitz.Point(x0, y), fitz.Point(x1, y), color=(1, 0, 0), width=1)

    for i, line in enumerate(lines):
        y_text = y0 + padding + i * line_height + line_height / 2
        rect = fitz.Rect(x0 + 5, y_text - line_height / 2, x1 - 5, y_text + line_height / 2)
        page.insert_textbox(rect, line, fontsize=10, color=(1, 0, 0), align=0)

    doc.save(output_pdf_path)
    doc.close()

    return {"message": "PDF tamponné sauvegardé", "output_path": output_pdf_path}
