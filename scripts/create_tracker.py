"""Create an empty tracker (tracker_vacantes_ia_backend.xlsx) with the layout the
other scripts expect: a "Vacantes" sheet with 13 columns and drop-down lists
for both status columns, and a "Resumen" sheet that counts postings by
application status.

It refuses to touch an existing tracker: that file is your data.

Usage:
    python scripts/create_tracker.py
"""

from __future__ import annotations

import argparse
import sys

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

from common import TRACKER_PATH

LAST_ROW = 200  # validations and Resumen formulas cover rows 2..LAST_ROW
COLUMNS = [  # (header, width)
    ("Fecha registrada", 14), ("Título", 30), ("Empresa", 20), ("Fuente", 22), ("Link", 40),
    ("Modalidad", 14), ("Ubicación", 18), ("Nivel", 16), ("Stack", 32), ("Compensación (aprox.)", 20),
    ("Estado publicación", 16), ("Estado postulación", 13), ("Notas", 42),
]
ESTADOS_PUBLICACION = ["Abierta", "Cerrada", "Programa recurrente", "Desconocido"]
ESTADOS_POSTULACION = ["No postulado", "Postulado", "En proceso", "Entrevista", "Rechazado", "Oferta"]

HEADER_FONT = Font(bold=True, color="FFFFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="FF0E7C86")
THIN = Side(style="thin", color="FFD0D7DE")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _list_validation(values: list[str], col: str) -> DataValidation:
    dv = DataValidation(type="list", formula1='"' + ",".join(values) + '"', allow_blank=True)
    dv.add(f"{col}2:{col}{LAST_ROW}")
    return dv


def build() -> openpyxl.Workbook:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Vacantes"
    for i, (header, width) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=i, value=header)
        cell.font, cell.fill, cell.border = HEADER_FONT, HEADER_FILL, BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[cell.column_letter].width = width
        # update_excel.py copies row 2's style to every new row: give it one.
        row2 = ws.cell(row=2, column=i)
        row2.alignment = Alignment(vertical="top", wrap_text=True)
        row2.border = BORDER
    ws.freeze_panes = "A2"
    ws.add_data_validation(_list_validation(ESTADOS_PUBLICACION, "K"))
    ws.add_data_validation(_list_validation(ESTADOS_POSTULACION, "L"))

    resumen = wb.create_sheet("Resumen")
    resumen["A1"] = "Resumen de seguimiento"
    resumen["A1"].font = Font(bold=True, size=14)
    resumen["A3"], resumen["B3"] = "Estado postulación", "Cantidad"
    for cell in (resumen["A3"], resumen["B3"]):
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL
    for row, estado in enumerate(ESTADOS_POSTULACION, start=4):
        resumen.cell(row=row, column=1, value=estado)
        resumen.cell(row=row, column=2, value=f"=COUNTIF(Vacantes!$L$2:$L${LAST_ROW},A{row})")
    total = 4 + len(ESTADOS_POSTULACION)
    resumen.cell(row=total, column=1, value="Total vacantes registradas").font = Font(bold=True)
    resumen.cell(row=total, column=2, value=f"=COUNTA(Vacantes!$B$2:$B${LAST_ROW})")
    resumen.cell(row=total + 3, column=1, value='Completá la columna "Estado postulación" en la hoja '
                                                'Vacantes: los conteos se recalculan solos al abrir el archivo.')
    resumen.column_dimensions["A"].width = 28
    resumen.column_dimensions["B"].width = 12
    return wb


def main() -> int:
    argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter).parse_args()
    if TRACKER_PATH.exists():
        print(f"{TRACKER_PATH.name} already exists; not touching it", file=sys.stderr)
        return 1
    TRACKER_PATH.parent.mkdir(parents=True, exist_ok=True)
    build().save(TRACKER_PATH)
    print(f"created {TRACKER_PATH.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
