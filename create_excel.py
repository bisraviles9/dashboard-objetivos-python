from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.table import Table, TableStyleInfo

OUT = Path(__file__).with_name("dashboard_objetivos.xlsx")
goals = [
    ["Ventas", "Aumentar ventas", "Ana", 100000, 68000, "USD", "2026-12-31", "T4 2026"],
    ["Marketing", "Generar oportunidades", "Luis", 500, 245, "leads", "2026-12-31", "T4 2026"],
    ["Producto", "Entregar mejoras", "Marta", 12, 8, "entregas", "2026-12-31", "T4 2026"],
    ["Atención al cliente", "Mejorar satisfacción", "Diego", 90, 74, "%", "2026-12-31", "T4 2026"],
    ["Operaciones", "Reducir tiempo de entrega", "Sofía", 30, 12, "% reducción", "2026-12-31", "T4 2026"],
]
months = []
for area, objective, unit, values in [
    ("Ventas", "Aumentar ventas", "USD", [4000, 4500, 5200, 5600, 6000, 6500, 7000, 7200, 8000]),
    ("Marketing", "Generar oportunidades", "leads", [22, 25, 28, 30, 32, 35, 36, 38, 42]),
    ("Producto", "Entregar mejoras", "entregas", [1, 1, 1, 1, 1, 1, 1, 0, 0]),
    ("Atención al cliente", "Mejorar satisfacción", "%", [70, 71, 72, 72, 73, 73, 74, 74, 74]),
]:
    for month, value in zip(["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre"], values):
        months.append([month, area, objective, value, unit])

wb = Workbook()
summary = wb.active
summary.title = "Resumen"
summary.append(["Área", "Objetivo", "Responsable", "Meta", "Actual", "Unidad", "Progreso", "Fecha límite", "Trimestre"])
for row in goals:
    area, obj, owner, target, actual, unit, due, quarter = row
    progress = min(actual / target, 1) if target else 0
    summary.append([area, obj, owner, target, actual, unit, progress, due, quarter])

ws = wb.create_sheet("Objetivos")
ws.append(["Área", "Objetivo", "Responsable", "Meta", "Actual", "Unidad", "Fecha límite", "Trimestre"])
for row in goals:
    ws.append(row)

monthly = wb.create_sheet("Seguimiento mensual")
monthly.append(["Mes", "Área", "Objetivo", "Avance mensual", "Unidad"])
for row in months:
    monthly.append(row)

for sheet in wb.worksheets:
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    sheet.sheet_view.showGridLines = False
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="E4793D")
        cell.alignment = Alignment(vertical="center")
    sheet.row_dimensions[1].height = 24
    for cells in sheet.columns:
        letter = cells[0].column_letter
        width = min(max(max(len(str(cell.value or "")) for cell in cells) + 2, 12), 34)
        sheet.column_dimensions[letter].width = width
    if sheet.max_row > 1:
        table = Table(displayName="Table" + str(wb.sheetnames.index(sheet.title) + 1), ref=sheet.dimensions)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False)
        sheet.add_table(table)

for row in range(2, summary.max_row + 1):
    summary.cell(row, 7).number_format = "0%"
    summary.cell(row, 4).number_format = '#,##0.##'
    summary.cell(row, 5).number_format = '#,##0.##'

chart = BarChart()
chart.type = "bar"
chart.style = 10
chart.title = "Progreso por área"
chart.y_axis.title = "Área"
chart.x_axis.title = "Progreso"
chart.x_axis.numFmt = "0%"
chart.add_data(Reference(summary, min_col=7, min_row=1, max_row=summary.max_row), titles_from_data=True)
chart.set_categories(Reference(summary, min_col=1, min_row=2, max_row=summary.max_row))
chart.height, chart.width = 8, 14
summary.add_chart(chart, "A9")

wb.save(OUT)
print(OUT)
