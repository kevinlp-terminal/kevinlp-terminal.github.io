"""Shared helpers for the UK M&A Terminal scripts."""
import shutil
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

TF = "Consolas"
AMBER, CYAN, LGREEN, GREY, ORANGE = "FFB000", "5FD7FF", "00E676", "9E9E9E", "FF8C00"
BG = "000000"
BOX = Border(*(Side(style="thin", color="3A3A3A"),) * 4)

LOG_FIRST = 5          # first data row of the Price Log tab
DEALS_FIRST = 5        # first data row of the Deals tab


def font(color=AMBER, bold=False, size=10, italic=False):
    return Font(name=TF, color=color, bold=bold, size=size, italic=italic)


def black_cell(cell, color=AMBER, bold=False, fmt=None):
    cell.fill = PatternFill("solid", fgColor=BG)
    cell.font = font(color, bold)
    cell.border = BOX
    if fmt:
        cell.number_format = fmt


def backup(path: Path) -> Path:
    bdir = path.parent / "backups"
    bdir.mkdir(exist_ok=True)
    dest = bdir / f"{path.stem}_{datetime.now():%Y%m%d_%H%M%S}{path.suffix}"
    shutil.copy2(path, dest)
    # keep the 20 most recent backups
    olds = sorted(bdir.glob(f"{path.stem}_*{path.suffix}"))[:-20]
    for o in olds:
        o.unlink()
    return dest


def open_workbook(path: Path):
    return load_workbook(path)


def deals_table(wb):
    """Return list of dicts for every row in the Deals tab (reads input values only)."""
    ws = wb["Deals"]
    out = []
    r = DEALS_FIRST
    while ws.cell(r, 2).value not in (None, ""):
        out.append({
            "id": ws.cell(r, 1).value, "target": ws.cell(r, 2).value, "bidder": ws.cell(r, 3).value,
            "total": ws.cell(r, 12).value, "undisturbed": ws.cell(r, 13).value, "status": ws.cell(r, 25).value,
        })
        r += 1
    return out


def add_precedents_chart(wb):
    """openpyxl drops charts when it re-saves a workbook, so the scripts rebuild this one."""
    ws = wb["Precedents"]
    th = None
    for r in range(1, 40):
        if ws.cell(r, 2).value == "#":
            th = r
            break
    if th is None:
        return
    n = 0
    while isinstance(ws.cell(th + 1 + n, 2).value, (int, float)):
        n += 1
    ch = BarChart()
    ch.type = "bar"
    ch.add_data(Reference(ws, min_col=9, min_row=th, max_row=th + n), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=3, min_row=th + 1, max_row=th + n))
    ch.legend = None
    ch.title = None
    ch.height, ch.width = 11, 18
    ch.roundedCorners = False
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.y_axis.numFmt = "0%"
    ch.y_axis.majorGridlines = None
    ch.x_axis.scaling.orientation = "maxMin"
    cp = CharacterProperties(solidFill=AMBER, sz=900)
    txt = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])
    ch.x_axis.txPr = txt
    ch.y_axis.txPr = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])
    ch.graphical_properties = GraphicalProperties(solidFill=BG, ln=LineProperties(solidFill="3A3A3A"))
    ch.plot_area.graphicalProperties = GraphicalProperties(solidFill=BG)
    ch.series[0].graphicalProperties.solidFill = ORANGE
    ch.series[0].graphicalProperties.line.solidFill = ORANGE
    ws._charts = []
    ws.add_chart(ch, "M5")


def save_workbook(wb, path: Path):
    add_precedents_chart(wb)
    wb.calculation.fullCalcOnLoad = True   # Excel recalculates everything on open
    wb.save(path)
