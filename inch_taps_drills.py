#!/usr/bin/env python3
"""
Create a properly formatted LibreOffice spreadsheet with merged cells
for the tap and drill sizes table.
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell

# Create workbook
wb = Workbook()
ws = wb.active
ws.title = "Tap & Drill Sizes"

# Define styles
header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
subheader_fill = PatternFill(start_color="E8E8E8", end_color="E8E8E8", fill_type="solid")
alt_row_fill = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")  # Light gray for alternating rows
bold_font = Font(bold=True)
center_align = Alignment(horizontal="center", vertical="center")
center_align_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)

# Row 1: Main headers with column spans
ws.merge_cells('A1:A2')  # Screw Size
ws['A1'] = "Screw Size"
ws['A1'].font = bold_font
ws['A1'].alignment = center_align_wrap

ws.merge_cells('B1:B2')  # Major Diameter
ws['B1'] = "Major Diameter"
ws['B1'].font = bold_font
ws['B1'].alignment = center_align_wrap

ws.merge_cells('C1:C2')  # Threads Per Inch
ws['C1'] = "TPI"
ws['C1'].font = bold_font
ws['C1'].alignment = center_align_wrap

ws.merge_cells('D1:D2')  # Minor Diameter
ws['D1'] = "Minor Diameter"
ws['D1'].font = bold_font
ws['D1'].alignment = center_align_wrap

# Tap Drill section
ws.merge_cells('E1:H1')  # Tap Drill header
ws['E1'] = "Tap Drill"
ws['E1'].font = bold_font
ws['E1'].alignment = center_align_wrap

ws.merge_cells('E2:F2')  # 75% Thread subsection
ws['E2'] = "75% Thread for Aluminum, Brass, Plastics"
ws['E2'].font = bold_font
ws['E2'].alignment = center_align_wrap

ws.merge_cells('G2:H2')  # 50% Thread subsection
ws['G2'] = "50% Thread for Stainless, Cast Iron & Iron"
ws['G2'].font = bold_font
ws['G2'].alignment = center_align_wrap

# Clearance Drill section
ws.merge_cells('I1:L1')  # Clearance Drill header
ws['I1'] = "Clearance Drill"
ws['I1'].font = bold_font
ws['I1'].alignment = center_align_wrap

ws.merge_cells('I2:J2')  # Close Fit subsection
ws['I2'] = "Close Fit"
ws['I2'].font = bold_font
ws['I2'].alignment = center_align_wrap

ws.merge_cells('K2:L2')  # Free Fit subsection
ws['K2'] = "Free Fit"
ws['K2'].font = bold_font
ws['K2'].alignment = center_align_wrap

# Row 3: Column detail headers
headers_row3 = [
    "", "", "", "",  # A-D (already merged from rows 1-2)
    "Drill Size", "Dec. Eq.",  # E-F (75% Thread)
    "Drill Size", "Dec. Eq.",  # G-H (50% Thread)
    "Drill Size", "Dec. Eq.",  # I-J (Close Fit)
    "Drill Size", "Dec. Eq."   # K-L (Free Fit)
]

for col, header in enumerate(headers_row3, start=1):
    if header:  # Skip empty cells (A-D)
        cell = ws.cell(row=3, column=col)
        cell.value = header
        cell.font = bold_font
        cell.alignment = center_align_wrap

# Data rows with row span information
# Reading CAREFULLY from the image, row by row
# Format: [screw_size, major_diam, threads, minor_diam, tap75_size, tap75_dec, tap50_size, tap50_dec, 
#          clear_close_size, clear_close_dec, clear_free_size, clear_free_dec, rowspan_for_screw, rowspan_for_clearance]
data = [
    # Row 1: Screw 0, 80 TPI (only one row for size 0)
    ["0", ".0600", "80", ".0447", "3/64", ".0469", "55", ".0520", "52", ".0635", "50", ".0700", 1, 1],
    # Rows 2-3: Screw 1, 64 TPI and 72 TPI (clearance spans both)
    ["1", ".0730", "64", ".0538", "53", ".0595", "1/16", ".0625", "48", ".0760", "46", ".0810", 2, 2],
    ["", "", "72", ".0560", "53", ".0595", "52", ".0635", "", "", "", "", 0, 0],
    # Rows 4-5: Screw 2, 56 TPI and 64 TPI (clearance spans both)
    ["2", ".0860", "56", ".0641", "50", ".0700", "49", ".0730", "43", ".0890", "41", ".0960", 2, 2],
    ["", "", "64", ".0668", "50", ".0700", "48", ".0760", "", "", "", "", 0, 0],
    # Rows 6-7: Screw 3, 48 TPI and 56 TPI (clearance spans both)
    ["3", ".0990", "48", ".0734", "47", ".0785", "44", ".0860", "37", ".1040", "35", ".1100", 2, 2],
    ["", "", "56", ".0771", "45", ".0820", "43", ".0890", "", "", "", "", 0, 0],
    # Rows 8-9: Screw 4, 40 TPI and 48 TPI (clearance spans both)
    ["4", ".1120", "40", ".0813", "43", ".0890", "41", ".0960", "32", ".1160", "30", ".1285", 2, 2],
    ["", "", "48", ".0864", "42", ".0935", "40", ".0980", "", "", "", "", 0, 0],
    # Rows 10-11: Screw 5, 40 TPI and 44 TPI (clearance spans both)
    ["5", ".125", "40", ".0943", "38", ".1015", "7/64", ".1094", "30", ".1285", "29", ".1360", 2, 2],
    ["", "", "44", ".0971", "37", ".1040", "35", ".1100", "", "", "", "", 0, 0],
    # Rows 12-13: Screw 6, 32 TPI and 40 TPI (clearance spans both)
    ["6", ".138", "32", ".0997", "36", ".1065", "32", ".1160", "27", ".1440", "25", ".1495", 2, 2],
    ["", "", "40", ".1073", "33", ".1130", "31", ".1200", "", "", "", "", 0, 0],
    # Rows 14-15: Screw 8, 32 TPI and 36 TPI (clearance spans both)
    ["8", ".1640", "32", ".1257", "29", ".1360", "27", ".1440", "18", ".1695", "16", ".1770", 2, 2],
    ["", "", "36", ".1299", "29", ".1360", "26", ".1470", "", "", "", "", 0, 0],
    # Rows 16-17: Screw 10, 24 TPI and 32 TPI (clearance spans both)
    ["10", ".1900", "24", ".1389", "25", ".1495", "20", ".1610", "9", ".1960", "7", ".2010", 2, 2],
    ["", "", "32", ".1517", "21", ".1590", "18", ".1695", "", "", "", "", 0, 0],
    # Rows 18-20: Screw 12, 24 TPI, 28 TPI, and 32 TPI (clearance spans all 3)
    ["12", ".2160", "24", ".1649", "16", ".1770", "12", ".1890", "2", ".2210", "1", ".2280", 3, 3],
    ["", "", "28", ".1722", "14", ".1820", "10", ".1935", "", "", "", "", 0, 0],
    ["", "", "32", ".1777", "13", ".1850", "9", ".1960", "", "", "", "", 0, 0],
    # Rows 21-23: Screw 1/4, 20 TPI, 28 TPI, and 32 TPI (clearance spans all 3)
    ["1/4", ".2500", "20", ".1887", "7", ".2010", "7/32", ".2188", "F", ".2570", "H", ".2660", 3, 3],
    ["", "", "28", ".2062", "3", ".2130", "1", ".2280", "", "", "", "", 0, 0],
    ["", "", "32", ".2117", "7/32", ".2188", "1", ".2280", "", "", "", "", 0, 0],
    # Rows 24-26: Screw 5/16, 18 TPI, 24 TPI, and 32 TPI (clearance spans all 3)
    ["5/16", ".3125", "18", ".2443", "F", ".2570", "J", ".2770", "P", ".3230", "Q", ".3320", 3, 3],
    ["", "", "24", ".2614", "I", ".2720", "9/32", ".2812", "", "", "", "", 0, 0],
    ["", "", "32", ".2742", "9/32", ".2812", "L", ".2900", "", "", "", "", 0, 0],
    # Rows 27-29: Screw 3/8, 16 TPI, 24 TPI, and 32 TPI (clearance spans all 3)
    ["3/8", ".3750", "16", ".2983", "5/16", ".3125", "Q", ".3320", "W", ".3860", "X", ".3970", 3, 3],
    ["", "", "24", ".3239", "Q", ".3320", "S", ".3480", "", "", "", "", 0, 0],
    ["", "", "32", ".3367", "11/32", ".3438", "T", ".3580", "", "", "", "", 0, 0],
    # Rows 30-32: Screw 7/16, 14 TPI, 20 TPI, and 28 TPI (clearance spans all 3)
    ["7/16", ".4375", "14", ".3499", "U", ".3680", "25/64", ".3906", "29/64", ".4531", "15/32", ".4687", 3, 3],
    ["", "", "20", ".3762", "25/64", ".3906", "13/32", ".4062", "", "", "", "", 0, 0],
    ["", "", "28", ".3937", "Y", ".4040", "Z", ".4130", "", "", "", "", 0, 0],
    # Rows 33-35: Screw 1/2, 13 TPI, 20 TPI, and 28 TPI (clearance spans all 3 rows)
    ["1/2", ".5000", "13", ".4056", "27/64", ".4219", "29/64", ".4531", "33/64", ".5156", "17/32", ".5312", 3, 3],
    ["", "", "20", ".4387", "29/64", ".4531", "15/32", ".4688", "", "", "", "", 0, 0],
    ["", "", "28", ".4562", "15/32", ".4688", "15/32", ".4688", "", "", "", "", 0, 0],
    # Rows 36-38: Screw 9/16, 12 TPI, 18 TPI, and 24 TPI (clearance spans all 3 rows)
    ["9/16", ".5625", "12", ".4603", "31/64", ".4844", "33/64", ".5156", "37/64", ".5781", "19/32", ".5938", 3, 3],
    ["", "", "18", ".4943", "33/64", ".5156", "17/32", ".5312", "", "", "", "", 0, 0],
    ["", "", "24", ".5114", "33/64", ".5156", "17/32", ".5312", "", "", "", "", 0, 0],
    # Rows 39-41: Screw 5/8, 11 TPI, 18 TPI, and 24 TPI (clearance spans all 3 rows)
    ["5/8", ".6250", "11", ".5135", "17/32", ".5312", "9/16", ".5625", "41/64", ".6406", "21/32", ".6562", 3, 3],
    ["", "", "18", ".5568", "37/64", ".5781", "19/32", ".5938", "", "", "", "", 0, 0],
    ["", "", "24", ".5739", "37/64", ".5781", "19/32", ".5938", "", "", "", "", 0, 0],
    # Row 42: Screw 11/16, 24 TPI only (clearance spans this row)
    ["11/16", ".6875", "24", ".6364", "41/64", ".6406", "21/32", ".6562", "45/64", ".7031", "23/32", ".7188", 1, 1],
    # Rows 43-45: Screw 3/4, 10 TPI, 16 TPI, and 20 TPI (clearance spans all 3)
    ["3/4", ".7500", "10", ".6273", "21/32", ".6562", "11/16", ".6875", "49/64", ".7656", "25/32", ".7812", 3, 3],
    ["", "", "16", ".6733", "11/16", ".6875", "45/64", ".7031", "", "", "", "", 0, 0],
    ["", "", "20", ".6887", "45/64", ".7031", "23/32", ".7188", "", "", "", "", 0, 0],
    # Row 46: Screw 13/16, 20 TPI only (clearance spans this row)
    ["13/16", ".8125", "20", ".7512", "49/64", ".7656", "25/32", ".7812", "53/64", ".8281", "27/32", ".8438", 1, 1],
    # Rows 47-49: Screw 7/8, 9 TPI, 14 TPI, and 20 TPI (clearance spans all 3)
    ["7/8", ".8750", "9", ".7387", "49/64", ".7656", "51/64", ".7969", "57/64", ".8906", "29/32", ".9062", 3, 3],
    ["", "", "14", ".7874", "13/16", ".8125", "53/64", ".8281", "", "", "", "", 0, 0],
    ["", "", "20", ".8137", "53/64", ".8281", "27/32", ".8438", "", "", "", "", 0, 0],
    # Row 50: Screw 15/16, 20 TPI only (clearance spans this row)
    ["15/16", ".9375", "20", ".8762", "57/64", ".8906", "29/32", ".9062", "61/64", ".9531", "31/32", ".9688", 1, 1],
    # Rows 51-53: Screw 1, 8 TPI, 12 TPI, and 20 TPI (clearance spans all 3)
    ["1", "1.000", "8", ".8466", "7/8", ".8750", "59/64", ".9219", "1-1/64", "1.0156", "1-1/32", "1.0313", 3, 3],
    ["", "", "12", ".8978", "15/16", ".9375", "61/64", ".9531", "", "", "", "", 0, 0],
    ["", "", "20", ".9387", "61/64", ".9531", "31/32", ".9688", "", "", "", "", 0, 0],
]

# Add data and handle row merging
current_row = 4
alternate_color = False  # Track alternating colors for screw sizes

for row_data in data:
    screw_rowspan = row_data[12]  # Rowspan for screw size/major diameter
    clear_rowspan = row_data[13]  # Rowspan for clearance drill columns
    
    # Toggle color when we hit a new screw size (when screw_rowspan > 0)
    if screw_rowspan > 0:
        alternate_color = not alternate_color
    
    # Merge screw size and major diameter first if needed
    if screw_rowspan > 1:
        ws.merge_cells(f'A{current_row}:A{current_row + screw_rowspan - 1}')
        ws.merge_cells(f'B{current_row}:B{current_row + screw_rowspan - 1}')
    
    # Merge clearance drill columns if needed
    if clear_rowspan > 1:
        ws.merge_cells(f'I{current_row}:I{current_row + clear_rowspan - 1}')
        ws.merge_cells(f'J{current_row}:J{current_row + clear_rowspan - 1}')
        ws.merge_cells(f'K{current_row}:K{current_row + clear_rowspan - 1}')
        ws.merge_cells(f'L{current_row}:L{current_row + clear_rowspan - 1}')
    
    # Write data to cells (first 12 columns)
    for col in range(1, 13):
        cell = ws.cell(row=current_row, column=col)
        if not isinstance(cell, MergedCell):
            cell.value = row_data[col-1] if row_data[col-1] else ""
        cell.alignment = center_align
        cell.border = thin_border
        # Apply alternating background color
        if alternate_color:
            cell.fill = alt_row_fill
    
    current_row += 1

# Apply borders to all header cells
for row in range(1, 4):
    for col in range(1, 13):
        ws.cell(row=row, column=col).border = thin_border

# Adjust column widths
ws.column_dimensions['A'].width = 10
ws.column_dimensions['B'].width = 10
ws.column_dimensions['C'].width = 10
ws.column_dimensions['D'].width = 10
for col in ['E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
    ws.column_dimensions[col].width = 11

# Set row heights for better readability
ws.row_dimensions[1].height = 30
ws.row_dimensions[2].height = 40
ws.row_dimensions[3].height = 20

# Save as Excel format (LibreOffice can open this)
wb.save('inch_taps_drills.xlsx')
print("Spreadsheet created: inch_taps_drills.xlsx")
print("This file can be opened in LibreOffice Calc with all merged cells preserved.")
