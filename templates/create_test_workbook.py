#!/usr/bin/env python3
"""
Create Surplus Lines Tax Excel Add-in Test Workbook
Mirrors the Google Sheets test template structure
"""

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import os

def create_test_workbook():
    wb = Workbook()

    # Define styles
    header_font = Font(bold=True, size=14, color="FFFFFF")
    subheader_font = Font(bold=True, size=11)
    header_fill = PatternFill(start_color="2E3C43", end_color="2E3C43", fill_type="solid")
    accent_fill = PatternFill(start_color="3DE2A0", end_color="3DE2A0", fill_type="solid")
    light_fill = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
    code_font = Font(name="Consolas", size=10)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # =========================================================================
    # Sheet 1: Calculator
    # =========================================================================
    ws1 = wb.active
    ws1.title = "Calculator"

    # Title
    ws1['A1'] = "Surplus Lines Tax Calculator"
    ws1['A1'].font = Font(bold=True, size=16)
    ws1.merge_cells('A1:E1')

    # Section 1: Simple Tax Calculation
    ws1['A3'] = "State"
    ws1['B3'] = "Premium"
    ws1['C3'] = "Total Tax"
    for col in ['A', 'B', 'C']:
        ws1[f'{col}3'].font = subheader_font
        ws1[f'{col}3'].fill = accent_fill

    # Test data
    test_data = [
        ("Texas", 10000),
        ("California", 25000),
        ("Florida", 15000),
        ("New York", 50000),
        ("Illinois", 7500),
    ]

    for i, (state, premium) in enumerate(test_data, start=4):
        ws1[f'A{i}'] = state
        ws1[f'B{i}'] = premium
        ws1[f'C{i}'] = f'=SLTAX.CALCULATE(A{i}, B{i})'

    # Section 2: Detailed Breakdown
    ws1['A11'] = "Detailed Breakdown (SLTAX.CALCULATE_DETAILS)"
    ws1['A11'].font = Font(bold=True, size=12)
    ws1.merge_cells('A11:E11')

    ws1['A12'] = "State"
    ws1['B12'] = "Premium"
    ws1['C12'] = "State (returned)"
    ws1['D12'] = "Premium (returned)"
    ws1['E12'] = "Total Tax"
    ws1['F12'] = "Total Due"
    for col in ['A', 'B', 'C', 'D', 'E', 'F']:
        ws1[f'{col}12'].font = subheader_font
        ws1[f'{col}12'].fill = light_fill

    ws1['A13'] = "Texas"
    ws1['B13'] = 10000
    ws1['C13'] = '=SLTAX.CALCULATE_DETAILS(A13, B13)'  # Spills to F13

    ws1['A14'] = "California"
    ws1['B14'] = 25000
    ws1['C14'] = '=SLTAX.CALCULATE_DETAILS(A14, B14)'

    # Section 3: With Premium
    ws1['A17'] = "Compact View (SLTAX.CALCULATE_WITHPREMIUM)"
    ws1['A17'].font = Font(bold=True, size=12)
    ws1.merge_cells('A17:E17')

    ws1['A18'] = "State"
    ws1['B18'] = "Premium"
    ws1['C18'] = "Total Tax"
    ws1['D18'] = "Total Due"
    for col in ['A', 'B', 'C', 'D']:
        ws1[f'{col}18'].font = subheader_font
        ws1[f'{col}18'].fill = light_fill

    ws1['A19'] = "Florida"
    ws1['B19'] = '=SLTAX.CALCULATE_WITHPREMIUM("Florida", 15000)'  # Spills to D19

    # Adjust column widths
    ws1.column_dimensions['A'].width = 15
    ws1.column_dimensions['B'].width = 15
    ws1.column_dimensions['C'].width = 18
    ws1.column_dimensions['D'].width = 18
    ws1.column_dimensions['E'].width = 15
    ws1.column_dimensions['F'].width = 15

    # =========================================================================
    # Sheet 2: Rates
    # =========================================================================
    ws2 = wb.create_sheet("Rates")

    ws2['A1'] = "Tax Rate Lookup (SLTAX.RATE)"
    ws2['A1'].font = Font(bold=True, size=14)
    ws2.merge_cells('A1:C1')

    ws2['A3'] = "State"
    ws2['B3'] = "Rate (%)"
    ws2['A3'].font = subheader_font
    ws2['B3'].font = subheader_font
    ws2['A3'].fill = accent_fill
    ws2['B3'].fill = accent_fill

    rate_states = ["Texas", "California", "Florida", "New York", "Iowa", "Alabama"]
    for i, state in enumerate(rate_states, start=4):
        ws2[f'A{i}'] = state
        ws2[f'B{i}'] = f'=SLTAX.RATE("{state}")'

    # All Rates section
    ws2['A12'] = "All States & Rates (SLTAX.RATES)"
    ws2['A12'].font = Font(bold=True, size=12)
    ws2.merge_cells('A12:C12')

    ws2['A13'] = "State"
    ws2['B13'] = "Tax Rate (%)"
    ws2['A13'].font = subheader_font
    ws2['B13'].font = subheader_font
    ws2['A13'].fill = light_fill
    ws2['B13'].fill = light_fill

    ws2['A14'] = '=SLTAX.RATES()'  # Spills 53 rows × 2 columns

    ws2.column_dimensions['A'].width = 20
    ws2.column_dimensions['B'].width = 15

    # =========================================================================
    # Sheet 3: All States
    # =========================================================================
    ws3 = wb.create_sheet("All States")

    ws3['A1'] = "All 53 Jurisdictions (SLTAX.STATES)"
    ws3['A1'].font = Font(bold=True, size=14)
    ws3.merge_cells('A1:B1')

    ws3['A3'] = "State Name"
    ws3['A3'].font = subheader_font
    ws3['A3'].fill = accent_fill

    ws3['A4'] = '=SLTAX.STATES()'  # Spills 53 rows

    ws3.column_dimensions['A'].width = 25

    # =========================================================================
    # Sheet 4: Detailed Rates
    # =========================================================================
    ws4 = wb.create_sheet("Detailed Rates")

    ws4['A1'] = "Complete Rate Details (SLTAX.RATES_DETAILS)"
    ws4['A1'].font = Font(bold=True, size=14)
    ws4.merge_cells('A1:K1')

    headers = ["State", "Tax Rate", "Stamping Fee", "Filing Fee", "Service Fee",
               "Surcharge", "Regulatory Fee", "Fire Marshal", "SLAS Fee", "Flat Fee", "Source"]
    for i, header in enumerate(headers, start=1):
        col = get_column_letter(i)
        ws4[f'{col}3'] = header
        ws4[f'{col}3'].font = subheader_font
        ws4[f'{col}3'].fill = light_fill

    ws4['A4'] = '=SLTAX.RATES_DETAILS()'  # Spills 53 rows × 11 columns

    for i in range(1, 12):
        ws4.column_dimensions[get_column_letter(i)].width = 14

    # =========================================================================
    # Sheet 5: Historical
    # =========================================================================
    ws5 = wb.create_sheet("Historical")

    ws5['A1'] = "Historical Rate Lookup"
    ws5['A1'].font = Font(bold=True, size=14)
    ws5.merge_cells('A1:D1')

    # Simple historical rate
    ws5['A3'] = "State"
    ws5['B3'] = "Date"
    ws5['C3'] = "Rate (%)"
    for col in ['A', 'B', 'C']:
        ws5[f'{col}3'].font = subheader_font
        ws5[f'{col}3'].fill = accent_fill

    historical_data = [
        ("Iowa", "2025-06-15"),
        ("Texas", "2024-01-01"),
        ("California", "2023-07-01"),
    ]

    for i, (state, date) in enumerate(historical_data, start=4):
        ws5[f'A{i}'] = state
        ws5[f'B{i}'] = date
        ws5[f'C{i}'] = f'=SLTAX.HISTORICALRATE(A{i}, B{i})'

    # Detailed historical
    ws5['A9'] = "Historical Rate Details (Horizontal)"
    ws5['A9'].font = Font(bold=True, size=12)
    ws5.merge_cells('A9:O9')

    detail_headers = ["State", "Date", "Tax Rate", "Stamping Fee", "Filing Fee", "Service Fee",
                      "Surcharge", "Regulatory Fee", "Fire Marshal", "SLAS Fee", "Flat Fee",
                      "Effective From", "Effective To", "Legislative Source", "Confidence"]
    for i, header in enumerate(detail_headers, start=1):
        col = get_column_letter(i)
        ws5[f'{col}10'] = header
        ws5[f'{col}10'].font = subheader_font
        ws5[f'{col}10'].fill = light_fill

    ws5['A11'] = '=SLTAX.HISTORICALRATE_DETAILS("Texas", "2024-01-01")'  # Spills 15 columns

    # Vertical view
    ws5['A14'] = "Historical Rate Details (Vertical - multiline=TRUE)"
    ws5['A14'].font = Font(bold=True, size=12)
    ws5.merge_cells('A14:D14')

    ws5['A15'] = '=SLTAX.HISTORICALRATE_DETAILS("Texas", "2024-01-01", TRUE)'  # Spills 15 rows

    ws5.column_dimensions['A'].width = 18
    ws5.column_dimensions['B'].width = 12
    ws5.column_dimensions['C'].width = 12

    # =========================================================================
    # Sheet 6: Quick Reference (FIXED - Plain text, no formulas)
    # =========================================================================
    ws6 = wb.create_sheet("Quick Reference")

    ws6['A1'] = "Surplus Lines Tax - Excel Function Reference"
    ws6['A1'].font = Font(bold=True, size=16)
    ws6.merge_cells('A1:E1')

    # Headers
    ws6['A3'] = "Function"
    ws6['B3'] = "Description"
    ws6['C3'] = "Example"
    ws6['D3'] = "Returns"
    for col in ['A', 'B', 'C', 'D']:
        ws6[f'{col}3'].font = header_font
        ws6[f'{col}3'].fill = header_fill

    # Function documentation - ALL TEXT, no formulas (prefix with apostrophe or just use plain text)
    functions = [
        ("SLTAX.CALCULATE(state, premium)", "Calculate total tax", 'SLTAX.CALCULATE("Texas", 10000)', "503 (number)"),
        ("SLTAX.CALCULATE_DETAILS(state, premium, [multiline])", "Full breakdown", 'SLTAX.CALCULATE_DETAILS("CA", 10000)', "state, premium, tax, due"),
        ("SLTAX.CALCULATE_WITHPREMIUM(state, premium)", "Compact breakdown", 'SLTAX.CALCULATE_WITHPREMIUM("FL", 15000)', "premium, tax, due"),
        ("SLTAX.RATE(state)", "Get tax rate %", 'SLTAX.RATE("California")', "3 (number)"),
        ("SLTAX.STATES()", "List all jurisdictions", "SLTAX.STATES()", "53 state names (vertical)"),
        ("SLTAX.RATES()", "All states with rates", "SLTAX.RATES()", "state, rate (53 rows)"),
        ("SLTAX.RATES_DETAILS()", "All rates with full fees", "SLTAX.RATES_DETAILS()", "11 columns x 53 rows"),
        ("SLTAX.HISTORICALRATE(state, date)", "Historical rate lookup", 'SLTAX.HISTORICALRATE("Iowa", "2025-06-15")', "0.95 (number)"),
        ("SLTAX.HISTORICALRATE_DETAILS(state, date, [multiline])", "Full historical info", 'SLTAX.HISTORICALRATE_DETAILS("TX", "2024-01-01")', "15 columns"),
    ]

    for i, (func, desc, example, returns) in enumerate(functions, start=4):
        # Use plain text (no = prefix means Excel treats it as text)
        ws6[f'A{i}'] = func
        ws6[f'A{i}'].font = code_font
        ws6[f'B{i}'] = desc
        ws6[f'C{i}'] = example
        ws6[f'C{i}'].font = code_font
        ws6[f'D{i}'] = returns
        if i % 2 == 0:
            for col in ['A', 'B', 'C', 'D']:
                ws6[f'{col}{i}'].fill = light_fill

    # API Key Setup section
    ws6['A15'] = "API Key Setup:"
    ws6['A15'].font = subheader_font
    ws6['A16'] = "1. Click 'Surplus Lines Tax' in the ribbon → 'Settings'"
    ws6['A17'] = "2. Enter your API key from app.surpluslinesapi.com"
    ws6['A18'] = "3. Click 'Save API Key'"

    ws6['A20'] = "Get your API key: https://app.surpluslinesapi.com"
    ws6['A20'].font = Font(color="0066CC", underline="single")

    # Usage note
    ws6['A22'] = "Usage Notes:"
    ws6['A22'].font = subheader_font
    ws6['A23'] = "• Functions that return multiple values will SPILL into adjacent cells"
    ws6['A24'] = "• Make sure cells to the right/below are empty for spill functions"
    ws6['A25'] = "• Use multiline=TRUE parameter for vertical output instead of horizontal"

    # Column widths
    ws6.column_dimensions['A'].width = 50
    ws6.column_dimensions['B'].width = 22
    ws6.column_dimensions['C'].width = 45
    ws6.column_dimensions['D'].width = 25

    # Save workbook
    output_path = os.path.join(os.path.dirname(__file__), "SurplusLinesTax-Test-Template.xlsx")
    wb.save(output_path)
    print(f"Created: {output_path}")
    return output_path

if __name__ == "__main__":
    create_test_workbook()
