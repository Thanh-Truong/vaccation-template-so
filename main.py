import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from openpyxl.styles import Color
from openpyxl.formatting.rule import CellIsRule


from copy import copy
import utils
from datetime import date

EXCEL_MONTH_ROW = 2
EXCEL_WEEKNUMBER_ROW = 3
EXCEL_DATE_ROW = 4
EXCEL_WEEK_DATE_ROW = 5
PASSWORD = '12345'

# source_wb, 
# destination_wb, 
# new_sheet_name
def clone_a_sheet(source_wb, destination_wb, new_sheet_name):
    # Load the Excel workbook
    workbook = openpyxl.load_workbook(source_wb)

    # Select the sheet you want to clone
    source_sheet = workbook['Template-sheet']

    # Create a new sheet to clone into
    new_sheet = workbook.copy_worksheet(source_sheet)
    new_sheet.title = new_sheet_name  # Set a new name for the cloned sheet

    # Save the updated Excel workbook
    workbook.save(destination_wb)

def set_start_date(workbook_path, sheet_name, start_date):
    # Load the Excel workbook
    wb = openpyxl.load_workbook(workbook_path)

    # Select the sheet
    ws = wb[sheet_name]

    # Create a new sheet to clone into
    ws['E4'].value = start_date

    # Save the updated Excel workbook
    wb.save(workbook_path)


def copy_date_formulas(destination_column, destination_column_letter, source_column_letter, date_diff):
    # Python starts index from 0 while Excel starts from 1
    destination_column[EXCEL_DATE_ROW - 1].value = "={}{}+{}".format(source_column_letter, EXCEL_DATE_ROW, date_diff)
    destination_column[EXCEL_WEEK_DATE_ROW -1].value = '=TEXT({}{},"ddd")'.format(destination_column_letter, EXCEL_DATE_ROW)

def copy_format_column(worksheet, source_column_letter, destination_column_letter, date_diff):
    source_column = worksheet[source_column_letter]
    destination_column = worksheet[destination_column_letter]

    # Copy the format from the source column to the newly inserted column
    for source_cell, new_cell in zip(source_column, destination_column):
        new_cell.font = copy(source_cell.font)
        new_cell.fill = copy(source_cell.fill)
        new_cell.border = copy(source_cell.border)
        new_cell.alignment = copy(source_cell.alignment)
        new_cell.number_format = source_cell.number_format
        new_cell.protection = copy(source_cell.protection)
    
    # Set the width of the newly inserted column to match the starting column
    worksheet.column_dimensions[destination_column_letter].width = worksheet.column_dimensions[source_column_letter].width
    
    # Copy date formulas to display "day" on row 4th and "weekday" on row 5th
    copy_date_formulas(destination_column, destination_column_letter, 
                       source_column_letter, date_diff)

def insert_columns_with_format(workbook_path, sheet_name, source_column_letter, date_count):
    # Load the Excel workbook
    workbook = openpyxl.load_workbook(workbook_path)
    
    # Select the worksheet
    worksheet = workbook[sheet_name]
    
    # Get the index of the starting column
    start_column_index = openpyxl.utils.column_index_from_string(source_column_letter)
    
    # Insert column_count to the right 
    worksheet.insert_cols(idx=start_column_index + 1, amount=date_count)

    for date_diff in range(1, date_count):
        destination_column_letter = get_column_letter(start_column_index + date_diff)
        copy_format_column(worksheet, source_column_letter, destination_column_letter, date_diff)

    # Save the updated Excel workbook
    workbook.save(workbook_path)

def colourize_and_lock_weekend(workbook_path, sheet_name, source_column_letter, start_date, date_count):
    # Load the Excel workbook
    workbook = openpyxl.load_workbook(workbook_path)
    
    # Select the worksheet
    worksheet = workbook[sheet_name]
    
    # Get the index of the starting column
    start_column_index = openpyxl.utils.column_index_from_string(source_column_letter)
    # Define the grey fill color
    grey_fill = PatternFill(start_color="808080", end_color="808080", fill_type="solid")

    for i in range(0, date_count):
        current_date = utils.date_add(input_date_obj=start_date, days=i)
        if utils.is_Weekend(current_date):
            # Lock cells and fill them with GREY
            for row in range(4, 124):
                cell = worksheet.cell(row=row, column=start_column_index + i)
                cell.fill = grey_fill
                cell.protection = openpyxl.styles.protection.Protection(locked=True)
    
    # Save the updated Excel workbook
    workbook.save(workbook_path)

def write_months_as_headers(workbook_path, sheet_name, source_column_letter, start_date, end_date):
    # Load the Excel workbook
    workbook = openpyxl.load_workbook(workbook_path)
    
    # Select the worksheet
    worksheet = workbook[sheet_name]
    
    # Get the index of the starting column
    start_column_index = openpyxl.utils.column_index_from_string(source_column_letter)
    monthly_ranges = utils.generate_monthly_ranges_within_year(start_date, end_date)    
    month = 0
    for (month_start, month_end) in monthly_ranges:
        if month_start and month_end:
            # Caculate the month_start column index
            start_date_diff = utils.count_dates_between_inclusive(start_date, month_start)
            month_start_column_index = start_column_index + start_date_diff - 1
            month_start_column_letter = get_column_letter(month_start_column_index)

            # Caculate the month_end column index
            end_date_diff = utils.count_dates_between_inclusive(start_date, month_end)
            month_end_column_index = start_column_index + end_date_diff - 1
            month_end_column_letter = get_column_letter(month_end_column_index)

            # Merge a range of cells to form a month
            merge_range = f"{month_start_column_letter}{EXCEL_MONTH_ROW}:{month_end_column_letter}{EXCEL_MONTH_ROW}"
            #print("merge_range start {} end {}".format(month_start, month_end))
            #print("merge_range start column {} end column {}".format(month_start_column_letter, month_end_column_letter))
            #print(merge_range)
            worksheet.merge_cells(merge_range)
            # Get the merged cell
            merged_cell = worksheet[f"{month_start_column_letter}{EXCEL_MONTH_ROW}"]
            # Set the month as string in the merged cells
            merged_cell.value = utils.MONTHS_OF_A_YEAR[month]
            # Colourize the month
            color_to_fill = utils.MONTHS_COLORS[month]

            merged_cell.fill = color_to_fill

        month = month + 1

    # Save the updated Excel workbook
    workbook.save(workbook_path)

def write_weeks_as_headers(workbook_path, sheet_name, source_column_letter, start_date, end_date):
    # Load the Excel workbook
    workbook = openpyxl.load_workbook(workbook_path)
    
    # Select the worksheet
    worksheet = workbook[sheet_name]
    
    # Get the index of the starting column
    start_column_index = openpyxl.utils.column_index_from_string(source_column_letter)
    week_ranges = utils.week_ranges_in_range(start_date, end_date)

    for (week_number, (week_start, week_end)) in week_ranges:
        # Caculate the week_start column index
        diff = utils.count_dates_between_inclusive(start_date, week_start)
        week_start_column_index = start_column_index + diff - 1
        week_start_column_letter = get_column_letter(week_start_column_index)

        # Caculate the week_end column index
        diff = utils.count_dates_between_inclusive(start_date, week_end)
        week_end_column_index = start_column_index + diff - 1
        week_end_column_letter = get_column_letter(week_end_column_index)

        # Merge a range of cells to form a month
        merge_range = f"{week_start_column_letter}{EXCEL_WEEKNUMBER_ROW}:{week_end_column_letter}{EXCEL_WEEKNUMBER_ROW}"
        #print("merge_range start {} end {}".format(week_start_column_letter, week_end_column_letter))
        #print(merge_range)
        worksheet.merge_cells(merge_range)
        # Get the merged cell
        merged_cell = worksheet[f"{week_start_column_letter}{EXCEL_WEEKNUMBER_ROW}"]
        # Set the month as string in the merged cells
        merged_cell.value = f"Week {week_number}"

    # Save the updated Excel workbook
    workbook.save(workbook_path)


def apply_conditional_formatting(workbook_path, sheet_name, range_to_format):
    """ # Load the Excel workbook
    workbook = openpyxl.load_workbook(workbook_path)

    # Select the worksheet
    worksheet = workbook[sheet_name]

    # Green fill
    green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")

    # Define the rule: Fill cell with green if cell value is equal to "v"
    rule = CellIsRule(operator='equal', formula=['v'], stopIfTrue=True, fill=redFill))
    
        formula=['$A1="v"'],
        stopIfTrue=False,
        fill=PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
    )

    # Add the rule to the worksheet
    worksheet.conditional_formatting.add(range_to_format, rule)

    # Save the updated Excel workbook
    workbook.save(workbook_path) """
    

def create_vaccation_period(source_wb, destination_wb, sheet_name, start_date, end_date):
    start_column_letter = 'E'

    # Clone the Template-sheet to a new sheet 
    clone_a_sheet(source_wb, destination_wb, sheet_name)
    # On the new workbook, let set the start_date at E4 
    set_start_date(destination_wb, sheet_name, start_date)

    date_count = utils.count_dates_between_inclusive(start_date, end_date)
    #print(f"There are {date_count} inclusive dates between {start_date} and {end_date}.")

    # Insert columns to the right with the same format as start_column_letter
    insert_columns_with_format(destination_wb, sheet_name, start_column_letter, date_count)
    colourize_and_lock_weekend(destination_wb, sheet_name, start_column_letter, start_date, date_count)
    write_months_as_headers(destination_wb,sheet_name, start_column_letter,start_date,end_date)
    write_weeks_as_headers(destination_wb,sheet_name, start_column_letter,start_date,end_date)
    
    apply_conditional_formatting(destination_wb, sheet_name, 'E6:DU123')

    # Lock the worksheet
    workbook = openpyxl.load_workbook(destination_wb)
    # Select the worksheet
    worksheet = workbook[sheet_name]
    # Protect the sheet with a password
    worksheet.protection.sheet = True
    worksheet.protection.password = PASSWORD

    
    workbook.save(destination_wb)

def main():
    source_wb='2023-11-03-Python.xlsx'
    destination_wb='2023-11-03-Python_updated.xlsx'
    sheet_name='Test'
    start_date = date(2024,1,1)
    end_date = date(2024,4,30)
    create_vaccation_period(source_wb=source_wb, destination_wb=destination_wb, sheet_name="Test-January-April",
                             start_date=start_date, end_date=end_date)

if __name__ == "__main__":
    main()
    #test()
