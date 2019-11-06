import openpyxl

def build_sheet_from_book(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.worksheets[0]
    return sheet

# returns site name stripped of location and customer information
def site_name_string_manager(site):
        return site[5:].split('(')[0].split('TX')[0]
