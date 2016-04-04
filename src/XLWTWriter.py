from Writer import Writer
from xlwt import easyxf, Workbook


# Implementation specifics for xlwt
class XLWTWriter(Writer):
    def __init__(self, filename):
        # for Python2 support; instead of using just super() 
        super(XLWTWriter, self).__init__(filename)
        self.excel_doc = Workbook()

    def new_sheet(self, sheet_name):
        return self.excel_doc.add_sheet(sheet_name)

    def write_cell(self, sheet_name, row_no, col_no, val, bold=False):
        sheet = self.get_sheet(sheet_name)

        if bold:
            style_bold = easyxf('font:bold True')
            sheet.write(row_no, col_no, val, style_bold) 
        else:
            sheet.write(row_no, col_no, val)

    # Set a specific column's width across all sheets.
    def set_column_width(self, col_no, width):
        for sheet_name in self.sheets:
            self.sheets[sheet_name].col(col_no).width = width

    def save_to_file(self, filename):
        self.excel_doc.save(filename) 
