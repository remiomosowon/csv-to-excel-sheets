from Writer import Writer


# Implementation specifics for OpenExcelWriter
class OpenExcelWriter(Writer):
    def __init__(self, filename):
        # for Python2 support; instead of using just super() 
        super(OpenExcelWriter, self).__init__(filename)

    def new_sheet(self, sheet_name):
        raise NotImplementedError

    def write_cell(self, sheet_name, row_no, col_no, val, bold=False):
        raise NotImplementedError

    # Set a specific column's width in all sheets.
    def set_column_width(self, col_no, width):
        raise NotImplementedError

    def save_to_file(self, filename):
        raise NotImplementedError
