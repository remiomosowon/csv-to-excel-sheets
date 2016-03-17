# ADT to abstract out whichever Excel writer is used
class Writer(object):
    def __init__(self, filename):
        self.filename = filename
        self.sheets = {}

    def get_sheet(self, sheet_name):
        sheet_name = str(sheet_name)

        try:
            return self.sheets[sheet_name]
        except KeyError:
            new_sheet = self.new_sheet(sheet_name)
            self.sheets[sheet_name] = new_sheet
            return new_sheet

    def write_row(self, sheet_name, row_no, list_of_vals, bold=False):
        for col_no, val in enumerate(list_of_vals):
            self.write_cell(sheet_name, row_no, col_no, val, bold)

    # Set column widths across all sheets.
    def set_column_widths(self, col_widths):
        for col_no, width in enumerate(col_widths):
            self.set_column_width(col_no, width)

    # Create and return an Excel Sheet object
    def new_sheet(self, sheet_name):
        raise NotImplementedError

    def write_cell(self, sheet_name, row_no, col_no, val, bold=False):
        raise NotImplementedError

    # Set a specific column's width across all sheets.
    def set_column_width(self, col_no, width):
        raise NotImplementedError

    def save_to_file(self, filename):
        raise NotImplementedError
