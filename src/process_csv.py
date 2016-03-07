import glob
from XLWTWriter import XLWTWriter

CSV_DELIMITER = ';'                 # Specific delimiter used to delimit the input CSV data
KEY_COLUMN_HDR_NAME = 'ITEM_ID'     # Name of column header used to filter the data

if __name__ == "__main__":
    input_csv_filename = glob.glob('*.csv').pop()   # Name of the only csv file in current directory.
    writer = XLWTWriter(input_csv_filename)         # Excel writer object (ADT implementation)
    header_data = []                                # Header data (description) to be repeated across generated sheets.
    sheet_has_headers = {}                          # Table keeping track of which sheets have header data.
    sheet_current_row = {}

    def has_headers(sheet_name):
        has_header = sheet_has_headers.get(sheet_name, False)
        if not has_header:
            sheet_has_headers[sheet_name] = True
        return has_header

    def current_sheet_row(sheet_name):
        current_row = sheet_current_row.get(sheet_name, 0)
        if current_row == 0:
            sheet_current_row[sheet_name] = 0
        sheet_current_row[sheet_name] += 1
        return current_row

    with open(input_csv_filename) as csv_file:
        # First copy the rows containing the data description.
        while True:
            line = csv_file.readline()
            header_data.append(line)

            if line.strip() == "":
                break

        # Process the row with the csv data's column headers.
        headers_row = csv_file.readline()
        headers = [header.strip().upper() for header in headers_row.split(CSV_DELIMITER)]

        key_col_no = headers.index(KEY_COLUMN_HDR_NAME.upper())

        # Then process the rest of the data.
        for line in csv_file:
            key_val = line.split(CSV_DELIMITER)[key_col_no].strip()

            sheet_name = key_val

            # If this sheet doesn't have any headers written to it because it hasn't been processed.
            if not has_headers(sheet_name):
                # Write the descriptive headers to the sheet.
                for header_data_line in header_data:
                    header_data_list = [data.strip() for data in header_data_line.split(CSV_DELIMITER)]
                    writer.write_row(sheet_name, current_sheet_row(sheet_name), header_data_list)

                # Write the data column headers.
                writer.write_row(sheet_name, current_sheet_row(sheet_name), headers, bold=True)

            row_data = [data.strip() for data in line.split(CSV_DELIMITER)]
            writer.write_row(sheet_name, current_sheet_row(sheet_name), row_data)

        # Adjust column widths.
        col_widths = [5600,8000,6000,7600,2400,6400,6400,3200]
        writer.set_column_widths(col_widths)

        # Save Excel file.
        writer.save_to_file(input_csv_filename.strip(".csv") + '_output.xls')
