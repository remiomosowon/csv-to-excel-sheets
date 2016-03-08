from xlwt import *
import glob

for inputfilename in glob.glob('*.csv'):
    input_filename = inputfilename

def addtodict(dict, itemid, row_values):
    if dict.has_key(itemid):
        dict[itemid].append(row_values)
    else:
        dict[itemid] = [row_values]

dict = {}
titlerow = []
toprows = []
csvfile = file(input_filename, 'r')
KEY_COLUMN = 'ITEM_ID'

i = 0
for line in csvfile.readlines():
    if i < 8:
        toprows.append(line)
    elif i == 8:
        pass
    elif i == 9:
        titlerow.append(line)
        col_hdrs = [hdr.strip().upper() for hdr in line.split(';')]
        key_col_no = col_hdrs.index(KEY_COLUMN)
    else:
        itemid = line.split(';')[key_col_no].strip()
        addtodict(dict, itemid, line)
    i = i + 1


book = Workbook()  # setup a new workbook (i.e. a new excel document)
styleBold = easyxf('font:bold True')
styleBoldAndLarge = easyxf('font:bold True, height 400')

for itemid in dict.keys():
    itemid_sheet = book.add_sheet(str(itemid))
    startrow = 0
    titlerownum = 0

    itemid_sheet.write(startrow, 3, itemid, styleBoldAndLarge)
    startrow = startrow + 2
    
    for line in toprows:
        if len(line.split(';')) > 1:
            key = line.split(';')[0].strip()
            value = line.split(';')[1].strip()
            itemid_sheet.write(startrow, 0, key, styleBold)
            itemid_sheet.write(startrow, 1, value)
        startrow = startrow + 1

    titlerownum = startrow + 1
    
    itemid_sheet.write(titlerownum,0,'INVOICE', styleBold)
    itemid_sheet.write(titlerownum,1,'REFNO', styleBold)
    itemid_sheet.write(titlerownum,2,'CUSTNO1', styleBold)
    itemid_sheet.write(titlerownum,3,'CUSTNO2', styleBold)
    itemid_sheet.write(titlerownum,4,'QTY', styleBold)
    itemid_sheet.write(titlerownum,5,'ORDERNO', styleBold)
    itemid_sheet.write(titlerownum,6,'ORDERNO2', styleBold)
    itemid_sheet.write(titlerownum,7,'ITEM', styleBold)
    itemid_sheet.write(titlerownum,8,'CASE', styleBold)

    i = titlerownum + 1
    for line in dict[itemid]:
        row = itemid_sheet.row(i)
        rowvalues = line.split(';')
        row.write(0, rowvalues[0].strip())  
        row.write(1, rowvalues[1].strip()) 
        row.write(2, rowvalues[2].strip())
        row.write(3, rowvalues[3].strip())
        row.write(4, rowvalues[4].strip())
        row.write(5, rowvalues[5].strip())
        row.write(6, rowvalues[6].strip())
        row.write(7, rowvalues[7].strip())
        row.write(8, rowvalues[8].strip())
        i = i + 1

    # specify a width of 4000 for both columns
    itemid_sheet.col(0).width = 5600
    itemid_sheet.col(1).width = 8000
    itemid_sheet.col(2).width = 6000
    itemid_sheet.col(3).width = 7600
    itemid_sheet.col(4).width = 2400
    itemid_sheet.col(5).width = 6400
    itemid_sheet.col(6).width = 6400
    itemid_sheet.col(7).width = 3200
    itemid_sheet.col(8).width = 3200

book.save(input_filename.strip(".csv") + '_output.xls')
