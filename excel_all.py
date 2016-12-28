import xlrd
#----------------------------------------------------------------------
def open_file(path):
    """
    Open and read an Excel file
    """
    book = xlrd.open_workbook(path)
    # print number of sheets
    print book.nsheets
    # print sheet names
    print book.sheet_names()
    # get the first worksheet
    first_sheet = book.sheet_by_index(0)
    # read a row
    print first_sheet.row_values(0)
    # read a cell
    cell = first_sheet.cell(1,0)
    print cell
    print cell.value
    # read a row slice
    print first_sheet.row_slice(rowx=0,
                                start_colx=0,
                                end_colx=2)
    i = 0
    j = 0
    nrows = first_sheet.nrows
    ncols = first_sheet.ncols
    for i in range(nrows):
        for j in range(ncols):
            if j == 0:
                print ('[%s]' % first_sheet.cell(i,0).value),
            else:
                print first_sheet.cell(i,j).value,
            print " ",
        j = 0
        print " "

      #print first_sheet.row_values(i)
#----------------------------------------------------------------------
if __name__ == "__main__":
    path = "list/FCST_Flex. MA_Juniper_0323_carol.xlsx"
    open_file(path)