import xlrd
import os
import traceback

#----------------------------------------------------------------------
def search_key_in_file(first_sheet,key):
    """
    Open ic require file with part number
    """
    #print "key = " + key
    for i in range(first_sheet.nrows):
        row = first_sheet.row_values(i)
        for j in range(len(row)):
            if row[j] == key:
                print "Matched at row[%d]" %(i)
#----------------------------------------------------------------------
def start():
    for file in os.listdir("list"):
        if file.endswith(".xlsx") | file.endswith(".xls"):
            print "Excel Files = %s" %(file)
            # Get file sheet
            book = xlrd.open_workbook("list/" + file)
            # get the first worksheet
            key_sheet = book.sheet_by_index(0)
            search_key_in_file(key_sheet,"XQ16D8E2GM-K-BC")

#----------------------------------------------------------------------
if __name__ == "__main__":
    start()
