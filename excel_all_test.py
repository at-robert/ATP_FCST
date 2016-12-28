import xlrd
import Junipter
import Mitac
import Fabrinet
#----------------------------------------------------------------------
def open_file(path):
    """
    Open and read an Excel file
    """
    book = xlrd.open_workbook(path)
    # print number of sheets
    #print book.nsheets
    # print sheet names
    #print book.sheet_names()
    # get the first worksheet
    first_sheet = book.sheet_by_index(0)
    # read a row
    #print first_sheet.row_values(0)
    # read a cell
    cell = first_sheet.cell(1,0)
    #print cell
    #print cell.value
    # read a row slice
    #print first_sheet.row_slice(rowx=0,start_colx=0,end_colx=2)

    """
    if Junipter.search_junipter_rule(first_sheet,1) == 0:
        print "Juniper rule doesn't match"
    else:
        print "Juniper rule match"
    """

    """
    if Mitac.search_mitac_rule(first_sheet,1) == 0:
        print "Mitac rule doesn't match"
    else:
        print "Mitac rule match"
    """

    if Fabrinet.search_fabrinet_rule(first_sheet,3) == 0:
        print "fabrinet rule doesn't match"
    else:
        print "fabrinet rule match"
#----------------------------------------------------------------------
if __name__ == "__main__":
    #path = "list/FCST_Flex. MA_Juniper_0323_carol.xlsx"
    #path = "list/FCST_MiTAC_0321.xls"
    path = "list/FCST_Fabrinet_0318.xls"
    open_file(path)