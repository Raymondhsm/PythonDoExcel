import xlrd
import xlwt

# define data class
class info:
    account = "undefined"
    name = "undefined"
    location = "undefined"
    salesAmount = 0.0
    profitRate = 0.0
    normal = 1      # to mark the cell is nornal or not



def getInfo(filePath):
    # open the file
    data = xlrd.open_workbook(filePath)

    # to store this file data which _platform and type is
    _platform = data.sheet_by_index(0).cell_value(0,1)
    _infoType = data.sheet_by_index(0).cell_value(1,1).split('/',2)[1]
    if _infoType == "C类" : _infoType += ("打底裤")

    # use info list to temporary store data 
    _infoList = []

    # for each table to process data
    sheets = data.sheets()
    for table in sheets:
        _row = 0
        nrow = table.max_row

        # for each module to find out useful info
        while _row < nrow:
            # add 1 to _row to ignore the title
            _row += 1

            # create the module info instance
            _infoInstance = info()

            _infoInstance.name = table.cell(_row,1).value.split('/',2)[0]
            if _infoInstance.name == "":
                _infoInstance.name = '/'

            # to process the account/location cell
            row_3 = table.cell(_row,2).value
            _infoInstance.account = row_3.split('/')[0]
            # remove the charactor "仓"
            _infoInstance.location = row_3.split('/')[1][0:-1]       

            # to judge this module is normal or not
            row_4 = table.cell(_row,3).value
            if row_4 == None:
                _infoInstance.normal = False
                _infoInstance.salesAmount = table.cell(_row,4).value
            else :
                _infoInstance.salesAmount = row_4
                _infoInstance.profitRate = table.cell(_row+1,4).value

            # append the module instance into the list
            _infoList.append(_infoInstance)

            # add 3 to variable _row to move to next module 
            _row += 3

    data.release_resources()
    return _platform, _infoType, _infoList







# open the file
xlrd.Book.encoding = "utf8"
data = xlrd.open_workbook("A.xlsx")

# for each file to process data 


# to store this file data which platform and type is
platform = "wish"
infoType = "A类"

# use info list to temporary store data 
infoList = []

# for each table to process data
sheets = data.sheets()
for table in sheets:
    row = 0
    nrow = table.nrows

    # for each module to find out useful info
    while row < nrow:
        # add 1 to row to ignore the title
        row += 1

        # create the module info instance
        infoInstance = info()

        infoInstance.name = table.cell_value(row,1).split('/',2)[0]

        # to process the account/location cell
        row_2 = table.cell_value(row,2)
        infoInstance.account = row_2.split('/')[0]
        # remove the charactor "仓"
        infoInstance.location = row_2.split('/')[1][0:-1]       

        # to judge this module is normal or not
        row_3 = table.cell_value(row,3)
        if row_3 == '':
            infoInstance.normal = 0
            infoInstance.salesAmount = table.cell_value(row,4)
        else :
            infoInstance.salesAmount = table.cell_value(row,3)
            infoInstance.profitRate = table.cell_value(row+1,4)

        # append the module instance into the list
        infoList.append(infoInstance)

        # add 3 to variable row to move to next module 
        row += 3


# open report to write down
# summary = xlwt.w

# for each data in infoList to write down in the report
# for infoInstance in infoList:
        
