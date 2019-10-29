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
        
