import os
import datetime
from openpyxl import Workbook, load_workbook

# define data class
class info:
    account = "undefined"
    name = "undefined"
    location = "undefined"
    salesAmount = 0.0
    profitRate = 0.0
    normal = True      # to mark the cell is nornal or not 


def getInfo(filePath):
    # open the file
    data = load_workbook(filePath, data_only=True)

    # to store this file data which _platform and type is
    _platform = data[data.sheetnames[0]]["A2"].value
    _infoType = data[data.sheetnames[0]]["B2"].value.split('/',2)[1]
    if _infoType == "C类" : _infoType += ("打底裤")

    # use info list to temporary store data 
    _infoList = []

    # for each table to process data
    sheets = data._sheets
    for table in sheets:
        _row = 1
        nrow = table.max_row

        # for each module to find out useful info
        while _row < nrow:
            # add 1 to _row to ignore the title
            _row += 1

            # create the module info instance
            _infoInstance = info()

            _infoInstance.name = table.cell(_row,2).value.split('/',2)[0]
            if _infoInstance.name == "":
                _infoInstance.name = '/'

            # to process the account/location cell
            row_3 = table.cell(_row,3).value
            _infoInstance.account = row_3.split('/')[0]
            # remove the charactor "仓"
            _infoInstance.location = row_3.split('/')[1][0:-1]       

            # to judge this module is normal or not
            row_4 = table.cell(_row,4).value
            if row_4 == None:
                _infoInstance.normal = False
                _infoInstance.salesAmount = table.cell(_row,5).value
            else :
                _infoInstance.salesAmount = row_4
                _infoInstance.profitRate = table.cell(_row+1,5).value

            # append the module instance into the list
            _infoList.append(_infoInstance)

            # add 3 to variable _row to move to next module 
            _row += 3

    data.close()
    return _platform, _infoType, _infoList


def setInfo(_platform, _infoType, _infoList):
    # find the infotype table
    reportTable = summary[_infoType]

    # store the merge cell's info
    mergeList = reportTable.merged_cells
    mergeDict = {}
    for mergeCell in mergeList:
        mergeDict[mergeCell.min_row] = mergeCell.max_row - mergeCell.min_row

    # to find the index of platform
    index = 1
    while index < reportTable.max_row:
        if reportTable["C" + str(index)].value.lower() == _platform:
            break
        else :
            if mergeDict.get(index) != None:
                index += mergeDict.get(index) + 1
            else :
                index +=1

    # for each data in infoList to write down in the report
    for infoInstance in infoList:
        isFind = False      # to mark the account is finded or not

        for row in range(index, index + mergeDict[index]):
            reportAccount = reportTable["D" + str(row)].value
            reportLocationList = reportTable["F" + str(row)].value.split(' ')
            reportLocation = reportLocationList[1] if len(reportLocationList) > 1 else reportLocationList[0]

            # print(infoInstance.account + "\t" + reportAccount + "\t" + infoInstance.location + "\t" + reportLocation)
            # to match corret account and location row
            if infoInstance.account == reportAccount and infoInstance.location == reportLocation :
                isFind = True

                # if the name is wrong, then change the name 
                reportName = reportTable["E" + str(row)].value
                if infoInstance.name !=  reportName : 
                    reportTable["E" + str(row)].value = infoInstance.name
                    reportTable["F" + str(row)].value = infoInstance.name + " " + reportLocation

                # to judge the normal is true or not
                if infoInstance.normal :
                    # if normal, then write down the salesAmount and profitRate
                    reportTable["G" + str(row)].value = infoInstance.salesAmount
                    reportTable["H" + str(row)].value = infoInstance.profitRate
                else :
                    # else write down the margin 
                    reportTable["J" + str(row)].value = infoInstance.salesAmount

                # print(row)

                # write down and break the for loop
                break

        if isFind:
            continue

        # 由于有合并表格的存在，插入一行真的极其的烦，功能后面在迭代吧，我不行了
        # if do not find the account in the table, then create it
        # reportTable.insert_rows(index)
        # reportTable["D" + str(index)].value = infoInstance.account
        # reportTable["E" + str(index)].value = infoInstance.name
        # reportTable["F" + str(index)].value = infoInstance.name + " " + infoInstance.location
        #  # to judge the normal is true or not
        # if infoInstance.normal :
        #     # if normal, then write down the salesAmount and profitRate
        #     reportTable["G" + str(index)].value = infoInstance.salesAmount
        #     reportTable["H" + str(index)].value = infoInstance.profitRate
        # else :
        #     # else write down the margin 
        #     reportTable["J" + str(index)].value = infoInstance.salesAmount

        # if do not find the account in the table, then print
        notfound = "this account not found: \n platform: " + _platform + "\n type: " + _infoType + "\n name: " + infoInstance.name + "\n location: " + infoInstance.location
        print(notfound)

    summary.save("output.xlsx")



# --------------------------------------------------------------------------------------------- #
# main function #
# --------------------------------------------------------------------------------------------- #

path = "input" #文件夹目录

# open report to write down
summary = load_workbook("B.xlsx")

#得到文件夹下的所有文件名称
files= os.listdir(path) 
print(files)

#遍历文件夹
for file in files: 
    #判断是否是文件夹
    upperPath = path + '/' + file
    if os.path.isdir(upperPath): 
        reportFiles = os.listdir(upperPath)

        # to process each file
        for reportFile in reportFiles:
            # ignore the dir
            lastPath = upperPath + '/' + reportFile
            if os.path.isdir(lastPath) :
                continue

            # ignore the files which not .xlsx file
            if not (".xlsx" in reportFile):

                # if find .xls file, print tips
                if ".xls" in reportFile:
                    print("warning: find .xls file in dir, please convert to .xlsx file!!!")
                continue
            else:
                # ignore the temporary files
                if "~$" in reportFile:
                    print("find \'~$\' in the file name, do use \'~$\' for file name in case we see it as temporary files")
                    continue
                
                print("正在处理文件：" + lastPath)
                startTime = datetime.datetime.now()

                platform, infoType, infoList = getInfo(lastPath)

                endTime = datetime.datetime.now()
                interval = (endTime-startTime).seconds
                print("文件已读取完成，用时 " + str(interval) + " 秒")

                setInfo(platform, infoType, infoList)

                endTime = datetime.datetime.now()
                interval = (endTime-startTime).seconds
                print("文件已处理完成，用时 " + str(interval) + " 秒\n")
