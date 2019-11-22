from error import processException, ErrorList, Error,Warning, NotFound
import utils
from datetime import datetime
from os import path as osPath
from os import listdir,system
from xlrd import open_workbook


class info:
    def __init__(self):
        self.account = "undefined"
        self.name = "undefined"
        self.location = "undefined"
        self.salesAmount = 0.0
        self.profitRate = 0.0
        self.normal = True      # to mark the cell is nornal or not 


class SaleAmount: 

    def __init__(self, _summary, _notfoundTable):
        self.__summary = _summary
        self.__notfoundTable = _notfoundTable
        self.__correctCount = 0
        self.__failCount = 0


    def processInfoWithTime(self,_lastpath):
        print("正在处理文件：" + _lastpath)
        _startTime = datetime.now()

        try:
            _platform, _infoType, _infoList = self.__getInfoByXlrd(_lastpath)
            _endTime = datetime.now()
            _interval = (_endTime-_startTime).seconds
            print("\r文件已读取完成，用时 " + str(_interval) + " 秒")

            # ignore writing when info list is none
            if _infoList == []:
                ErrorList.addError(Warning(_lastpath, "读取数据为空，注意检查"))
                print("读取数据为空，已跳过写入！！")
                return 

            _startTime = datetime.now()
            self.__setInfo(_platform, _infoType, _infoList, _lastpath)
            _endTime = datetime.now()
            _interval = (_endTime-_startTime).seconds
            print("数据已处理完成，用时 " + str(_interval) + " 秒\n")

        except Exception:
            processException()


    def processDir(self, _filepath):
        files = listdir(_filepath) 
        for file in files: 
            filePath = _filepath + '/' + file
            if osPath.isdir(filePath): 
                self.processSecondLevelDir(filePath)


    def processSecondLevelDir(self, _filepath):
        _files = listdir(_filepath) 

        for _file in _files:
            _lastpath = _filepath + '/' + _file

            #判断是否是文件夹
            if osPath.isdir(_lastpath): 
                # 递归搜索
                self.processDir(_lastpath)
            else:
                # to process each file
                
                # ignore the files which not .xlsx or .xls file
                if not (".xls" in _file):
                    continue
                else:
                    # ignore the temporary files
                    if "~$" in _file:
                        ErrorList.addError(Warning(_lastpath, "find \'~$\' in the file name, do use \'~$\' for file name in case we see it as temporary files"))
                        continue

                    self.processInfoWithTime(_lastpath)
   


    def __getInfoByXlrd(self, filePath):
        try:
            # open the file
            data = open_workbook(filePath)

            # to store this file data which _platform and type is
            _platform = data.sheet_by_index(0).cell_value(1,0).split('/')[0]
            _infoType = data.sheet_by_index(0).cell_value(1,1).split('/',2)[1]

            # use info list to temporary store data 
            _infoList = []

            # for each table to process data
            sheets = data.sheets()
            for table in sheets:
                # ignore the sheet of "原始"
                if table.name == "原始":
                    continue

                _row = 1
                nrow = table.nrows

                # for each module to find out useful info
                while _row < nrow:
                    # if find the none row, add 1 to row for finding the next
                    if table.cell(_row,1).value == "":
                        _row += 1
                        continue

                    # if account is none, the wo think the line is bad
                    row_3List = table.cell(_row,2).value.split('/')
                    if row_3List[0] == "" or utils.is_contains_chinese(row_3List[0]):
                        _row += 1
                        continue

                    # create the module info instance
                    _infoInstance = info()

                    # to process the account/location cell
                    _infoInstance.account = row_3List[0]
                    _infoInstance.location = row_3List[1] if not utils.is_contains_chinese(row_3List[1]) else row_3List[1][0:-1]

                    # if cell has no name, then wo think this line is bad 
                    _nameInTable = table.cell(_row,1).value.split('/',2)
                    _infoInstance.name = "/" if _nameInTable[0] == "" else _nameInTable[0]
                    if _infoType == "类": _infoType = _nameInTable[1]
            
                    # to judge this module is normal or not
                    row_4 = table.cell(_row,3).value
                    if row_4 == "":
                        _infoInstance.normal = False
                        _infoInstance.salesAmount = table.cell(_row,4).value
                    else :
                        _infoInstance.salesAmount = row_4
                        _infoInstance.profitRate = table.cell(_row+1,4).value

                    # append the module instance into the list
                    _infoList.append(_infoInstance)

                    # add 3 to variable _row to move to next module 
                    _row += 2

            data.release_resources()

        except Exception:
            processException()
            return "", "", []
        else :
            return _platform, _infoType, _infoList


    def __processNotFoundInfo(self, _platform, _infoType, _infoInstance):
        _row = self.__notfoundTable.max_row + 1
        self.__notfoundTable["A" + str(_row)] = _infoType
        self.__notfoundTable["B" + str(_row)] = _platform
        self.__notfoundTable["C" + str(_row)] = _infoInstance.account
        self.__notfoundTable["D" + str(_row)] = _infoInstance.name
        self.__notfoundTable["E" + str(_row)] = _infoInstance.name + " " + _infoInstance.location
        # to judge the normal is true or not
        if _infoInstance.normal :
            # if normal, then write down the salesAmount and profitRate
            self.__notfoundTable["F" + str(_row)].value = _infoInstance.salesAmount
            self.__notfoundTable["G" + str(_row)].value = _infoInstance.profitRate
        else :
            # else write down the margin 
            self.__notfoundTable["I" + str(_row)].value = _infoInstance.salesAmount



    def __setInfo(self, _platform, _infoType, _infoList, _path):
        try:
            # find the infotype table
            if _infoType not in self.__summary.sheetnames:
                ErrorList.addError(Error(_path,"can not find correct sheet: " + _infoType))
                return 
            reportTable = self.__summary[_infoType]

            # to find the index of platform
            index, offset = utils.findPlatformIndex(reportTable, _platform)
            if index == -1:
                ErrorList.addError(Error(_path, "can not find correct platform: " + _platform))
                return

            # for each data in infoList to write down in the report
            for infoInstance in _infoList:
                isFind = False      # to mark the account is finded or not

                for row in range(index, index + offset):
                    reportAccount = reportTable["D" + str(row)].value
                    reportLocationList = reportTable["F" + str(row)].value.split(' ')
                    reportLocation = reportLocationList[1] if len(reportLocationList) > 1 else reportLocationList[0]

                    # to match corret account and location row
                    if infoInstance.account == reportAccount and infoInstance.location == reportLocation :
                        isFind = True
                        self.__correctCount += 1

                        # if the name is wrong, then change the name 
                        reportName = reportTable["E" + str(row)].value
                        if infoInstance.name !=  reportName : 
                            reportTable["E" + str(row)].value = infoInstance.name
                            reportTable["F" + str(row)].value = infoInstance.name + " " + infoInstance.location

                        # to judge the normal is true or not
                        if infoInstance.normal :
                            # if normal, then write down the salesAmount and profitRate
                            reportTable["G" + str(row)].value = infoInstance.salesAmount
                            reportTable["H" + str(row)].value = infoInstance.profitRate
                        else :
                            # else write down the margin 
                            reportTable["J" + str(row)].value = infoInstance.salesAmount

                        # write down and break the for loop
                        break

                if isFind:
                    continue
                else:
                    self.__failCount += 1
                    ErrorList.addError(NotFound(_path, "存在新增数据，请自行插入，数据已录入 notfound.xlsx"))
                    self.__processNotFoundInfo(_platform, _infoType, infoInstance)

                # 由于有合并表格的存在，插入一行真的极其的烦，功能后面在迭代吧，我不行了
        except Exception:
            processException()

            
    def getFailCount(self):
        return self.__failCount

    def getCorrectCount(self):
        return self.__correctCount

