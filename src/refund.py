from error import processException, ErrorList, Error, NotFound
from log import Logger
import utils
from xlrd import open_workbook


class refundInfo:
    def __init__(self, _platform):
        self.accountList = []
        self.locationList = []
        self.refundList = []
        self.length = 0
        self.platform = _platform

    def add(self, _account, _location, _refund):
        self.accountList.append(_account)
        self.locationList.append(_location)
        self.refundList.append(_refund)
        self.length += 1

    def get(self, index):
        return self.accountList[index], self.locationList[index], self.refundList[index]


class Refund:
    
    def __init__(self, _summary, _notfoundTable):
        Logger.addLog("CREATE refund!!")
        self.summary = _summary
        self.notfoundTable = _notfoundTable
        self.__correctCount = 0
        self.__failCount = 0

    
    def processRefundInfo(self, _filePath):

        try:
            Logger.addLog("OPEN REFUND FILE!! Path = " + _filePath)
            _refundBook = open_workbook(_filePath,formatting_info=True)
            _refundSheets = _refundBook.sheets()

        except Exception:
            processException()

        for _refundsheet in _refundSheets:
            _infoType = _refundsheet.name

            print("正在处理 " + _infoType + " 退款金额...")
            Logger.addLog("PROCESS {} 退款".format(_infoType))

            _refundInfoList = self.__getRefundInfo(_refundsheet)
            self.__setRefundInfo(_infoType, _refundInfoList)


    def __getRefundInfo(self, _refundTable):
        try:
            _refundInfoList = []

            # use the merge cell to locate the useful cell
            for merge in _refundTable.merged_cells:
                rs, re, cs, ce = merge

                # ignore some merge cell
                if re-rs != 1 and ce-cs != 2:
                    continue

                # ignore when it has no info
                if _refundTable.cell_value(re,cs) == "":
                    continue
                
                # read the platform
                _platform = _refundTable.cell_value(rs,cs).strip().lower()
                if utils.is_contains_chinese(_platform):
                    ErrorList.addError(Error("REFUND ERROR! Type = " + _refundTable.name + "\tPlatform: " + _platform, "We can not identify the chinese as platform name"))
                    continue
                else:
                    _refundInfoInstance = refundInfo(_platform)

                _row = re + 1 if _refundTable.cell_value(re,cs) == "账号" else re
                while True: 
                    _acclo = _refundTable.cell_value(_row,cs)

                    if _acclo == "":
                        break
                    else:
                        # split the account and location
                        if '(' in _acclo:
                            _accloList = _acclo.split("(",1)
                            _account = _accloList[0].strip()
                            _location = _accloList[1].split(')',1)[0].strip()
                        elif '（' in _acclo:
                            _accloList = _acclo.split("（",1)
                            _account = _accloList[0].strip()
                            _location = _accloList[1].split('）',1)[0].strip()
                        else:
                            _acclo = _acclo.strip()
                            _accloList = _acclo.split(" ",1)
                            _account = _accloList[0].strip()

                            # set the default value
                            if len(_accloList) == 1:
                                _location = "CN"
                            else:
                                _location = _accloList[1].strip()
                        
                        # read the refund
                        _refund = _refundTable.cell_value(_row,cs+1)
                        
                        _refundInfoInstance.add(_account, _location, _refund)
                    
                    # add 1 to row
                    _row += 1
                
                # add to the list
                _refundInfoList.append(_refundInfoInstance)
        except Exception:
            processException()

        return _refundInfoList


    def __setRefundInfo(self, _infoType,  _refundInfoList):
        try:
            # if can not find correct type, return
            if _infoType not in self.summary.sheetnames:
                ErrorList.addError(Error("REFUND ERROR","can not find correct sheet: " + _infoType))
                return 
            _refundTable = self.summary[_infoType]

            for _refundInstance in _refundInfoList:
                _index, _offset = utils.findPlatformIndex(_refundTable, _refundInstance.platform)
                
                # if can not find platform, send error
                if _index == -1:
                    ErrorList.addError(Error("REFUND ERROR","can not find correct platform: " + _refundInstance.platform))
                    continue
                
                # create the dictionary 
                accountDict = {}
                for _row in range(_index, _index + _offset):
                    _account = _refundTable["D" + str(_row)].value.strip()
                    _locationList = _refundTable["F" + str(_row)].value.split(' ')
                    _location = _locationList[1].strip() if len(_locationList) > 1 else _locationList[0].strip()

                    # add to the dict
                    accountDict[_account + '_' + _location] = _row

                # process the refund
                for _refundIndex in range(0, _refundInstance.length):
                    _refundAccount, _refundLocation, _refund = _refundInstance.get(_refundIndex)

                    if _refundAccount + '_' + _refundLocation in accountDict:
                        self.__correctCount += 1
                        _refundTable["I" + str(accountDict[_refundAccount + '_' + _refundLocation])].value = _refund
                    else:
                        self.__processRefundNotFound(_infoType, _refundInstance.platform, _refundAccount, _refundLocation, _refund)
        except Exception:
            processException()

        return


    def __processRefundNotFound(self, _infoType, _platform, _refundAccount, _refundLocation, _refund):

        try:
            # find the account and location in the notfound table
            for _row in range(2, self.notfoundTable.max_row + 1):
                _nfType = self.notfoundTable["A" + str(_row)].value
                _nfPlatform = self.notfoundTable["B" + str(_row)].value
                _nfAccount = self.notfoundTable["C" + str(_row)].value
                _nfLocation = self.notfoundTable["E" + str(_row)].value.split(' ')[1]

                if _infoType == _nfType and _platform == _nfPlatform and _refundAccount == _nfAccount and _refundLocation == _nfLocation:
                    self.notfoundTable["H" + str(_row)] = _refund
                    return

        except Exception:
            processException()
            return


        # if do not be found, we insert and send warning
        _row = self.notfoundTable.max_row + 1
        self.notfoundTable["A" + str(_row)].value = _infoType
        self.notfoundTable["B" + str(_row)].value = _platform
        self.notfoundTable["C" + str(_row)].value = _refundAccount
        self.notfoundTable["E" + str(_row)].value = "unkown " + _refundLocation
        self.notfoundTable["I" + str(_row)].value = _refund

        # count the failcount and send the error
        self.__failCount += 1
        ErrorList.addError(NotFound("REFUND ERROR!\tType = " + _infoType + "\tPlatform: " + _platform, "Can not match the refund account and location!!"))


    def getFailCount(self):
        return self.__failCount

    def getCorrectCount(self):
        return self.__correctCount