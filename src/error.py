import sys
import linecache
from os import system


class ErrorList:
    __errorList = []
    __warningList = []
    __notFoundList = []

    @classmethod
    def printErrorList(ErrorList):
        if len(ErrorList.__errorList) != 0 :
            print("\n操作过程有以下错误：")
        else:
            print("\n恭喜！！操作过程无错误")
        for _error in ErrorList.__errorList:
            _error.printInfo()

        if len(ErrorList.__warningList) != 0 :
            print("\n操作过程有以下警告：")
        else:
            print("\n恭喜！！操作过程无警告")
        for _warning in ErrorList.__warningList:
            _warning.printInfo()

        if len(ErrorList.__notFoundList) != 0 : print("\n存在新增数据，请自行插入，数据已录入 notfound.xlsx")
        for _notfound in ErrorList.__notFoundList:
            _notfound.printInfo()
        

    @classmethod
    def addError(ErrorList, _error):
        if type(_error) == Error:
            ErrorList.__errorList.append(_error)

        elif type(_error) == Warning:
            ErrorList.__warningList.append(_error)

        elif type(_error) == NotFound:
            ErrorList.__notFoundList.append(_error)
    
    @classmethod
    def getErrorCount(ErrorList):
        return len(ErrorList.__errorList)

    @classmethod
    def getWarningCount(ErrorList):
        return len(ErrorList.__warningList) + (1 if len(ErrorList.__notFoundList) > 0 else 0)


class ErrorBase:
    TYPE_ERROR = 1
    TYPE_WARNING = 2
    TYPE_NOTFOUND = 3

    def __init__(self,_path,_message):
        self.path = _path
        self.message = _message

    def printInfo(self):
        pass


class Error(ErrorBase):
    def __init__(self,_path,_message):
        self.errorType = ErrorBase.TYPE_ERROR
        ErrorBase.__init__(self, _path, _message)


    def printInfo(self):
        print("ERROR: " + self.message)
        print("PATH: " + self.path + "\n")


class Warning(ErrorBase):
    def __init__(self,_path,_message):
        self.errorType = ErrorBase.TYPE_WARNING
        ErrorBase.__init__(self, _path, _message)

    def printInfo(self):
        print("Warning: " + self.message)
        print("PATH: " + self.path + "\n")


class NotFound(ErrorBase):
    def __init__(self,_path,_message):
        self.errorType = ErrorBase.TYPE_NOTFOUND
        ErrorBase.__init__(self, _path, _message)

    def printInfo(self):
        print("PATH: " + self.path)


def processException():
    exc_type, exc_obj, tb = sys.exc_info()
    f = tb.tb_frame
    lineno = tb.tb_lineno
    filename = f.f_code.co_filename
    linecache.checkcache(filename)
    line = linecache.getline(filename, lineno, f.f_globals)

    # send error
    ErrorList.addError(Error(filename,'APPLICATION EXCEPTION (LINE {} "{}"): {}'.format(lineno, line.strip(), exc_obj)))
    print('APPLICATION EXCEPTION (LINE {} "{}"): {}'.format(lineno, line.strip(), exc_obj))

    # stop the system and exit
    system("pause")
    exit() 