from datetime import datetime

class Log:
    def __init__(self, message, date = datetime.now()):
        self.date = date
        self.message = message

    def printLog(self):
        logstr = "[ {} ]: {}".format(self.date, self.message)
        print(logstr)


class Logger:

    __logList = []

    LOG_TYPE_DO = 1
    LOG_TYPE_SAVE = 2


    @classmethod
    def addLog(logger, message:str):
        logger.__logList.append(Log(message))

    @classmethod
    def addDateLog(logger, message:str, date):
        logger.__logList.append(Log(message, date))

    @classmethod
    def addPrefabLog(logger, logType, path = "" ):
        if logType == logger.LOG_TYPE_DO:
            logger.addLog("正在处理文件： " + path)
        
        elif logType == logger.LOG_TYPE_SAVE:
            logger.addLog("正在保存文件： " + path)

    @classmethod
    def printLog(logger):
        for log in logger.__logList:
            log.printLog()


