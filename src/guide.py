from os import listdir,system
from openpyxl import load_workbook,Workbook

from error import ErrorList,Error,Warning,processException
from log import Logger
import utils
import saleAmount
import refund
import profit
import notfound


def __doSA(path = "./"):
    _option = input("是否要做金额录入，是就输入点东西，否就啥也不输：")
    _doSA = False if _option == "" else True
    Logger.addLog("输入：{}，退款{}".format(_option, _doSA))

    _option = input("是否要做退款录入，是就输入点东西，否就啥也不输：")
    _doRefund = False if _option == "" else True
    Logger.addLog("输入：{}，退款{}".format(_option, _doRefund))

    _summaryPath = utils.getReportPath(path, "summary")
    try:
        _summary = load_workbook(_summaryPath)
    except:
        processException()
    
    # 获取notfound表
    NF = notfound.NotFound()
    _notfoundTable = NF.getNotfoundTable()

    # 处理销售额
    if _doSA:
        SA = saleAmount.SaleAmount(_summary,_notfoundTable)
        SA.processDir(path)

    # 处理退款
    if _doRefund:
        _refundPath = utils.getReportPath(path,"refund",onlyXlsx=False, canSkip=True)

        if _refundPath is not None:
            RF = refund.Refund(_summary, _notfoundTable)
            RF.processRefundInfo(_refundPath)

    # 保存文件
    _summaryName = input("输点什么东西当输出文件名呗：")
    if _summaryName == "":
        _summary.save("summary.xlsx")
    else:
        _summary.save(_summaryName)
    Logger.addPrefabLog(Logger.LOG_TYPE_SAVE,_summaryName)

    NF.save()


def __doPF(path = "./"):
    _originPath = utils.getReportPath(path,"originfile",False)
    _updatePath = utils.getReportPath(path,"updatefile")

    PF = profit.Profit(_originPath, _updatePath)
    PF.processProfitUpdate()
    PF.save()
    pass


def doGuide(path = "./"):
    option = input("输入点东西就进入更新成本功能\n什么都不输就进入更新汇总表功能：")
    if option == "":
        Logger.addLog("输入：{}，处理saleAmount。".format(option))
        __doSA(path)
    else :
        Logger.addLog("输入：{}，处理更新成本".format(option))
        __doPF(path)


