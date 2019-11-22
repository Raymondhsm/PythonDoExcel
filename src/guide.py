from os import listdir,system
from openpyxl import load_workbook,Workbook

from error import ErrorList,Error,Warning,processException
import utils
import saleAmount
import refund
import profit
import notfound


def __doSA(path = "./"):
    _option = input("是否要做退款录入，是就输入点东西，否就啥也不输：")
    _doRefund = False if _option == "" else True

    _summaryPath = utils.getReportPath(path, "选择汇总文件")
    try:
        _summary = load_workbook(_summaryPath)
    except:
        processException()
    
    # 获取notfound表
    NF = notfound.NotFound()
    _notfoundTable = NF.getNotfoundTable()

    # 处理销售额
    SA = saleAmount.SaleAmount(_summary,_notfoundTable)
    SA.processDir(path)

    # 处理退款
    if _doRefund:
        _refundPath = utils.getReportPath(path,"选择退款文件",False)

        RF = refund.Refund(_summary, _notfoundTable)
        RF.processRefundInfo(_refundPath)

    # 保存文件
    _summaryName = input("输点什么东西当输出文件名呗：")
    if _summaryName == "":
        _summary.save("summary.xlsx")
    else:
        _summary.save(_summaryName)

    NF.save()


def __doPF(path = "./"):
    _originPath = utils.getReportPath(path,"选择参考成本文件",False)
    _updatePath = utils.getReportPath(path,"选择更新成本文件")

    PF = profit.Profit(_originPath, _updatePath)
    PF.processProfitUpdate()
    pass


def doGuide(path = "./"):
    option = input("输入点东西就进入更新汇总表功能\n什么都不输就进入更新成本功能：")
    if option == "":
        __doSA(path)
    else :
        __doPF(path)


