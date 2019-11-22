from error import ErrorList, Error, Warning, processException
import utils
import saleAmount
import refund

from os import listdir,system
from datetime import datetime
from time import sleep
from threading import Thread
from openpyxl import load_workbook,Workbook

def save(_fileName):
    print("\n正在保存文件")
    _fileName = "output" if _fileName == "" else _fileName

    try:
        summary.save(path + "/" + _fileName + ".xlsx")
        notfound.save(path + "/notfound.xlsx")

    except Exception:
        processException()
        print("文件保存失败，请检查 " + _fileName + ".xlsx 或 notfound.xlsx 是否在打开状态。请关闭文件后重试！！！")

    else:
        print("保存文件成功，路径：" + path + "/" + _fileName + ".xlsx") 


def get_user_input(user_input_ref):
    user_input_ref[0] = input("输入文件名（直接回车或 20 秒后将使用默认文件名保存）：")


def initNotFoundTable():
    notfoundTable["A1"] = "类型"
    notfoundTable["B1"] = "渠道"
    notfoundTable["C1"] = "账号"
    notfoundTable["D1"] = "姓名"
    notfoundTable["E1"] = "姓名/仓"
    notfoundTable["F1"] = "销售额"
    notfoundTable["G1"] = "利润率"
    notfoundTable["H1"] = "退款金额"
    notfoundTable["I1"] = "毛利"



def openReport(_reportPath):
    print("正在读取汇总文件\r")
    try:
        _summary = load_workbook(_reportPath)
    except Exception:
        print("读取汇总文件发生错误，可能汇总文件已被打开，请关闭文件后重试！！")
        processException()
        return None
    else:    
        print("汇总文件读取成功：" +  _reportPath + "\n")
        return _summary




# --------------------------------------------------------------------------------------------- #
# main function #
# --------------------------------------------------------------------------------------------- #


#文件夹目录
path = "../data"

#得到文件夹下的所有文件名称
files= listdir(path) 

# open report to write down
reportPath = utils.getReportPath(path, files, "summary")
refundPath = utils.getReportPath(path, files, "refund", onlyXlsx = False, canSkip = True)

# start timer
startTime = datetime.now()

if reportPath is not None:
    # open the report
    summary = openReport(reportPath)
    if summary is not None : 

        # init notfound table
        notfound = Workbook()
        notfoundTable = notfound.active
        initNotFoundTable()

        # instance the saleAmount and process
        SA = saleAmount.SaleAmount(summary, notfoundTable)
        SA.processDir(path)

        # instance the refund and process
        RF = refund.Refund(summary, notfoundTable)
        if refundPath is not None:
            RF.processRefundInfo(refundPath)

        # Declare a mutable object so that it can be pass via reference
        user_input = [None]

        mythread = Thread(target=get_user_input, args=(user_input,))
        mythread.daemon = True
        mythread.start()

        for increment in range(0, 21):
            sleep(1)
            if user_input[0] is not None:
                save(user_input[0])
                break
            
        if user_input[0] is None:
            save("output")


    endTime = datetime.now()
    interval = (endTime-startTime).seconds

    # print the error list
    ErrorList.printErrorList()

    print("\n已处理完成! 共耗时 " + str(interval) + " 秒")
    print("成功命中数据 " + str(SA.getCorrectCount() + RF.getCorrectCount()) + " 条，失败命中数据 " + str(SA.getFailCount() + RF.getFailCount()) + " 条。")
    print("出现错误 " + str(ErrorList.getErrorCount()) + " 项，警告 " + str(ErrorList.getWarningCount()) + " 项。")

# pause the os in case disappear
system("pause")



