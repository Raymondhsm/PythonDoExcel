from os import listdir,system
from openpyxl import load_workbook,Workbook

from error import ErrorList,Error,Warning,processException
import utils
import saleAmount
import refund


def __doSA(path):
    _option = input("是否要做退款录入，是就输入点东西，否就啥也不输：")

    
    




def __doPF():
    pass


def doGuide(path = "./"):
    option = input("输入点东西就进入更新汇总表功能\n什么都不输就进入更新成本功能：")
    if option == "":
        __doSA(path)
    else :
        __doPF()





    # files = listdir(path)

    # open report to write down

