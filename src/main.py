import command
import guide
from os import system


print("欢迎使用6.0版本的本系统了（连名字都没有，哎呀我去")
print("有不懂不要问，去命令行模式输入 do -help ")
print("汇总文件仅支持“.xlsx“格式, 退款文件仅支持”.xls“格式，出错了检查一下是不是这个问题")
print("更新成本功能暂时仅支持命令行模式\n")

mode = input("输入点东西就进入命令行模式，啥也不输就进入引导模式：")

if mode == "":
    guide.doGuide("./")
else:
    command.doCommand()




system("pause")