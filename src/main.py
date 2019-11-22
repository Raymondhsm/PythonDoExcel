import command
from os import system


print("欢迎使用6.0版本的本系统了（连名字都没有，哎呀我去")
print("有不懂不要问，去命令行模式输入 do -help ")
print("更新成本功能暂时仅支持命令行模式\n")

mode = input("输入点东西就进入命令行模式，啥也不输就进入引导模式：")

if mode == "":
    print("引导模式还没开发！！！再见！！！")
    pass
    # guide.doGuide()
else:
    command.doCommand()




system("pause")