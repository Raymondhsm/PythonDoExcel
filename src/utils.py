from error import ErrorList, Error, Warning, processException
from os import path as osPath

def is_contains_chinese(strs):
    for _char in strs:
        if '\u4e00' <= _char <= '\u9fa5':
            return True
    return False


def getReportPath(path, _files, _tips, onlyXlsx = True, canSkip = False):
    _count = 0
    _pathList = []

    _isXls = False

    # print file list
    for _file in _files:
        # 去除文件夹，非.xlsx文件以及临时文件
        if (not osPath.isdir(path + '/' + _file)) and "~$" not in _file:
            if ".xlsx" in _file or (not onlyXlsx and ".xls" in _file):
                _count += 1
                _pathList.append(path + '/' + _file)
                print(str(_count) + "、" + _file)

        
        # only at the onlyxlsx mode, will send the warning
        if onlyXlsx and ".xls" in _file and ".xlsx" not in _file :
            _isXls = True

    # do not find file
    if _count == 0:
        ErrorList.addError(Error("./", "do not find files"))
        print("do not find files")
        return None

    # warning find .xls file
    if _isXls:
        ErrorList.addError(Warning("./", "We have found .xls file in report list"))
        print("We have found .xls file, but the software can not read the .xls as report table!")
        print("Please convert it into .xlsx file, if you want to read it!!!")

    # input file number
    _index = 0
    _tipStr = "input number to choose {} file{}:".format(_tips, "(zero for skipping)" if canSkip else "")
    
    while _index <= 0 or _index > _count:
        _input = input(_tipStr)
        if _input == "" or not _input.isdigit():
            continue
        
        _index = int(_input)
        if canSkip and _index == 0 :
            return None

        if _index < 0 or _index > _count:
            print("wrong number\n")
    
    # output \n
    print("\n")
    
    return _pathList[_index - 1]


def findPlatformIndex(_reportTable, _platform):
    try:
        # store the merge cell's info
        mergeList = _reportTable.merged_cells
        mergeDict = {}
        for mergeCell in mergeList:
            mergeDict[mergeCell.min_row] = mergeCell.max_row - mergeCell.min_row

        # to find the index of platform
        isFindPlatform = False
        index = 1
        while index < _reportTable.max_row:
            _reportPlatform = _reportTable["C" + str(index)].value
            if _reportPlatform is None:
                index += 1
                continue
            _reportPlatform = _reportPlatform if not _reportPlatform.strip().isalpha() else _reportPlatform.lower()
            if _reportPlatform.strip() == _platform.strip().lower():
                isFindPlatform = True
                break
            else :
                if mergeDict.get(index) != None:
                    index += mergeDict.get(index) + 1
                else :
                    index +=1

        offset = mergeDict[index] + 1 if index in mergeDict else 1
    
    except Exception:
        processException()
        return -1, 0

    else:
        if isFindPlatform:
            return index, offset
        else:
            return -1, 0