# PythonDoExcel

利用Python的OpenXlPy库对Excel表格进行处理

> 这个程序就是一个python小白的糊弄之作啊。由于有一天我姐突然找我干excel枯燥烦闷的复制粘贴操作。做完一想，消灭这些无聊的操作不就是程序员存在的意义嘛。所以用python开干了，虽然自己就没写过python，但看网上python操作excel还不错。

## 1、环境准备

> 听信了网上的建议，我还是使用了anaconda软件来管理我的python环境。最后感觉还是挺舒心的。可以再软件可视化的看见自己的已安装的包。
    
> 项目用到的库基本在python的自带的库中，包括xlrd， openpyxl
    
> 但打包时使用pipenv进行创建虚拟环境，并使用pyinstaller来对python脚本打包成EXE文件。
    
> 这里还是提醒一下，注意一下要用这个程序的人的电脑是32位还是64位的，来选择一下使用python的版本。一开始我没注意，使用64位的python打包了exe，在我姐32位的电脑不兼容。最后32位和64位python环境共存也没搞好，就直接把64位卸载重装32位的了，难受。（有人会告诉我一声！！！

#### 

## 2、xlrd操作

> 看了晚上的一堆建议，然后一开始使用了`openpyxl`作为`excel`文件的读取和写入数据的库。但是后面写完才发现有点坑这个`openpyxl`，读取实在是有点慢，我一个行数不多，只有三十来K的文件，竟然在调试环境读取了一分多钟，运行环境也用了四十多秒。最后实在忍不了，投入`xlrd`的怀抱。
>     
> `xlrd`读得快，但可能内存消耗多一点，我没有仔细的比较。`openpyxl`读取较慢可能和我使用`cell`函数较多有关，导致`xml`文件多次解析。另外，`xlrd`还支持 `.xls` 和 `.xlsx` 文件，`openpyxl`仅支持`.xlsx`文件。

##### 

- 工作表
  
  ```python
  wb =  open_workbook(filePath)   # 打开工作表
  wb.close()        # 关闭工作表
  ```

- `sheet`操作
  
  ```python
  table = data.sheets()[0]          # 通过索引顺序获取
  
  table = data.sheet_by_index(sheet_indx)) # 通过索引顺序获取
  
  table = data.sheet_by_name(sheet_name)   # 通过名称获取
  
  names = data.sheet_names()    # 返回book中所有工作表的名字
  
  data.sheet_loaded(sheet_name or indx)   # 检查某个sheet是否导入完毕
  ```

- `sheet`的行操作
  
  ```python
  nrows = table.nrows  # 获取该sheet中的有效行数
  
  table.row(rowx)  # 返回由该行中所有的单元格对象组成的列表
  
  table.row_slice(rowx)  # 返回由该列中所有的单元格对象组成的列表
  
  table.row_types(rowx, start_colx=0, end_colx=None)    # 返回由该行中所有单元格的数据类型组成的列表
  
  table.row_values(rowx, start_colx=0, end_colx=None)   # 返回由该行中所有单元格的数据组成的列表
  
  table.row_len(rowx) # 返回该列的有效单元格长度
  ```

- `sheet`的列操作
  
  ```python
  ncols = table.ncols # 获取列表的有效列数
  
  table.col(colx, start_rowx=0, end_rowx=None) # 返回由该列中所有的单元格对象组成的列表
  
  table.col_slice(colx, start_rowx=0, end_rowx=None) # 返回由该列中所有的单元格对象组成的列表
  
  table.col_types(colx, start_rowx=0, end_rowx=None) # 返回由该列中所有单元格的数据类型组成的列表
  
  table.col_values(colx, start_rowx=0, end_rowx=None) # 返回由该列中所有单元格的数据组成的列表
  ```

- `sheet`单元格操作
  
  ```python
  table.cell(row,col)   # 返回对应的单元格对象
  table.cell_value(row,col)    # 返回对应单元格对象的值
  ```

- `sheet`获取合并表格
  
  > 这个地方还是有个坑的吧，如果你在打开表格的时候不把formatting_info属性设为True的话，是获取不到合并单元格的信息的
  
  ```python
  _refundTable.merged_cells    # 获取合并单元格的列表
  rs, re, cs, ce = merge        # 开始行结束行，开始列结束列信息
  ```

### 

## 3、openpyxl操作

>         `xlrd`和`openpyxl`的操作还是有一定的相似的。但**有一个比较明显的不同之处就是`xlrd`的行和列是由0开始的，而`openpyxl`的行和列是由1开始的。**另外，`openpyxl`也支持使用（如`A1`）这类写法。 
>     
>         还有就是，当遇到没有内容的单元格时，`xlrd`返回的是`“”`，而`openpyxl`返回的是`none`。如果`openpyxl`在打开时不选择`dataonly = True`的模式，`openpyxl`单元格如果为公式，则会返回公式而非值。
>     
>         这里使用`openpyxl`库作为`excel`的写入库，还是觉得他比`xlutils`和 `xlwriter`要方便挺多的。另外，在写入模式下，`openpyxl`的性能也没有他在读入模式下的那么逊色，一番操作后只需一个save`函数`即可。
> 
>  [参考博客](https://juejin.im/post/5cae014c6fb9a0686c0186df#heading-7)

##### 

- 打开工作表
  
  ```python
  wb = load_workbook(_reportPath)    # 打开工作表
  
  wb = workbook()    # 创建新的工作表
  table = wb.active()    # 激活默认的sheet
  
  ws1 = wb.create_sheet() #默认插在工作簿末尾
  
  ws2 = wb.create_sheet(0) # 插入在工作簿的第一个位置
  ws.title = "New Title"    # 修改sheet的title
  ```

- sheet操作
  
  ```python
  wb.get_sheet_names()    # 获取所有sheet的name
  
  wb["New Title"]        # 获取名字为那个的sheet
  wb.get_sheet_by_name("New Title")       # 和上面一样，但官方似乎不建议用这种方法了
  
  wb.remove(ws1)    #删除sheet
  ```

- sheet行操作
  
  ```python
  table.max_row    # 获取最大有效行数
  table.max_col    # 获取最大有效列数
  ```

- 单元格操作
  
  ```python
  cell = ws['A4'] # 获取第4行第A列的单元格
  
  ws['A4'] = 4 # 给第4行第A列的单元格赋值为4
  
  ws.cell(row=4, column=2, value=10) # 给第4行第2列的单元格赋值为10
  ws.cell(4, 2, 10) # 同上
  ```

##### 

## 4、打包exe

- #### 使用anaconda环境打包
  
  >         不太推荐在此环境下进行exe的打包。因为anaconda会将很多无关的库依赖链接进入程序，导致最后的exe文件高达几百兆的大小，极其吓人，还会在打包的过程中引发错误。但自己在这条路上遇到了一点问题，还是简单记录一下。
  
  1. ##### 使用pip安装pyinstaller
     
     ```python
     pip install pyinstaller
     ```
  
  2. ##### 安装完成后，使用pyinstaller命令
     
     ```python
     # @filePath 打包的python文件的路径
     # @iconPath 打包的icon的文件路径，可选
     pyinstaller -F filePath -i iconPath
     ```
     
            还有一堆的pyinstaller的命令，建议还是百度一下吧，就不列出来了。
  
  3. ##### 遇到一点报错了
     
     - ###### 错误一：installer maximum recursion depth exceeded
       
               超过最大的递归深度。这很可能是由于anaconda的库依赖太过复杂导致的。因为我在使用虚拟环境打包是就没这个错误。
           
               解决方法：打开生成的.spec文件，在文件开头添加
       
       ```python
       import sys
       sys.setrecursionlimit(1000000)
       ```
       
       .继续执行打包，但是改文件名：pyinstaller -F XXX.spec ,执行该文件。
     
     - ###### 错误二：**UnicodeDecodeError: 'utf-8' codec can't decode byte 0xce in position 110: invalid continuation byte**
       
                 修改D:\Python34\Lib\site-packages\PyInstaller\compat.py文件中[参考](https://stackoverflow.com/questions/47692960/error-when-using-pyinstaller-unicodedecodeerror-utf-8-codec-cant-decode-byt)
       
       ```python
       out = out.decode(encoding)
       为
       out = out.decode(encoding, errors='ignore')
       或
       out = out.decode(encoding, "replace")
       ```
     
     - ###### 错误三：pyinstaller module 'win32ctypes.pywin32.win32api' has no attribute 'error
       
       - 细看这个错误，似乎是由于icon文件的copy是造成的。
       
       - google一下，发现问题有点难解决，但还是找到了不明原因的解决方案
       
       - 将你的icon文件转换成ico格式的文件，并把文件放到打包目录的根目录下，问题就好了。
     
     - ###### 问题四：pyinstaller unpack requires a buffer of 16 bytes
       
       - 这个问题有点傻，是因为我直接将jpg的icon文件直接改拓展名的方式转换为ico格式文件，pyinstaller无法识别造成的。
       
       - google搜一下在线转ico就好了，一大堆
  
  ##### 

- #### 使用pipenv虚拟环境打包
  
  - ###### 安装pipenv
    
    `pip install pipenv`
  
  - ###### 选一个好目录做我们的虚拟环境，然后在该目录下:
    
    `pipenv install --python 3.7`
  
  - ###### 在命令行下激活环境
    
    `pipenv shell`
    
    输入这个命令，我们就进入到了新建的虚拟环境。如果你这时候使用命令 `pip list` 并发现里面只有很少的库，这就说明我们成功进入虚拟环境了
  
  - ###### 安装依赖的库
    
    ```python
    # 我就用到这几个了，所以就安装了这几个
    pipenv install pyinstaller
    pipenv install openpyxl
    pipenv install xlrd
    ```
  
  - ###### 把你的脚本放到这个目录下面，运行 pyinstaller，方法同前
  
  - ###### 你会发现exe小了太多了，我从三百多兆变成了六兆，你敢信
