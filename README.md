# pandasrw
一个高效易用的pandas  I/O库
pandas的I/O相对各类包一直偏慢且存在易用性问题，特别是对于大文件的读写，瓶颈非常明显。pandasio 库通过将各类库进一步封装，提高了pandas 读写excel、csv等文件的性能和易用性。性能和易用性提升详细说明见第二部分。

通过pip进行安装

```
pip install pandasrw
```

通过import导入包

```
import pandasrw
```

或者导入各函数简化使用，本库支持的函数见下文API部分。

```
from pandasrw import load,dump
```

一、常用API

本库的常用API为三个为load、dump和view，分别实现了文件的读取、写入和查看；其中查看功能为通过excel打开DataFrame或者文件，主要用于upyter等交互环境。此外还有流式加载表、csv转换为utf8编码、xlsx转换为csv等功能。

1、加载表
能够自动识别后缀和修改编码方式来实现加载。
```
df=load(file_path, col_name=None,sheetname='Sheet1',engine="polars")
```
示例：输入路径读取Sheet1表的全部列，生成pandas的DataFrame。 默认使用polars引擎，该表可以是xlsx、xlsx、csv和pkl格式。

```
df=load(file_path)
```

2、写入表
能够自动识别后缀写入，还实现了追加写和在后缀增加写入时间的功能。
```
dump(df,file_path,sheetname='Sheet1',engine="polars")
```
示例：输入路径，将pandas的DataFrame写入Sheet1表，默认使用polars引擎。该表可以是xlsx、xlsx、csv和pkl格式



能够自动识别后缀写入，还实现了追加写和在后缀增加写入时间的功能。
```
dump(df, file_path, mode=None, sheetname='Sheet1', time=False, engine="polars", cell='A1', visible=False,close=True)
```
示例：输入路径，将pandas的DataFrame写入Sheet1表，默认使用polars引擎，该表可以是xlsx、xlsx、csv和pkl格式。

```
dump(df,file_path)
```
2.1、追加写
当参数mode="a"时能够追加写。
```
dump(df, file_path, mode="a", sheetname='Sheet1',cell='A1')
```
支持csv和excel的追加写。pandas对csv的追加写支持较好，但是对excel的追加写比较繁琐。本库通过分别通过pandas和xlwings实现了追加写，其中小数据集通过pandas较为快速方便，对于大数据集或者需要指定写入文件单元格的情况使用xlwings库。 注意：
1、追加写时文档必须关闭，否则使用pandas引擎会报错，xlwings不会报错但是无法追加写入
2、需要指定写入文件单元格的情况引擎必须使用xlwings库，既engine == "xlwings"
```
dump(df, file_path, mode="a", sheetname='Sheet1',engine="xlwings", cell='A1')
```
2.3、后缀增加时间戳
当参数 time=False 时能够在文件后缀上自动添加写入时间的功能，时间格式为  -年月日_时分
```
dump(df, file_path, sheetname='Sheet1', time=False)
```
3、查看表
可以在excel中打开DataFrame和文件路径进行查看，方便在jupyter等交互环境中使用。输入参数f既可以是文件路径也可以是DataFrame。
```
view(f)
```

4、流式加载表

file_path是路径， row_count是没错读取的行
```
load_stream_row(file_path, row_count,col_name=None)
```
生成一个pandas.io.parsers.readers.TextFileReader对象

对于该迭代器对象，通过遍历迭代器分块运算

4.1、遍历迭代器

4.2、对于迭代器中的每个DataFrame进行运算

4.3、采用追加写（功能mode="a"或者mode="a+"）的方式写入csv。

注意：file_result_csv和上文函数中的file_path_csv绝对不能相同，即读取的csv和存入的csv不能同路径。否则会不停的迭代下去，不能退出循环。

原因：使用chunksize分块读取后，pandas并没有真正的将csv的内容加载入内存，只是解析了csv的内容和建立了连接（类似浅拷贝），在调用迭代器时再从csv中加载。所以再使用追加写的时候，会一边写入csv，一边再从csv中读取，形成死循环。

```
for df in df_iter:
    df=运算结果
    df.to_csv(file_result_csv,mode="a", index=False,encoding='UTF-8',header=False)
```

示例3：输入路径读取Sheet1表的全部列，生产pandas的DataFrame 默认使用polars引擎。该表可以是xlsx、xlsx、csv和pkl格式

```
df_iter=load_stream_row(file_path, row_count）
```
5、将csv转化为utf8编码

```
encode_to_utf8(filename, des_encode):
```
6、将xlsx转换为csv

```
 xlsxtocsv(file_path)
```

二、功能特点

1、性能提升
通过封装polars、pandas、xlwings库实现性能的提升和保证兼容性。通过参数engine来选择采用的读写引擎，默认采用polars库。对于polars和xlwings不支持的格式后端自动选择pandas进行兼容。各类引擎的特点如下。

polars库：
一个rust写的高性能库。读写快，支持xlsx和csv文件类型，其中xlsx为通过xlsx2csv库转换为csv来实现。兼容性较好，但是不如pandas。xlsx可以比pandas快2倍，csv可以快3倍.
建议：推荐使用。截至2023年3月该库的star已经超过1.3k，性能和质量已经经过检验。

pandas库：
支持xlsx、xls、csv、pkl四种文件类型，读写较慢，但是兼容性最好。
建议：备选。

xlwings库：
支持xlsx和xls两种文件类型，该库本身支持csv，但是在测试时性能较差未采用。
读写xlsx文件速度最快，可以达到pandas的3倍，但是在数据类型推断上较为粗糙，数据类型只能区分为object和float64两种类型，不能进一步区分int64等类型。本库通过调用 win32实现对excel的操作，需要使用windows操作协同且安装了excel软件。
建议：在excel数据较大一般超过100M时使用，且对细分数据类型不敏感或者采用手动修改数据类型的场景。


![在这里插入图片描述](https://img-blog.csdnimg.cn/cf1423ada429433e9a775beba2435e86.png#pic_center)


![在这里插入图片描述](https://img-blog.csdnimg.cn/b8347a47a68d419bbfaaa02f58126e9a.png#pic_center)

2、易用性提升
易用性方面主要做了以下三个方面的提升

2.1、对csv格式的编码自动修改为utf-8解决了 “‘utf-8’ codec can’t decode”等编码类报错。

2.2、对各类后缀进行了自适应，无需在手动指定excel、csv、pickle等

2.3、方便的实现csv、excel的追加写和在后缀加写入时间。

2.4、可以通过excel中随时查看DataFrame或文件，方便在jupyter等交互环境中使用。

3、大内存表的流式加载和计算
为了API的简洁暂时只实现了一个sheet的表格的流式加载和计算
<pandas.io.parsers.readers.TextFileReader at 0x2597d4080d0>

4、pickle支持
如果一个较大的表会多次使用，请转pickle存储，后续读写pickle文件会大大加快读写性能，本库支持pickle文件的读取和写入。
