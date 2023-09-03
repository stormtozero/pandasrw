import os
import pandas as pd
import polars as pl
import xlwings as xw
import datetime
from xlsx2csv import Xlsx2csv
from chardet.universaldetector import UniversalDetector

# 将csv转化为utf8编码
def encode_to_utf8(filename, des_encode):
    # 读取文件的编码方式
    with open(filename, 'rb') as f:
        detector = UniversalDetector()
        for line in f.readlines():
            detector.feed(line)
            if detector.done:
                break
        original_encode = detector.result['encoding']
    # 读取文件的内容
    with open(filename, 'rb') as f:
        file_content = f.read()
    # 修改编码
    file_decode = file_content.decode(original_encode, 'ignore')
    file_encode = file_decode.encode(des_encode)
    with open(filename, 'wb') as f:
        f.write(file_encode)


# 通过xlwings读取表
def xw_open(file_path, sheetname='Sheet1', visible=False):
    # 数据类型可能推断不正确。测试时发现可以正确区分object和数据类型，但是数据类型都推断为float64不能区分int64
    app = xw.App(visible=visible,
                 add_book=False)

    book = app.books.open(file_path)

    sheet = book.sheets[sheetname]
    data = sheet.used_range.options(pd.DataFrame, header=1, index=False, expand='table').value
    data = data.convert_dtypes()
    if visible == False:
        book.close()
        app.quit()

    return data


# 通过xlwings写入表
def xw_write(df, file_path, sheetname='Sheet1', visible=False):
    # 数据类型可能推断不正确。测试时发现可以正确区分object和数据类型，但是数据类型都推断为float64不能区分int64
    app = xw.App(visible=visible,
                 add_book=False)
    wb = app.books.add()

    if sheetname == 'Sheet1':
        sheet = wb.sheets[0]
    else:
        sheet = wb.sheets.add(sheetname, after='Sheet1')

    # 将dataframe写入Sheet
    sheet.range('A1').value = df
    # 保存并关闭Workbook
    wb.save(file_path)
    if visible == False:
        wb.close()
        app.quit()


###通过xlwings追加写入
def xw_write_a(df, file_path, sheetname='Sheet1', cell='A1', visible=False, close=True):
    if not (os.path.exists(file_path)):
        wb = xw.Book()
        wb.save(file_path)

    app = xw.App(visible=visible, add_book=False)
    wb = app.books.open(file_path)

    sheet_names = [sht.name for sht in wb.sheets]
    if sheetname not in sheet_names:
        wb.sheets.add(sheetname, after=sheet_names[-1])

    sheet = wb.sheets[sheetname]
    sheet.range(cell).value = df
    wb.save()
    if close:
        wb.close()
        app.quit()

##通过xlwings查看df
def xw_view(df):
    # 启动Excel程序，不新建工作表薄（否则在创建工作薄时会报错），这时会弹出一个excel
    app= xw.App(visible = True, add_book= False)
    # 新建一个工作簿，默认sheet为Sheet1
    wb = app.books.add()
    #将工作表赋值给sht变量
    sht = wb.sheets('Sheet1')
    sht.range('A1').value =df


###通过pandas追加写入
def pd_write_a(df, file_path, sheetname='Sheet1'):
    with pd.ExcelWriter(file_path, mode="a", engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheetname)


# xlsx转换为csv
def xlsxtocsv(file_path):
    file_path_csv = file_path.replace(".xlsx", ".csv")
    Xlsx2csv(file_path, outputencoding="utf-8").convert(file_path_csv)
    return file_path_csv


# row_count每次读取的行数
def load_stream_row(file_path, row_count, col_name=None):
    name, ext = os.path.splitext(file_path)
    if '.csv' == ext:
        encode_to_utf8(file_path, des_encode="utf-8")
        df_read = pd.read_csv(file_path, usecols=col_name, chunksize=row_count)
    if ".xls" == ext:
        df_read = pd.read_excel(file_path, usecols=col_name)
        # 转化为csv再分块读
        file_path_csv = file_path.replace(".xls", ".csv")
        df_read.to_csv(file_path_csv, index=False, encoding='UTF-8')
        # encode_to_utf8(file_path_csv, des_encode="utf-8")
        df_read = pd.read_csv(file_path_csv, usecols=col_name, chunksize=row_count)
    if ".xlsx" == ext:
        file_path_csv = xlsxtocsv(file_path)
        df_read = pd.read_csv(file_path_csv, usecols=col_name, chunksize=row_count)
    return df_read


# 第一行是列名，从第二行开始读内容
def load_excel(file_path, sheetname='Sheet1', start_row=2, end_row=None):
    lst = []
    wb = load_workbook(filename=file_path, read_only=True)
    ws = wb[sheetname]
    max_row = ws.max_row
    # [*迭代器]方法性能高于list(迭代器)和逐行取值的方法
    row_columns = [*ws.iter_rows(min_row=1, max_row=1, values_only=True)]
    if end_row != None:
        row_data = [*ws.iter_rows(min_row=start_row, max_row=end_row, values_only=True)]
    else:
        row_data = [*ws.iter_rows(min_row=start_row, max_row=max_row, values_only=True)]
    # 将列名和数据合并
    row_columns.extend(row_data)
    df = pd.DataFrame(row_columns)
    return df


########主函数#############################################################################################################

# 自适用后缀、多引擎读取表，默认为polars
def load(file_path, col_name=None, sheetname='Sheet1', engine="polars", read_csv_options=None):
    name, ext = os.path.splitext(file_path)
    try:
        if '.csv' == ext:
            if engine == "polars":
                encode_to_utf8(file_path, des_encode="utf-8")
                df_read = pl.read_csv(file_path, columns=col_name)
                df_read = df_read.to_pandas()
            if engine == "pandas":
                encode_to_utf8(file_path, des_encode="utf-8")
                df_read = pd.read_csv(file_path, usecols=col_name)
            if engine == "xlwings":
                # xlwings读取csv兼容性和效率都较差调用pandas读取
                encode_to_utf8(file_path, des_encode="utf-8")
                df_read = pd.read_csv(file_path, usecols=col_name)

        if ".xlsx" == ext:
            if engine == "polars":
                df_read = pl.read_excel(file_path,
                                        read_csv_options=read_csv_options,
                                        sheet_name=sheetname)
                # 删除所有空行
                df_read = df_read.filter(~pl.all(pl.all().is_null()))
                df_read = df_read[[s.name for s in df_read if not (s.null_count() == df_read.height)]]
                df_read = df_read.to_pandas()
            if engine == "pandas":
                df_read = pd.read_excel(file_path, usecols=col_name, sheet_name=sheetname)
            if engine == "xlwings":
                df_read = xw_open(file_path, sheetname=sheetname, visible=False)

        if ".xls" == ext:
            if engine == "polars":
                # polars不能读xls格式，调用pandas解决
                df_read = pd.read_excel(file_path, usecols=col_name, sheet_name=sheetname)
            if engine == "pandas":
                df_read = pd.read_excel(file_path, usecols=col_name, sheet_name=sheetname)
            if engine == "xlwings":
                df_read = xw_open(file_path, sheetname=sheetname, visible=False)

        if ".pkl" == ext:
            df_read = pd.read_pickle(file_path)

    except Exception as e:
        if '.csv' == ext:
            encode_to_utf8(file_path, des_encode="utf-8")
            df_read = pd.read_csv(file_path, usecols=col_name)

        if ".xlsx" == ext:
            df_read = pd.read_excel(file_path, usecols=col_name, sheet_name=sheetname)

        print(f"读取文件发生错误：{e}")
        print(
            f'已自动切换兼容性更好的pandas引擎。下次读取该文件可以手动选择pandas引擎，语法为load(file_path,engine="pandas")，对于大文件尝试使用engine="xlwings"。')

    return df_read


# 自适应后缀、多引擎写入表,默认为polars，带有追加写入功能
def dump(df, file_path, mode=None, sheetname='Sheet1', time=False, engine="polars", cell='A1', visible=False,
         close=True):
    name, ext = os.path.splitext(file_path)
    if time:
        timestamp = datetime.datetime.now().strftime("%y%m%d_%H%M")
        """ 
        strftime="%y%m%d%H%M"  2位数年份 2306010101 Y大写则为4位数年份
        strftime="%m%d%H%M%S"   添加秒的信息 H M S时分秒 不能换为小写字母

        """
        base_path, ext_path = os.path.splitext(file_path)
        file_path_with_time = f"{base_path}-{timestamp}{ext}"
        file_path = file_path_with_time
    try:
        if mode == None:
            if '.csv' == ext:
                if engine == "polars":
                    df = pl.from_pandas(df)
                    df.write_csv(file_path, separator=",")
                if engine == "pandas":
                    df.to_csv(file_path, index=False)
                if engine == "xlwings":
                    # xlwings不能写入csv，调用pandas写入
                    df.to_csv(file_path, index=False)

            if ".xlsx" == ext:
                if engine == "polars":
                    df = pl.from_pandas(df)
                    df.write_excel(file_path, worksheet=sheetname)
                if engine == "pandas":
                    df.to_excel(file_path, index=False, sheet_name=sheetname)
                if engine == "xlwings":
                    xw_write(df, file_path, sheetname=sheetname, visible=False)

            if ".xls" == ext:
                if engine == "polars":
                    df.to_excel(file_path, index=False, sheet_name=sheetname)
                if engine == "pandas":
                    df.to_excel(file_path, index=False, sheet_name=sheetname)
                if engine == "xlwings":
                    xw_write(df, file_path, sheetname=sheetname, visible=False)

            if ".pkl" == ext:
                df.to_pickle(file_path)
        if mode == "a":
            if '.csv' == ext:
                df.to_csv(file_path, index=False, mode='a')

            if ".xlsx" == ext:
                if engine == "polars":
                    pd_write_a(df, file_path, sheetname=sheetname)
                if engine == "pandas":
                    pd_write_a(df, file_path, sheetname=sheetname)
                if engine == "xlwings":
                    xw_write_a(df, file_path, sheetname=sheetname, cell=cell, visible=False, close=True)

            if ".xls" == ext:
                if engine == "polars":
                    pd_write_a(df, file_path, sheetname=sheetname)
                if engine == "pandas":
                    pd_write_a(df, file_path, sheetname=sheetname)
                if engine == "xlwings":
                    xw_write_a(df, file_path, sheetname=sheetname, cell=cell, visible=False, close=True)

    except Exception as e:
        if '.csv' == ext:
            df.to_csv(file_path, index=False)

        if ".xlsx" == ext:
            df.to_excel(file_path, index=False, sheet_name=sheetname)

        print(f"写入文件发生错误：{e}")
        print(
            f'已自动切换兼容性更好的pandas引擎。下次写入该文件可以手动选择pandas引擎，语法为dump(file_path,engine="pandas")，对于大文件尝试使用engine="xlwings"。')

##通过excel查看数据，输入参数f既可以是文件路径也可以是DataFrame
def view(f):
    if type(f)==str:
        xw_open(f, sheetname='Sheet1', visible=True)
    else:
        xw_view(f)
