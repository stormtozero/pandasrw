import os
import pandas as pd
import polars as pl
import xlwings as xw
from chardet.universaldetector import UniversalDetector

#将csv转化为utf8编码
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
    #修改编码
    file_decode = file_content.decode(original_encode, 'ignore')
    file_encode = file_decode.encode(des_encode)
    with open(filename, 'wb') as f:
        f.write(file_encode)


#通过xlwings读取表
def xw_open(file_path, sheetname='Sheet1', visible=False):
    #数据类型可能推断不正确。测试时发现可以正确区分object和数据类型，但是数据类型都推断为float64不能区分int64
    app = xw.App(visible=visible,
                 add_book=False)

    book = app.books.open(file_path)

    sheet = book.sheets[sheetname]
    data = sheet.used_range.options(pd.DataFrame, header=1, index=False, expand='table').value

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

#xlsx转换为csv
def xlsxtocsv(file_path):
    file_path_csv=file_path.replace(".xlsx", ".csv")
    Xlsx2csv(file_path, outputencoding="utf-8").convert(file_path_csv)
    return file_path_csv


#自适用后缀、多引擎读取表，默认为polars
def load(file_path, col_name=None,sheetname='Sheet1',engine="polars"):

    name, ext = os.path.splitext(file_path)

    if '.csv' == ext:
        if engine=="polars":
            encode_to_utf8(file_path, des_encode="utf-8")
            df_read = pl.read_csv(file_path, columns=col_name)
            df_read = df_read.to_pandas()
        if engine=="pandas":
            encode_to_utf8(file_path, des_encode="utf-8")
            df_read = pd.read_csv(file_path, usecols=col_name)
        if engine == "xlwings":
            #xlwings读取csv兼容性和效率都较差调用pandas读取
            encode_to_utf8(file_path, des_encode="utf-8")
            df_read = pd.read_csv(file_path, usecols=col_name)

    if ".xlsx" == ext:
        if engine == "polars":
            df_read = pl.read_excel(file_path, read_csv_options={"columns": col_name}, sheet_name=sheetname)
            df_read = df_read.to_pandas()
        if engine == "pandas":
            df_read = pd.read_excel(file_path, usecols=col_name, sheet_name=sheetname)
        if engine == "xlwings":
            df_read = xw_open(file_path, sheetname=sheetname, visible=False)

    if ".xls" == ext:
        if engine == "polars":
            #polars不能读xls格式，调用pandas解决
            df_read = pd.read_excel(file_path, usecols=col_name, sheet_name=sheetname)
        if engine == "pandas":
            df_read = pd.read_excel(file_path, usecols=col_name, sheet_name=sheetname)
        if engine == "xlwings":
            df_read = xw_open(file_path, sheetname=sheetname, visible=False)

    if ".pkl" == ext:
        df_read=pd.read_pickle(file_path)

    return df_read


#自适应后缀、多引擎写入表
def dump(df,file_path,sheetname='Sheet1',engine="polars"):
    name, ext = os.path.splitext(file_path)
    if '.csv' ==ext:
        if engine == "polars":
            df=pl.from_pandas(df)
            df.write_csv(file_path, separator=",")
        if engine == "pandas":
            df.to_csv(file_path,index=False)
        if engine == "xlwings":
            #xlwings不能写入csv，调用pandas写入
            df.to_csv(file_path,index=False)

    if ".xlsx" == ext:
        if engine == "polars":
            df=pl.from_pandas(df)
            df.write_excel(file_path, worksheet=sheetname)
        if engine == "pandas":
            df.to_excel(file_path, index=False,sheet_name=sheetname)
        if engine == "xlwings":
            xw_write(df, file_path, sheetname=sheetname, visible=False)

    if ".xls" == ext:
        if engine == "polars":
            df.to_excel(file_path, index=False,sheet_name=sheetname)
        if engine == "pandas":
            df.to_excel(file_path, index=False,sheet_name=sheetname)
        if engine == "xlwings":
            xw_write(df, file_path, sheetname=sheetname, visible=False)

    if ".pkl" == ext:
        df.to_pickle(file_path)

#row_count每次读取的行数
def load_stream_row(file_path, row_count,col_name=None):

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
        file_path_csv = xlsxtocsv(path_xlsx)
        df_read = pd.read_csv(file_path_csv, usecols=col_name, chunksize=row_count)
    return df_read