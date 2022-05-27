import tushare as ts
import tkinter as tk
import tkinter.messagebox
from tkinter import *
import datetime
import sqlite3
import pandas as pd
import pandas_ta as ta
import re


# 获取数据函数，可传三个参数，num为股票代码，startday为起始日期 默认一年前，endday为结束日期 默认今天
def get_stock(num, startday=datetime.datetime.today() + datetime.timedelta(days=-365), endday=datetime.datetime.today()):
    stocknum = num
    today = datetime.datetime.today()

    startday_for_technical = startday + datetime.timedelta(days=-60)
    startday_for_technical = startday_for_technical.strftime('%Y%m%d')
    startday = startday.strftime('%Y%m%d')
    endday = endday.strftime('%Y%m%d')

    stock_df = pro.daily(ts_code=stocknum, start_date=startday, end_date=endday)
    stock_df_for_technical = pro.daily(ts_code=stocknum, start_date=startday_for_technical, end_date=endday)

    # 用于获取所有指标的
    # stock_df_for_technical['trade_date'] = pd.to_datetime(stock_df_for_technical['trade_date'])
    stock_df_for_technical = stock_df_for_technical.rename(columns={'trade_date': '日期'})
    stock_df_for_technical.set_index('日期', inplace=True)
    stock_df_for_technical = stock_df_for_technical.rename(columns={'vol': 'volume'})  # 列vol改名为volume
    stock_df_for_technical = stock_df_for_technical.iloc[::-1]  # 数据倒叙显示

    # 由于指标计算时有日期要求，因此最前部分的会有NaN，用stock_df_for_technical来解决
    # stock_df['trade_date'] = pd.to_datetime(stock_df['trade_date'])
    stock_df = stock_df.rename(columns={'trade_date': '日期'})
    stock_df.set_index('日期', inplace=True)
    stock_df = stock_df.rename(columns={'vol': 'volume'})  # 列vol改名为volume
    stock_df = stock_df.iloc[::-1]  # 数据倒叙显示

    return stock_df, stock_df_for_technical


def get_technical(stock_df):
    # MACD 返回结果的df有三列，MACD_12_26_9 MACDh_12_26_9 MACDs_12_26_9 上面会有NaN是由于数据不足26天无法计算，想要一年完整的数据只要把时间提前26天即可
    # MACD_12_26_9的数据表示DIF线、MACDh_12_26_9的数据表示柱方图（乘以2才是同花顺上的MACD柱方图）、MACDs_12_26_9的数据表示DEA线
    # MACD计算方法 MACD=2×（DIF-DEA）
    # 其中DIF由快的指数移动平均线（EMA12）减去慢的指数移动平均线（EMA26）得到  DEA为DIF的9日加权移动均线
    macd_df = ta.macd(close=stock_df['close'])
    macd_df['MACDh_12_26_9'] = macd_df['MACDh_12_26_9'] * 2

    # KDJ 根据默认的返回有三列，K_9_3      D_9_3      J_9_3
    # K_9_3的数据表示K线、D_9_3的数据表示D线、J_9_3的数据表示J线
    kdj_df = ta.kdj(stock_df['high'], stock_df['low'], stock_df['close'])

    # 均线 5、10、20天
    # 结果分别为SMA_5 SMA_10 SMA_20 表示5、10、20天均线
    # N日均线计算公式为  N日收盘价之和/N
    ma5_df = pd.DataFrame(ta.sma(stock_df['close'], length=5))
    ma10_df = pd.DataFrame(ta.sma(stock_df['close'], length=10))
    ma20_df = pd.DataFrame(ta.sma(stock_df['close'], length=20))

    # 连接所有技术指标结果与收盘价以列形式在一个DataFrame
    total_df = pd.concat([macd_df, kdj_df, ma5_df, ma10_df, ma20_df, stock_df['close']], axis=1)
    return total_df


def date_check():
    endday = datetime.datetime.today()
    startday = datetime.datetime.today() + datetime.timedelta(days=-365)

    # 日期有错误时赋值为0，让程序不继续往下执行
    if len(entry3.get()) == 0 and len(entry2.get()) == 0:
        pass
    elif len(entry3.get()) == 0 and len(entry2.get()) != 0:
        if len(entry2.get()) != 8:
            tk.messagebox.showerror(title='错误！', message='起始日期有误，请检查')
            startday = 0
        elif pd.to_datetime(entry2.get(), format='%Y%m%d', errors='coerce') > datetime.datetime.today():
            tk.messagebox.showerror(title='错误！', message='起始日期有误，请检查')
            startday = 0
        else:
            startday = pd.to_datetime(entry2.get(), format='%Y%m%d', errors='coerce')
    elif len(entry3.get()) != 0 and len(entry2.get()) == 0:
        if len(entry3.get()) != 8:
            tk.messagebox.showerror(title='错误！', message='结束日期有误，请检查')
            endday = 0
        elif pd.to_datetime(entry3.get(), format='%Y%m%d', errors='coerce') > datetime.datetime.today():
            tk.messagebox.showerror(title='错误！', message='结束日期有误，请检查')
            endday = 0
        else:
            endday = pd.to_datetime(entry3.get(), format='%Y%m%d', errors='coerce')
    else:
        if len(entry3.get()) != 8 or len(entry2.get()) != 8:
            tk.messagebox.showerror(title='错误！', message='日期有误，请检查')
            endday = 0
        elif pd.to_datetime(entry3.get(), format='%Y%m%d', errors='coerce') > datetime.datetime.today() or pd.to_datetime(entry2.get(), format='%Y%m%d', errors='coerce') > datetime.datetime.today():
            tk.messagebox.showerror(title='错误！', message='日期有误，请检查')
            endday = 0
        else:
            startday = pd.to_datetime(entry2.get(), format='%Y%m%d', errors='coerce')
            endday = pd.to_datetime(entry3.get(), format='%Y%m%d', errors='coerce')

    return startday, endday


def callback():
    writer = pd.ExcelWriter(r'数据结果.xlsx')
    workbook = writer.book

    num = re.split(r'[,，]', entry1.get())
    conn = sqlite3.connect('stock.db')
    cursor = conn.cursor()
    for i in num:
        sql_str = "select ts_code,symbol from stock_list_info where symbol='"+i+"'"
        cursor.execute(sql_str)
        result = cursor.fetchall()
        if len(result) == 0:
            tk.messagebox.showerror(title='错误！', message=i + '股票代码不存在')
            return 0
        else:
            stocknum = result[0][0]
            startday, endday = date_check()

            stock_df, stock_df_for_technical = get_stock(stocknum, startday, endday)
            total_df = get_technical(stock_df)
            total_df2 = get_technical(stock_df_for_technical)
            total_df2 = total_df2[len(total_df2)-len(total_df):]

            result_df = pd.concat([total_df2, stock_df['ts_code'], stock_df['open'], stock_df['high'], stock_df['low'], stock_df['pre_close'], stock_df['change'], stock_df['pct_chg'], stock_df['volume'], stock_df['amount']], axis=1)
            result_df = result_df.rename(columns={'MACD_12_26_9': 'DIF',
                                                  'MACDh_12_26_9': 'MACD',
                                                  'MACDs_12_26_9': 'DEA',
                                                  'K_9_3': 'K值',
                                                  'D_9_3': 'D值',
                                                  'J_9_3': 'J值',
                                                  'SMA_5': 'MA5',
                                                  'SMA_10': 'MA10',
                                                  'SMA_20': 'MA20',
                                                  'close': '收盘价',
                                                  'ts_code': '股票代码',
                                                  'open': '开盘价',
                                                  'high': '日内最高价',
                                                  'low': '日内最低价',
                                                  'pre_close': '昨收价',
                                                  'change': '涨跌额',
                                                  'pct_chg': '涨跌幅',
                                                  'volume': '成交量(手)',
                                                  'amount': '成交额(千元)'})  # 列改名

            result_df.to_excel(writer, sheet_name=i+'数据')
            writer_sheet = writer.sheets["%s数据" % i]
            writer_sheet.set_column('A:G', 15)
            writer_sheet.set_column('L:L', 10)
            writer_sheet.set_column('N:O', 12)
            writer_sheet.set_column('S:T', 15)
            print('保存完毕')
    writer.save()

    tk.messagebox.showinfo(title='提示', message='数据保存完毕！')


token = '9e932f78306d0a9f0fa9fdfe5a8b66ba13a5e23313de8ca996edda3f'
ts.set_token(token)
pro = ts.pro_api()

root_window = tk.Tk()
root_window.title('获取数据')
root_window.geometry('500x500')
root_window.resizable(0, 0)

# 创建一个frame窗体对象，用来包裹标签
frame = Frame(root_window, relief=SUNKEN, borderwidth=2, width=450, height=250)
# 在水平、垂直方向上填充窗体
frame.pack(side=TOP, fill=BOTH, expand=1)

label1 = tk.Label(frame, text="请输入股票代码(查询多个用逗号间隔)：")
label2 = tk.Label(frame, text="请输入数据起始日期(格式为YYYYMMDD)：")
label3 = tk.Label(frame, text="请输入数据结束日期(格式为YYYYMMDD)：")

entry1 = tk.Entry(root_window, width=50)
entry2 = tk.Entry(root_window, width=50)
entry3 = tk.Entry(root_window, width=50)


button1 = tk.Button(root_window, text="获取数据", width=10, command=callback)
button2 = tk.Button(root_window, text="关闭", width=10, command=root_window.quit)

label1.place(x=40, y=40)
label2.place(x=40, y=100)
label3.place(x=40, y=160)

entry1.place(x=45, y=70)
entry2.place(x=45, y=130)
entry3.place(x=45, y=190)

button1.place(x=40, y=310)
button2.place(x=180, y=310)

root_window.mainloop()
