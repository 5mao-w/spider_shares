import urllib.request, urllib.error
import re
import xlwt
import xlrd
import xlutils.copy
from xlutils.copy import copy

book_name_xls = 'shares.xls'  # excl文件名
sheet_name_xls = 'shares'  # 工作簿名
value_title = [['股票代码', '上升日期', '上升日成交量较前日增值百分比']]  # excl列表内容名称

# 个人中心网址（未登录版）
cent_url = "http://myfavor1.eastmoney.com/v4/anonymwebouter/gstkinfos?appkey=d41d8cd98f00b204e9800998ecf8427e&bid=cc7a025f9247031f208ecfc2adbc70f0&cb=jQuery331008911013707299698_1645183581107&g=1&_=1645183581158"
# 网址请求头
head = {
    'Referer': 'http://quote.eastmoney.com/',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.80 Safari/537.36 Edg/98.0.1108.50'
}


# 主函数：所有方法的入口
def main():
    html = ask_url(cent_url, head)  ## 得到指定一个URL的网页内容
    code_list = get_code(html)
    day_data = get_data(code_list)
    up_data = get_up_data(day_data)
    try:
        xlrd.open_workbook(book_name_xls)
    except FileNotFoundError:
        write_excel_xls(book_name_xls, sheet_name_xls, up_data, value_title)
    else:
        write_excel_xls_append(book_name_xls, up_data)
    # get_down_data(day_data)
    input('结果已成功获得请退出')


# 得到指定一个URL的网页内容
def ask_url(url, head):
    request = urllib.request.Request(url, headers=head)
    html = ''
    try:
        response = urllib.request.urlopen(request)
        html = response.read().decode('utf-8')
        # print(html)
    except urllib.error.URLError as e:
        if hasattr(e, 'code'):
            print(e.code)
        if hasattr(e, 'reason'):
            print(e.reason)
    return html


# 获得自选股编码
def get_code(html):
    code_rule = re.compile('"security":"((\d)\$(\d+)\$)\d+"')
    code = re.findall(code_rule, html)
    return code


# 获得所有自选股的每日数据
def get_data(code_list):
    day_data = {}
    for code in code_list:
        # 每支股票日数据网址
        ##
        home_url = "http://push2his.eastmoney.com/api/qt/stock/kline/get?cb=jQuery112407568100035914511_1644991027184&fields1=f1%2Cf2%2Cf3%2Cf4%2Cf5%2Cf6&fields2=f51%2Cf52%2Cf53%2Cf54%2Cf55%2Cf56%2Cf57%2Cf58%2Cf59%2Cf60%2Cf61&ut=7eea3edcaed734bea9cbfc24409ed989&klt=101&fqt=1&secid="
        center_url = code[1] + '.' + code[2]
        end_url = "&beg=0&end=20500000&_=1644991027361"
        ##
        code_url = home_url + center_url + end_url
        html = ask_url(code_url, head)
        data_rule = re.compile(
            "(\d+-\d+-\d+),\d+\.\d+,\d+\.\d+,(\d+\.\d+),(\d+\.\d+),(\d+),\d+\.\d+,\d+\.\d+,-?\d+\.\d+,-?\d+\.\d+,\d+\.\d+")
        html_data = re.findall(data_rule, html)
        day_data[code[2]] = html_data
    # print(day_data)
    return day_data


# 判断出满足上升条件的自选股数据
def get_up_data(dayData):
    up_data = []
    for k in dayData:
        # 判断条件为i-1天最低价大于前后两天，且（i天成交量 - i-1天成交量） / i-1天成交量 >= 5 / 100
        if dayData[k][-2][2] < dayData[k][-1][2] and dayData[k][-2][2] < dayData[k][-3][2] and abs(
                int(dayData[k][-2][-1]) - int(dayData[k][-1][-1])) / int(dayData[k][-2][-1]) >= 5 / 100:
            day_list = [k, dayData[k][-1][0], '{:.2f}%'.format(
                (int(dayData[k][-1][-1]) - int(dayData[k][-2][-1])) / int(dayData[k][-2][-1]) * 100)]
            up_data.append(day_list)
    if len(up_data) == 0:
        print('今天没有处于上升拐点的股票')
    else:
        print(up_data)
        return up_data


# 判断出满足下降条件的自选股数据
def get_down_data(dayData):
    up_data = []
    for k in dayData:
        # 判断条件为i-1天最低价小于前后两天，且（i天成交量 - i-1天成交量） / i-1天成交量 >= 5 / 100
        if dayData[k][-2][2] > dayData[k][-1][2] and dayData[k][-2][2] > dayData[k][-3][2] and abs(
                int(dayData[k][-2][-1]) - int(dayData[k][-1][-1])) / int(dayData[k][-2][-1]) >= 5 / 100:
            day_list = [k, dayData[k][-1][0], '{:.2f}%'.format(
                (int(dayData[k][-1][-1]) - int(dayData[k][-2][-1])) / int(dayData[k][-2][-1]) * 100)]
            up_data.append(day_list)
    if len(up_data) == 0:
        print('今天没有处于下降拐点的股票')
    else:
        print(up_data)
        return up_data


def write_excel_xls(path, sheet_name, value, title):  # 保存每日满足条件的自选股数据   ##初次启动新建一个xls文件
    workBook = xlwt.Workbook(encoding='utf-8')  # 新建一个工作簿
    workSheet = workBook.add_sheet(sheet_name)  # 在工作簿中新建一个表格
    # workSheet.col(0).width = 256*10
    # workSheet.col(1).width = 256*15
    # workSheet.col(2).width = 256*30
    # style = xlwt.XFStyle()                      # 创建一个样式对象，初始化
    # al = xlwt.Alignment()
    # al.horz = 0x02    #设置水平居中
    # al.vert = 0x01    #设置垂直居中
    # style.alignment = al
    for i in range(len(value)):
        for k in range(len(value[i])):
            if i == 0:
                workSheet.write(i, k, title[0][k])  # 在表格中标题
                workSheet.write(i + 1, k, value[i][k])  # 在表格中写入数据（对应的列和行）
            else:
                workSheet.write(i + 1, k, value[i][k])  # 在表格中写入数据（对应的列和行）
    workBook.save(path)  # 保存工作簿


##后续追加数据
def write_excel_xls_append(path, value):
    workbook = xlrd.open_workbook(path, formatting_info=True)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的第一个表格
    row_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    for i in range(len(value)):
        for j in range(len(value[i])):
            new_worksheet.write(i + row_old, j, value[i][j])  # 追加写入数据
    new_workbook.save(path)  # 保存工作簿


if __name__ == '__main__':
    main()
