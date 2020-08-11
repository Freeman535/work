import xlrd
import winreg
import time


def get_desktop():  # 获取桌面地址
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
    return winreg.QueryValueEx(key, 'Desktop')[0]


def read_x():  # 读取表格
    desktop = get_desktop()
    data = xlrd.open_workbook(str(desktop) + '\out.xls')
    table = data.sheets()[0]
    rows = table.nrows  # 行
    cols = table.ncols  # 列
    h = []
    data = '['
    for y in range(cols):
        h.append(str(table.row_values(0)[y]))
    print(h)
    for x in range(1, rows):
        data2 = ''
        for xy in range(cols):
            values = table.row_values(x)
            data2 = data2 + '"' + h[xy] + '": "' + str(values[xy]) + '", '
        data2 = '{' + data2[:-2] + '}, '
        data = data + data2
    data = data[:-2] + ']'
    return data


if __name__ == '__main__':
    print('把需要转换为json的xls文件放到桌面，文件名改为out.xls')
    time_c = time.strftime("%Y-%m-%d", time.localtime(time.time()))  # 获取本地日期
    author = input("今天是" + str(time_c) + "\n输入作者名")  # 获取作者
    desktop = get_desktop()
    d1 = read_x()
    d2 = str(d1).replace("\'", "\"")
    print(d2)
    d2 = "{\"date\":\"" + str(time_c) + "\",\"author\":\"" + str(author) + "\",\"main\":" + d2 + "}"
    jsFile = open(str(desktop) + '\code.json', "w+", encoding='utf-8')
    jsFile.write(d2)
    jsFile.close()
