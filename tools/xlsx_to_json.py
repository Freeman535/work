import xlrd
import winreg
import time


def get_desktop():  # 获取桌面地址
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
    return winreg.QueryValueEx(key, 'Desktop')[0]


def read_x():  # 读取表格
    desktop = get_desktop()
    data = xlrd.open_workbook(str(desktop) + '\data_json_format.xlsx')
    table = data.sheets()[0]
    rows = table.nrows
    data = []
    for x in range(1, rows):
        values = table.row_values(x)
        data.append(
            (
                {
                    "sp_code": str(values[0]),
                    "code": str(values[1]),
                    "wl_code": str(values[2]),
                    "sp_gg": str(values[3]),
                    "gg": str(values[4]),
                    "name": str(values[5]),
                    "jhj": str(values[6]),
                    "sj": str(values[7]),
                    "class1": str(values[8]),
                    "class2": str(values[9]),
                    "class3": str(values[10]),
                }
            )
        )
    return data


if __name__ == '__main__':
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
