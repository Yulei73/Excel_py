import xlrd  # 读取
import xlwt  # 生成
from os.path import join
from xlrd import open_workbook  # 桥梁
from xlutils.copy import copy


num8 = [15211206,
        16221222,
        16221071,
        16231234,
        16231230,
        16231206,
        16211411,
        16211405,
        16211403,
        16211389,
        16211409,
        16211402,
        16211399,
        16211368,
        16211366,
        16211378,
        16211377,
        16211379,
        16211373,
        16211380,
        16221013,
        16211375,
        16211369,
        16211393,
        16211372,
        16211412,
        16221241,
        16211406,
        16211388,
        16211401,
        16231261]

rb = open_workbook('finalfile.xlsx')

print(num8[1])
#wb = copy(rb)

#workbook = xlwt.Workbook()
workbook = copy(rb)
#sheet1 = workbook.add_sheet('sheet1', cell_overwrite_ok=True)
sheet1 = workbook.get_sheet('sheet1')

for i in range(31):
    # tworkbook = xlrd.open_workbook(u'pytest.xlsx')
    name = str(num8[i]) + '.xlsx'
    try:
        tworkbook = xlrd.open_workbook(str(name))
        tsheet = tworkbook.sheet_by_name('Sheet1')
        rows = 3  # 获取行数
        cols = 20  # 获取列数
        for j in range(cols):  # 读取每一列的数据
            # print(tsheet.cell(3,j))
            con1 = tsheet.cell(3, j)
            con2 = str(con1).strip('\'').split(':')
            print(con2)
            sheet1.write((i+3), j, con2[1].strip('.0').strip('\''))

        # sheet1.write(i, 0, namestr[0])
    except Exception as e:
        print(e)

workbook.save('finalfile.xls')
