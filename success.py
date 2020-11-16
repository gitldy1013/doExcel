all_list = []
list_name = []
list_data = []

import xlrd
import xlutils.copy
import xlwt


def get_list(col_name):
    list_index = []
    list_no = {}
    for i in range(0, len(all_list)):
        list_index.append('')
    for j in col_name:
        if j in all_list:
            list_index[all_list.index(j)] = j
        else:
            list_no[col_name.index(j)] = j
    list_two = [list_index, list_no, all_list]
    return list_two


def set_style(name, height, bold=False):
    style = xlwt.XFStyle()  # 初始化样式
    font = xlwt.Font()  # 为样式创建字体
    font.name = name
    font.bold = bold
    font.color_index = 4
    font.height = height
    style.font = font
    pattern = xlwt.Pattern()
    pattern.pattern_fore_colour = 13  # 设置背景颜色为亮黄色
    return style


def write_excel(list_data, filename):
    # 创建工作簿
    # workbook = xlwt.Workbook(encoding='utf-8')
    data = xlrd.open_workbook(filename, formatting_info=True)  # 保留原格式
    workbook = xlutils.copy.copy(data)  # 复制之前表里存在的数据
    # 获取sheet
    data_sheet = workbook.get_sheet('Sheet4')
    # data_sheet.write(0, 0, '', set_style('Times New Roman', 220, True))
    j = 0
    for i in list_name:
        # data_sheet.write(0, j + 2, i, set_style('Times New Roman', 220, True))
        for x, element in enumerate(get_list(list_data[j])[0]):
            # data_sheet.write(x + 1, 1, get_list(list_data[j])[2][x], set_style('Times New Roman', 220, True))
            data_sheet.write(x + 1, j + 2, element, set_style('Times New Roman', 220, True))
            print(x, element)
            workbook.save(filename)
        j += 1

    # 确定栏位宽度
    col_width = []
    for i in range(len(list_name)):
        for j in range(len(list_name[i])):
            if i == 0:
                col_width.append(len_byte(list_name[i][j]))
            else:
                if col_width[j] < len_byte(str(list_name[i][j])):
                    col_width[j] = len_byte(list_name[i][j])

    # 设置栏位宽度，栏位宽度小于10时候采用默认宽度
    for i in range(len(col_width)):
        if col_width[i] > 10:
            data_sheet.col(i).width = 256 * (col_width[i] + 1)


# 获取字符串长度，一个中文的长度为2
def len_byte(value):
    length = len(value)
    utf8_length = len(value.encode('utf-8'))
    length = (utf8_length - length) / 2 + length
    return int(length)


def extract(filename):
    datacols = []
    data = xlrd.open_workbook(filename, encoding_override='utf-8')
    table = data.sheets()[3]  # 选定表
    nrows = table.nrows  # 获取行号
    ncols = table.ncols  # 获取列号
    x = 1
    for j in range(1, ncols):
        for i in range(1, nrows):  # 第0行为表头
            alldata = table.row_values(i)  # 循环输出excel表中每一行，即所有数据
            if j == 1:
                all_list.append(alldata[j])  # 取出表中第二列数据
            if j == x & j != 1:
                if alldata[j] != '':
                    datacols.append(alldata[j])
        if len(datacols)>0:
            list_data.append(datacols)
            datacols = []
        x += 1
    print(list_data)
    for j in range(2, ncols):
        namedata = table.row_values(0)  # 循环输出excel表中每一行，即所有数据
        if j >= 2:
            list_name.append(namedata[j])
    # print(list_name)


if __name__ == '__main__':
    filename = "20201112.xls"
    extract(filename)
    write_excel(list_data,filename)
    print('更新文件成功')
