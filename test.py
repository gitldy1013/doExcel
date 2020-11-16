all_list = ['申请人',
            '部门',
            '申请日期',
            '项目名称：',
            '项目编号：',
            '合同名称：',
            '合同编号：',
            '公司名称：',
            '计税方法：',
            '电话：',
            '纳税人识别号：',
            '备注：',
            '开户行',
            '银行账号：',
            '发票税率：',
            '税额(元)：',
            '不含税金额(元)：',
            '合同款项类别：',
            '合同类别：',
            '合同金额（元）：',
            '发票类型：',
            '【增加发票】',
            '发票金额(元)：',
            '发票号：',
            '发票附件：',
            '开票内容：',
            '发票日期：',
            '付款单编号：',
            '款项类别：',
            '付款方式：',
            '付款金额：',
            '跨区涉税事项联系人：',
            '合同期限有效期：',
            '项目所在地：',
            '开票金额（元）：',
            '经营项目类别：',
            '经办人：',
            '经营详细地址：',
            '跨区登记机关：',
            '公司地址',
            '发票种类：',
            '开票金额(元)：',
            '是否预缴税：',
            '预缴税金额(元)：',
            '上传预缴税凭证：',
            '增值税税额(元)：',
            '城市维护建设税(元)：',
            '教育费附加(元)：',
            '地方教育附加(元)：',
            '水利建设专项收入(元)：',
            '印花税(元)：',
            '企业所得税(元)：',
            '个人所得税(元)：',
            '其他（元）：',
            '是否已回款：',
            '设备清单：',
            '是否有发票：',
            '审批意见',
            '增值税税额：',
            '城市维护建设税：',
            '教育费附加：',
            '地方教育附加：',
            '跨区税源附件：',
            '挂账单编号：',
            '发票金额（元）：',
            '税额（元）：',
            '公司地址：',
            '银行回单/保函文件：',
            '付款日期：',
            '手续费-元：',
            '是否已挂账：']

waijingzheng = ['申请人',
                '部门',
                '申请日期',
                '合同款项类别：',
                '合同类别：',
                '项目名称：',
                '项目编号：',
                '合同名称：',
                '合同编号：',
                '合同金额（元）：',
                '公司名称：',
                '电话：',
                '跨区涉税事项联系人：',
                '纳税人识别号：',
                '合同期限有效期：',
                '项目所在地：',
                '备注：']

kuashuiqudengji = ['申请人',
                   '部门1',
                   '申请日期',
                   '合同款项类别：',
                   '合同类别：',
                   '项目名称：',
                   '项目编号：',
                   '合同名称：',
                   '合同编号：',
                   '合同金额（元）：',
                   '开票金额（元）：',
                   '计税方法：',
                   '公司名称：',
                   '纳税人识别号：',
                   '经营项目类别：',
                   '经办人：',
                   '电话：',
                   '经营详细地址：',
                   '跨区登记机关：']

kuashuiqushangchuan = ['申请人',
                       '部门',
                       '申请日期',
                       '合同款项类别：',
                       '合同类别：',
                       '项目名称：',
                       '项目编号：',
                       '合同名称：',
                       '合同编号：',
                       '合同金额（元）：',
                       '开票金额（元）：',
                       '计税方法：',
                       '公司名称：',
                       '纳税人识别号：',
                       '经营项目类别：',
                       '经办人：',
                       '电话：',
                       '经营详细地址：',
                       '跨区登记机关：',
                       '审批意见',
                       '增值税税额：',
                       '城市维护建设税：',
                       '教育费附加：',
                       '地方教育附加：',
                       '跨区税源附件：']

list_name = ['跨区涉税【外经证】上传', '跨区税源登记', '跨区税源上传']
list_data = [waijingzheng, kuashuiqudengji, kuashuiqushangchuan]

import xlwt


def get_list(col_name):
    list_index = []
    list_no = {}
    for i in range(0, 70):
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


def write_excel(list_data):
    # 创建工作簿
    workbook = xlwt.Workbook(encoding='utf-8')
    # 创建sheet
    data_sheet = workbook.add_sheet('demo', cell_overwrite_ok=True)
    data_sheet.write(0, 0, '', set_style('Times New Roman', 220, True))
    j = 0
    for i in list_name:
        data_sheet.write(0, j + 1, i, set_style('Times New Roman', 220, True))
        for x, element in enumerate(get_list(list_data[j])[0]):
            data_sheet.write(x + 1, 0, get_list(list_data[j])[2][x], set_style('Times New Roman', 220, True))
            data_sheet.write(x + 1, j + 1, element, set_style('Times New Roman', 220, True))
            print(x, element)
            workbook.save('demo.xls')
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


if __name__ == '__main__':
    write_excel(list_data)
    print('创建demo.xlsx文件成功')
