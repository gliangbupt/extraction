import os
import re
import xlrd
import xlsxwriter
import datetime


def name_split(pathname):  # 提取名称信息（包括光源和强度）
    symbol = "[ _-]"

    [path, f_name] = os.path.split(pathname)  # 将路径与名称分隔开保存为Tuple
    pre_result = re.split(symbol, f_name)  # 利用正则表达式分隔开名称
    result = '_'.join(pre_result[:3])

    # print(result)
    return result


def read_excel(pathname):
    # 打开文件
    workbook = xlrd.open_workbook(pathname)
    sheet1 = workbook.sheet_by_name('Sheet1')
    if sheet1.cell_value(25,1)=='WARNING':# 在这儿加了一个判断，如果有warning则选取另一位置
        tac = sheet1.cell_value(84, 9)
        ea = sheet1.cell_value(84, 12)
        vn = sheet1.cell_value(90, 1)
    else:
        tac = sheet1.cell_value(60, 9)
        ea = sheet1.cell_value(60, 12)
        vn = sheet1.cell_value(66, 1)

    name_str = name_split(pathname)

    values = {'illuminant': name_str, 'Texture_acutance_clipped': tac,
              'Edge_acutance': ea, 'Visual_Noise': vn}

    return values


def batch_process(path):
    parents = os.listdir(path)
    list_data = []
    for i in range(0, len(parents)):
        pathname = path + parents[i]
        if ('.xls' or '.xlsx') in parents[i]:
            list_data.append(read_excel(pathname))

    list_max = compare_extraction(list_data)

    final_max = list_max
    final_list = list_data

    return final_list, final_max


def max(list_compare):
    sort_key = lambda x: x['Texture_acutance_clipped']
    list_sort = sorted(list_compare, key=sort_key)
    list_max = list_sort[int(len(list_sort) - 1)]

    return list_max


def compare_extraction(Lightlist):  # 生成指定光源含各个照度需求值的字典的列表（我说了个啥...）
    Lightlist_compare = []
    Lightlist_final = []
    i = 0
    j = 1
    # 分别设置i和j两个游标，j比i大1
    while i < len(Lightlist) and j < len(Lightlist):
        if j == len(Lightlist) - 1:  # 单独对最后几项处理，否则会丢失
            Lightlist_compare.append(Lightlist[i])  # 这里是加上了倒数第二项
            Lightlist_compare.append(Lightlist[j])
            Lightlist_final.append(max(Lightlist_compare))

        if Lightlist[i]['illuminant'] == Lightlist[j]['illuminant']:
            Lightlist_compare.append(Lightlist[i])
        else:
            Lightlist_compare.append(Lightlist[i])

            Lightlist_final.append(max(Lightlist_compare))
            Lightlist_compare = []
        i += 1
        j += 1

    return Lightlist_final

def write_excel(list_data,list_ext):
    today = datetime.datetime.now()
    time = today.strftime('%y%m%d%H%M')
    workbook = xlsxwriter.Workbook('Texture_output_' + time + '.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.set_column('A:Z',10)# 更改列宽为10
    format1 = workbook.add_format({'num_format': '0.000'})
    format2 = workbook.add_format({'num_format': '0.00'})

    # 写入所有数据
    row=1
    col=17
    for i in range(len(list_data)):
        item=list_data[i]#item现在是一个字典
        worksheet.write(row,col,item['illuminant'])
        worksheet.write(row+1,col,str('Texture_acutance_clipped'))
        worksheet.write(row+1,col+1,item['Texture_acutance_clipped'],format1)
        worksheet.write(row+2,col,str('Edge_acutance'))
        worksheet.write(row+2,col+1,item['Edge_acutance'],format1)
        worksheet.write(row +3 , col, str('Visual_Noise'))
        worksheet.write(row + 3, col+1,item['Visual_Noise'],format2)
        row+=4
    # 写入最大值数据
    row_2=2
    col_2=2
    worksheet.write(row_2+1, col_2-1, str('Texture_acutance_clipped'))
    worksheet.write(row_2+2, col_2-1, str('Edge_acutance'))
    worksheet.write(row_2 +3, col_2-1, str('Visual_Noise'))

    for i in range(len(list_ext)):
        item=list_ext[i]
        worksheet.write(row_2,col_2,item['illuminant'])
        worksheet.write(row_2 + 1,col_2,item['Texture_acutance_clipped'],format1)
        worksheet.write(row_2 + 2, col_2, item['Edge_acutance'],format1)
        worksheet.write(row_2 + 3, col_2 , item['Visual_Noise'],format2)
        col_2+=1
    workbook.close()


path = 'D:\My document\work\Execl提取数据\Data_extraction\\texture\\'
a,b=batch_process(path)
write_excel(a,b)

