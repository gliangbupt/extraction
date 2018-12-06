import os
import re
import xlrd
import xlsxwriter
import datetime


def name_split(pathname):  # 提取所有名称信息（包括光源和强度）
    symbol = "[ _-]"

    [path, f_name] = os.path.split(pathname)  # 将路径与名称分隔开保存为Tuple
    pre_result = re.split(symbol, f_name)  # 利用正则表达式分隔开名称
    result = '_'.join(pre_result[0:5])#这里要提取1到5的关键字,如‘VC2_F_0_50_(1)’

    # print(result)
    return result

def name_symbol(pathname):# 用来判别需要同时提取的文件名关键字
    symbol = "[ _-]"

    [path, f_name] = os.path.split(pathname)  # 将路径与名称分隔开保存为Tuple
    pre_result = re.split(symbol, f_name)  # 利用正则表达式分隔开名称
    result = '_'.join(pre_result[1:5])  # 这里要提取2到5的关键字，如‘F_0_50_(1)’

    return result

def name_ill(pathname):# 仅提取光源名
    symbol = "[ _-]"

    [path, f_name] = os.path.split(pathname)  # 将路径与名称分隔开保存为Tuple
    pre_result = re.split(symbol, f_name)  # 利用正则表达式分隔开名称
    result = '_'.join(pre_result[1:4])  # 这里要提取2到4的关键字，如‘F_0_50’

    return result

def stimul_read(path):#重点在于同时读取相匹配的两个文件
    parents = os.listdir(path)#文件名列表parents

    list_data=[]#初始化数据列表

    for i in range(len(parents)):
        pathname = path + parents[i]
        if ('.xls' or '.xlsx') in parents[i]:
            s=name_symbol(pathname)
            S='VC2_'+s #定义后面的正则匹配模式
            flash='FLASH_'+s
            name_str1=name_split(pathname)

            if flash == name_str1:#保证先读取的部分为flash文件
                workbook1=xlrd.open_workbook(pathname)
                sheet1=workbook1.sheet_by_name('Sheet1')
                for j in range(len(parents)):# 遍历寻找相同关键字的VC2文件
                    pathname2=path+parents[j]
                    name_str=name_split(pathname2)

                    # vc_file=re.search(S,name_str)
                    if S==name_str :

                        name=name_ill(pathname2)



                        workbook2=xlrd.open_workbook(pathname2)
                        sheet2=workbook2.sheet_by_name('Sheet1')
                        ev=sheet1.cell_value(34,7)
                        In=sheet1.cell_value(34,9)
                        ds=sheet1.cell_value(34,11)
                        oc=sheet1.cell_value(38,7)
                        grey=sheet1.cell_value(40,7)
                        sd=sheet2.cell_value(632,8)



                        value={'illuminant':name,'in EV':ev,'In %':In,
                               'in density':ds,'off-Centering':oc,
                               'Grey level at flash center':grey,
                               'Standard deviation':sd}
                        list_data.append(value)

    list_data = add_order(list_data)

    list_max = compare_extraction(list_data)
    list_mid = compare_extraction2(list_data)

    final_max = sorted(list_max, key=lambda x: x['order'])
    final_list = sorted(list_data, key=lambda x: x['order'])
    final_mid =sorted(list_mid,key=lambda x:x['order'])

    return final_list, final_max,final_mid

def compare_extraction(Lightlist):  # 生成指定光源含各个照度需求值的字典的列表（我说了个啥...）
    Lightlist_compare = []
    Lightlist_final = []
    i = 0
    j = 1
    # 分别设置i和j两个游标，j比i大1
    while i < len(Lightlist) and j < len(Lightlist):
        if j == len(Lightlist) - 1:#单独对最后几项处理，否则会丢失
            Lightlist_compare.append(Lightlist[i])#这里是加上了倒数第二项
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

def compare_extraction2(Lightlist):  # 因为flash要加上IN%的中间值提取
    Lightlist_compare = []
    Lightlist_final = []
    i = 0
    j = 1
    # 分别设置i和j两个游标，j比i大1
    while i < len(Lightlist) and j < len(Lightlist):
        if j == len(Lightlist) - 1:#单独对最后几项处理，否则会丢失
            Lightlist_compare.append(Lightlist[i])#这里是加上了倒数第二项
            Lightlist_compare.append(Lightlist[j])
            Lightlist_final.append(mid(Lightlist_compare))

        if Lightlist[i]['illuminant'] == Lightlist[j]['illuminant']:
            Lightlist_compare.append(Lightlist[i])
        else:
            Lightlist_compare.append(Lightlist[i])

            Lightlist_final.append(mid(Lightlist_compare))
            Lightlist_compare = []
        i += 1
        j += 1


    return Lightlist_final

def max(list_compare):# 每次都要修改key的参数，即排序规则
    sort_key = lambda x: x['Standard deviation']
    list_sort = sorted(list_compare, key=sort_key)
    list_max = list_sort[int(len(list_sort) - 1)]

    return list_max

def mid(list_compare):                                #返回中间值，list_compare存的是相同照度的含字典列表
    sort_key =lambda x:x['Grey level at flash center']            #排序因子
    list_sort=sorted(list_compare,key=sort_key)
    list_mid= list_sort[int((len(list_sort))/2)]

    return list_mid

def write_excel(list_data,list_ext,list_mid):
    today = datetime.datetime.now()
    time = today.strftime('%y%m%d%H%M')
    workbook = xlsxwriter.Workbook('flash_output_' + time + '.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.set_column('A:Z', 10)  # 更改列宽为10
    format1 = workbook.add_format({'num_format': '0.00'})
    format2 = workbook.add_format({'num_format': '0.00%'})
    # 写入所有数据
    row = 1
    col = 17
    for i in range(len(list_data)):
        item = list_data[i]  # item现在是一个字典
        worksheet.write(row, col, item['illuminant'])
        # worksheet.write(row + 1, col, str('in EV'))
        # worksheet.write(row + 1, col + 1, item['in EV'], format1)
        worksheet.write(row + 1, col, str('In %'))
        worksheet.write(row + 1, col + 1, item['In %'], format2)
        # worksheet.write(row + 3, col, str('in density'))
        # worksheet.write(row + 3, col + 1, item['in density'], format1)
        worksheet.write(row + 2, col, str('off-Centering'))
        worksheet.write(row + 2, col + 1, item['off-Centering'], format1)
        worksheet.write(row + 3, col, str('Grey level at flash center'))
        worksheet.write(row + 3, col + 1, item['Grey level at flash center'], format1)
        worksheet.write(row + 4, col, str('Standard deviation'))
        worksheet.write(row + 4, col + 1, item['Standard deviation'], format2)
        row += 5
    # 写入最大值数据
    row_2 = 2
    col_2 = 2
    worksheet.write(row_2 + 1, col_2 - 1, str('Standard deviation'))
    # worksheet.write(row_2 + 2, col_2 - 1, str('in EV'))
    worksheet.write(row_2 + 2, col_2 - 1, str('In %'))
    # worksheet.write(row_2 + 4, col_2 - 1, str('in density'))
    worksheet.write(row_2 + 3, col_2 - 1, str('off-Centering'))
    worksheet.write(row_2 + 4, col_2 - 1, str('Grey level at flash center'))
    for i in range(len(list_ext)):
        item = list_ext[i]
        worksheet.write(row_2, col_2, item['illuminant'])
        worksheet.write(row_2 + 1, col_2, item['Standard deviation'], format2)
        # worksheet.write(row_2 + 2, col_2, item['in EV'], format1)
        worksheet.write(row_2 + 2, col_2, item['In %'], format2)
        # worksheet.write(row_2 + 4, col_2, item['in density'], format1)
        worksheet.write(row_2 + 3, col_2, item['off-Centering'], format1)
        worksheet.write(row_2 + 4, col_2, item['Grey level at flash center'], format1)
        col_2+=1

    row_3 = 10
    col_3=2
    worksheet.write(row_3 + 4, col_3 - 1, str('Standard deviation'))
    # worksheet.write(row_2 + 2, col_2 - 1, str('in EV'))
    worksheet.write(row_3 + 2, col_3 - 1, str('In %'))
    # worksheet.write(row_2 + 4, col_2 - 1, str('in density'))
    worksheet.write(row_3 + 3, col_3 - 1, str('off-Centering'))
    worksheet.write(row_3 + 1, col_3 - 1, str('Grey level at flash center'))
    for i in range(len(list_mid)):
        item = list_mid[i]
        worksheet.write(row_3, col_3, item['illuminant'])
        worksheet.write(row_3 + 4, col_3, item['Standard deviation'], format2)
        # worksheet.write(row_2 + 2, col_2, item['in EV'], format1)
        worksheet.write(row_3 + 2, col_3, item['In %'], format2)
        # worksheet.write(row_2 + 4, col_2, item['in density'], format1)
        worksheet.write(row_3 + 3, col_3, item['off-Centering'], format1)
        worksheet.write(row_3 + 1, col_3, item['Grey level at flash center'], format1)
        col_3 += 1
    workbook.close()

def add_order(list_ext):
    for i in range(len(list_ext)):
        item=list_ext[i]
        s=item['illuminant'].split('_')
        illu=s[0]
        if illu == 'D65':
            item.update({'order':1})
        elif illu == 'TL84':
            item.update({'order':2})
        elif illu == 'TL83':
            item.update({'order':3})
        elif illu == 'A':
                item.update({'order': 4})
        elif illu == 'H':
            item.update({'order':5})
        elif illu == 'F':
            item.update({'order':6})
        elif illu == 'FA':
            item.update({'order':7})

    return list_ext


if __name__ == '__main__':
    print('选择路径')
    path = 'D:\My document\work\Execl提取数据\Data_extraction\\flash\\'
    a, b ,c= stimul_read(path)
    write_excel(a, b,c)