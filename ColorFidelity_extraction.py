import os
import re
import xlrd
import xlsxwriter
import datetime


def name_split(pathname):                         # 提取名称信息（包括光源和强度）
    symbol = "[ _-]"

    [path, f_name] = os.path.split(pathname)      # 将路径与名称分隔开保存为Tuple
    pre_result = re.split(symbol, f_name)         # 利用正则表达式分隔开名称
    result = '_'.join(pre_result[:3])

    # print(result)
    return result


def read_excel(pathname):
    # 打开文件
    workbook = xlrd.open_workbook(pathname)
    # 获取所有sheet
    # print(workbook.sheet_names()) # [u'sheet1', u'sheet2']

    # 根据sheet索引或者名称获取sheet内容
    # sheet2 = workbook.sheet_by_index(0) # sheet索引从0开始
    sheet1 = workbook.sheet_by_name('Sheet1')

    # sheet的名称，行数，列数
    # print(sheet1.name,sheet1.nrows,sheet1.ncols)
    # 提取出需求的值
    contrast_a = sheet1.cell_value(210, 5)
    contrast_b = sheet1.cell_value(210, 25)
    contrast = (contrast_a - contrast_b) / (contrast_b + contrast_a)

    Color_fid = sheet1.cell_value(56, 17)
    White_blnc = sheet1.cell_value(57, 17)
    L22 = sheet1.cell_value(210, 17)

    name_str = name_split(pathname)
    namelist=name_str.split('_')
    light_level=float(namelist[2])
    # print(namelist)
    # 字典存储函数返回值
    values = {'illuminant':name_str,'light_level':light_level,
              'contrast': contrast, 'Color_fidelity': Color_fid,
              'L*_22nd': L22, 'White_balance': White_blnc}
    return values


def batch_process(path):
    parents = os.listdir(path)
    list_data= []
    for i in range(0, len(parents)):
        pathname = path + parents[i]
        if ('.xls' or '.xlsx') in parents[i]:
            list_data.append(read_excel(pathname))

    list_data=add_order(list_data)

    list_ext = compare_extraction(list_data)


    final_mid=sorted(list_ext,key=lambda x:x['order'])
    final_list=sorted(list_data,key=lambda x:x['order'])

    return final_list,final_mid

def write_excel(list_data,list_mid):


    today = datetime.datetime.now()
    time = today.strftime('%y%m%d%H%M')
    workbook = xlsxwriter.Workbook('ColorFidelity_output_' + time + '.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.set_column('A:Z',10)# 更改列宽为10
    format1=workbook.add_format({'num_format':'0.00'})

    # 写入所有数据
    row=1
    col=17
    for i in range(len(list_data)):
        item=list_data[i]#item现在是一个字典
        worksheet.write(row,col,item['illuminant'])
        worksheet.write(row+1,col,str('contrast'))
        worksheet.write(row+1,col+1,item['contrast'],format1)
        worksheet.write(row+2,col,str('Color_fidelity'))
        worksheet.write(row+2,col+1,item['Color_fidelity'],format1)
        worksheet.write(row +3 , col, str('L*_22nd'))
        worksheet.write(row + 3, col+1,item['L*_22nd'],format1 )
        worksheet.write(row + 4, col, str('White_balance'))
        worksheet.write(row + 4, col+1, item['White_balance'],format1)
        row += 5                                      #这儿的坑太多了...卧槽加5才对呀，每个光源占5行...
    #写入中间值数据
    row_2=2
    col_2=2
    worksheet.write(row_2+1, col_2-1, str('Color_fidelity'))
    worksheet.write(row_2+2, col_2-1, str('White_balance'))
    worksheet.write(row_2 +3, col_2-1, str('L*22nd'))
    worksheet.write(row_2 +4, col_2-1, str('contrast'))
    for i in range(len(list_mid)):
        item=list_mid[i]
        worksheet.write(row_2,col_2,item['illuminant'])
        worksheet.write(row_2 + 1,col_2,item['Color_fidelity'],format1)
        worksheet.write(row_2 + 2, col_2, item['White_balance'],format1)
        worksheet.write(row_2 + 3, col_2 , item['L*_22nd'],format1)
        worksheet.write(row_2 + 4, col_2, item['contrast'],format1)
        col_2+=1
    workbook.close()


def mid(list_compare):                                #返回中间值，list_compare存的是相同照度的含字典列表
    sort_key =lambda x:x['Color_fidelity']            #排序因子是色温（？）
    list_sort=sorted(list_compare,key=sort_key)
    list_mid= list_sort[int((len(list_sort))/2)]

    return list_mid


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

def add_order(list_ext):
    for i in range(len(list_ext)):
        item=list_ext[i]
        s=item['illuminant'].split('_')
        illu=s[1]
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
    path = 'D:\My document\work\Execl提取数据\Data_extraction\sample\\'
    a,b=batch_process(path)
    write_excel(a,b)

