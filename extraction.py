import os
import re
import xlrd
import xlwt
import numpy as np


def name_split(pathname):            #提取名称信息（包括光源和强度）
    symbol="[ _-]"

    [path,f_name]=os.path.split(pathname)   #将路径与名称分隔开保存为Tuple
    pre_result=re.split(symbol,f_name)      #利用正则表达式分隔开名称
    result='_'.join(pre_result[:3])

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
    #提取出需求的值
    contrast_a=sheet1.cell_value(210,5)
    contrast_b=sheet1.cell_value(210,25)
    contrast=(contrast_a - contrast_b)/(contrast_b + contrast_a)

    Color_fid=sheet1.cell_value(56,17)
    White_blnc=sheet1.cell_value(57,17)
    L22=sheet1.cell_value(210,17)
    #字典存储函数返回值
    values={'contrast':contrast,'Color_fidelity':Color_fid,'L*22nd':L22,'White_balance':White_blnc}
    return values

def batch_procsess(path):
    parents = os.listdir(path)
    D65=[]                                   #初始化代表各个光源的列表，用来存储字典
    TL84=[]
    TL83=[]
    A=[]
    H=[]
    F=[]

    for i in range(0, len(parents)):
        pathname=path+parents[i]
        # print(pathname)
        name_str=name_split(pathname)          #name_str存储分割出来的文件名，例如F_A_5
        if re.search('[_TL84_]',name_str):
            # dic=read_excel(pathname)           #按固定顺序匹配
            # TL84=TL84.append(dic)             #若是这样写则会修改obj本身并且返回None
            TL84.append(read_excel(pathname))
        if re.search('[_TL83_]',name_str):     #另外此处不可以用elif,因为elif隐含条件是
            TL83.append(read_excel(pathname))  #当第一个if为假时才会向下判定elif
        if re.search('[_D65_]',name_str):      #然而若if为真则下面不进行判断，直接if下面的
            D65.append(read_excel(pathname))   #语句块进行处理
        if re.search('[_A_]',name_str):
            A.append(read_excel(pathname))
        if re.search('[_H_]',name_str):
            H.append(read_excel(pathname))
        if re.search('[_F_]',name_str):
            F.append(read_excel(pathname))

    final_list=[D65,TL84,TL83,A,H,F]
    # print(final_list)
    return final_list








path='D:\My document\work\Execl提取数据\Data_extraction\sample\\'
batch_procsess(path)





