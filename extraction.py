import os
import re



def name_split(pathname,symbol):            #提取名称信息（包括光源和强度）
    # symbol=[_-]

    [path,f_name]=os.path.split(pathname)   #将路径与名称分隔开保存为Tuple
    pre_result=re.split(symbol,f_name)      #利用正则表达式分隔开名称
    result='_'.join(pre_result[:3])

    # print(result)
    return result



