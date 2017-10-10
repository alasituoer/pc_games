#coding:utf-8
import os
import pandas as pd
from openpyxl import Workbook

# 将指定路径下的多个csv文件, 提取指定字段后存为一个Excel文件
# 同时在第一列补上文件日期

working_space = '/Users/Alas/Documents/TD_handover/PC_Games/'

# 每期仅需修改下列三项
curdate = '2017-08-07'
dir_csv_files = 'icafe_2017-7-4_to_2017-8-6_CSV'
# 只有四个字段或是全部字段(选择一个, 注释其他)
list_columns = ['PkgId', 'PkgName', 'IdcClass', 'IdcHotExp']
#list_columns = ['PkgId', 'PkgName', 'IdcClass', 'LocalClass', 'PkgSize', 'MainExe', 'RunExe', 'AssocPkgId', 'PkgType', 'PkgVersion', 'IdcVersion', 'IdcUpdateDate', 'LocalUpdateDate', 'VDiskUpdateDate', 'LocalPath', 'LargeUrl', 'SmallUrl', 'PkgIdxCRC', 'LocalHotExp', 'LocalRecomExp', 'IdcHotExp', 'IdcRecomExp', 'AutoUpdate', 'PkgPriority', 'PkgOption', 'ServiceOption', 'CmdParam']


path_csv_files = working_space + 'data_source/' + curdate + '/' + dir_csv_files + '/'
path_to_write = working_space + 'results/' + curdate + '/'

# 得到指定文件夹下所有文件(包括文件夹)的名称(列表的形式)
list_csv_files = os.listdir(path_csv_files)
#print list_csv_files

# 去除文件列表中不需要的元素(去除的方法视情况而定)
# 如隐藏文件: '.DS_Store'
remove_from_file_list = [i for i,x in enumerate(list_csv_files) if 'DS_Store' in x]
# 从后往前删除不会影响未删除元素的位置数据
for i in remove_from_file_list[::-1]:
    #print i, file_list[i]
    del list_csv_files[i]
#print list_csv_files

def select_columns_to_excel(list_columns):
    wb = Workbook()
    # 设置默认活动表的标题和表头
    ws = wb.active
    ws.title = 'PC_Games'
    ws.append(['Date'] + list_columns)

    #if 'li' in 'limingzhi':
    for fs in list_csv_files:
	# 依次读取各CSV文件, 提取需要的指标
	#df_file_csv = pd.read_csv('/Users/Alas/Downloads/Game_Data_CSV/20151008230002.csv')    
	df_csv_file = pd.read_csv(path_csv_files + fs)    
#	print df_csv_file

	df_t1_csv_file = df_csv_file[list_columns]
	#print list(df_t1_csv_file.ix[0].values)

	# 将提取数据加上时间(第一列)写入Excel文件
	for i in range(len(df_t1_csv_file)):
	    #print [fs[:-4]] + list(df_t1_csv_file.ix[i].values)
	    ws.append([fs[:-4]] + list(df_t1_csv_file.ix[i].values))
        print fs
    to_write_filename = dir_csv_files + '.xlsx'
    if len(list_columns) > 5:
	to_write_filename = 'to_excel_to_mysql.xlsx'
	print 'Selected columns: ALL'
    wb.save(path_to_write + to_write_filename)


# 调用此函数
select_columns_to_excel(list_columns)


