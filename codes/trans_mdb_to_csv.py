#coding:utf-8
import os

# 将指定路径下的mdb文件转为csv文件

working_space = '/Users/Alas/Documents/TD_handover/PC_Games/'

# 每期仅需要修改下列两项
curdate = '2017-08-07'
dir_mdb_files = 'icafe_2017-7-4_to_2017-8-6/'

dir_csv_files = dir_mdb_files.replace('/', '_CSV')
path_mdb_files = working_space + 'data_source/' + curdate + '/' + dir_mdb_files + '/'
path_csv_files = working_space + 'data_source/' + curdate + '/' + dir_csv_files + '/'
#print path_csv_files

# 根据 "iCafeData_20170208-20170304"新建一个文件夹"iCafeData_20170208-20170304_CSV"
try:
    os.system('mkdir ' + path_csv_files)
    os.system('mkdir ' + working_space + 'results/' + curdate)
except Exception, e:
    print e


# 原始文件日文件夹
list_mdb_files = os.listdir(path_mdb_files)
#print list_mdb_files
remove_from_file_list = [i for i,x in enumerate(list_mdb_files) if 'DS_Store' in x]
# 从后往前删除不会影响未删除元素的位置数据
for i in remove_from_file_list[::-1]:
    del list_mdb_files[i]
#print list_mdb_files

for f in list_mdb_files:
    cmd = 'mdb-export ' + path_mdb_files + f + '/Data/Data.mdb Package > ' +\
            path_csv_files + f[:8] + '.csv'
    os.system(cmd)


