#coding:utf-8
import os
import pandas as pd
import MySQLdb

try:
    conn = MySQLdb.connect(
    host = ‘*’,
    port = *,
    user = ‘*’,
    passwd = ‘*’,
    db = ‘*’,
    charset = 'utf8',)
except Exception, e:
    print e
    sys.exit()
cursor = conn.cursor()

# 每期仅需要修改curdate
curdate = '2017-08-07'

working_space = ‘/PC_Games/'
list_excelfile = os.listdir(working_space + 'results/' + curdate + '/')
list_excelfile = [item for item in list_excelfile if 'to_mysql' in item]
#print list_excelfile

for excelfile in list_excelfile:
    df_onefile  = pd.read_excel(working_space + excelfile, sheetname=0)
    df_onefile.fillna('', inplace=True)
    #print df_onefile.head() 

    for i in df_onefile.index:
        try:
            sql = """INSERT INTO data VALUES ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}', '{14}', '{15}', '{16}', '{17}', '{18}', '{19}', '{20}', '{21}', '{22}', '{23}', '{24}', '{25}', '{26}', '{27}')"""
            Date, PkgId, PkgName, IdcClass, LocalClass, PkgSize, MainExe, RunExe, AssocPkgId, PkgType, PkgVersion, IdcVersion, IdcUpdateDate, LocalUpdateDate, VDiskUpdateDate, LocalPath, LargeUrl, SmallUrl, PkgIdxCRC, LocalHotExp, LocalRecomExp, IdcHotExp, IdcRecomExp, AutoUpdate, PkgPriority, PkgOption, ServiceOption, CmdParam = df_onefile.ix[i].values
            sql = sql.format(Date, PkgId, PkgName, IdcClass, LocalClass, PkgSize, MainExe, RunExe, AssocPkgId, PkgType, PkgVersion, IdcVersion, IdcUpdateDate, LocalUpdateDate, VDiskUpdateDate, LocalPath, LargeUrl, SmallUrl, PkgIdxCRC, LocalHotExp, LocalRecomExp, IdcHotExp, IdcRecomExp, AutoUpdate, PkgPriority, PkgOption, ServiceOption, CmdParam)
#            print sql
            cursor.execute(sql)
            conn.commit()
        except Exception, e:
            print e
            conn.rollback()
#        print len(df_onefile.ix[i].values)
        print df_onefile.ix[i].values[:3]

