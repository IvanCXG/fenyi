import xlrd
import datetime
import win32api
from time import *

# #修改对应的Excel点表
# IO_execl_name = 'CU07_20190112.xlsx'
# #保存到对应控制器的组态，分别为01_CU99.txt/02_CU99.txt……
# cu_99 = '07_CU99.txt'
cu_id = '11'

ao_name = []
do_name = []
ho_name = []

def CU_Generate():
    #写CU文件头
    now1=datetime.datetime.now()
    RevTime1=str(now1.strftime('%Y-%m-%d %H:%M:%S'))
    with open('./update/cu_need/%s_CU99.txt' %cu_id, 'w',encoding='UTF-8') as f:
        f.write('NuCon Cu File\n\n')
        f.write('FileHead\n')
        f.write('Version=3.0.0.0\n')
        f.write('Drop=99\n')
        f.write('Description=\n')
        f.write('Project=\n')
        f.write('Profile=10A319\n')
        f.write('Temperature=80\n')
        f.write('CpuLoad=80\n')
        f.write('MemLoad=80\n')
        f.write('MaxAxId=99\n')
        f.write('MaxDxId=99\n')
        f.write('MaxExchangeId=60\n')
        f.write('NetworkRedundancy=2\n')
        f.write('FileLastUpdate='+str(RevTime1)+'\n')

        t = str(int(time()))
        #print(t)
        now2=datetime.datetime.now()
        RevTime2=str(now2.strftime('%Y-%m-%d %H:%M:%S'))
        f.write('PointDirLastUpdate='+str(RevTime2)+' V1\n')
        f.write('FileHeadEnd\n\n')
        f.write('Class1OutputTimestamp='+ t +'\n')
        f.write('Class1OutputExchange,1,1,1,'+ t +',0,40000000\n\n\n')
    #读取Excel中的点名，并存储到list中
        data = xlrd.open_workbook('./update/IOlist_source/CU%s_201901.xlsx' %cu_id)
        sheet1_text = data.sheets()[0]
        nrows = sheet1_text.nrows
        print(nrows)
        for i in range(1,nrows):
            if sheet1_text.cell_value(i,6) == 'AO':
                new_name_ao = 'net-'+sheet1_text.cell_value(i,0)
                ao_name.append(new_name_ao)
            if sheet1_text.cell_value(i,6) == 'DOA':
                new_name_do = 'net-'+sheet1_text.cell_value(i,0)
                do_name.append(new_name_do)
            if sheet1_text.cell_value(i, 6) == 'HO':
                new_name_ho = 'net-' + sheet1_text.cell_value(i, 0)
                ho_name.append(new_name_ho)
        print(ao_name)
        print(do_name)
        do_len = len(do_name)
        print(do_len)
        ao_len = len(ao_name)
        print(ao_len)
        ho_len = len(ho_name)

        def p_calc(num):
            if num % 100 != 0:
                page_cst = int(num / 100) + 1
                return (page_cst)
            elif num % 100 == 0:
                page_cst = int(num / 100)
                # print(page_cst)
                return (page_cst)

        ao_page = p_calc(ao_len)
        do_page = p_calc(do_len)
        ho_page = p_calc((ho_len))
        print('AO页数 =', ao_page)
        print('DOA页数 =', do_page)
        print('HO页数 = ',ho_page)

        #AO
        for i in range(0,ao_page):
            a = 80
            b = 80
            f.write('Page, '+str(i+1)+':'+str((i+1)*2)+', 100 x10ms 6 0 0\n')
            f.write('	Description=AO-UP-DOWN\n')
            f.write('	RevTime=2019-01-15 21:25:29\n')
            f.write('	Sub=\n')
            for j in range(0,ao_len):
                temp1 = int(j / 16)
                temp2 = j % 16
                f.write('	Func, NetAO, '+str(j+1)+':'+str(10*(j+1))+', ('+str(a+150*temp1)+','+str(b+30*temp2)+'), 1, 0\n')
                f.write('		In= ,Null, \n')
                f.write('		Para= '+str(j+1)+','+ao_name[j]+',1,100,1,1,0,-99999.9,8,-99999.9,8,-99999.9,8,-99999.9,8,-99999.9,8,-99999.9,8,-99999.9,8,-99999.9,8,-99999.9,8,-99999.9,8,-99999.9,8,-99999.9,8,0,0,0,8,0,0,\n')
                f.write('		Out= ,\n')
                f.write('Page='+str(i+1)+',AO-UP-DOWN\n')
                f.write('	FuncEnd\n')
            f.write('	EndDesc=\n')
            f.write('	SubProfile=BE3C32B\n')
            f.write('PageEnd\n')

        # HO
        for i in range(0, ho_page):
            a = 80
            b = 80
            f.write('Page, ' + str(i + ao_page+1) + ':' + str((i + ao_page+1) * 2) + ', 100 x10ms 6 0 0\n')
            f.write('	Description=AO-UP-DOWN\n')
            f.write('	RevTime=2019-01-15 21:25:29\n')
            f.write('	Sub=\n')
            for j in range(0, ho_len):
                temp1 = int(j / 16)
                temp2 = j % 16
                f.write('	Func, NetAO, ' + str(j + 1) + ':' + str(10 * (j + 1)) + ', (' + str(a + 150 * temp1) + ',' + str(b + 30 * temp2) + '), 1, 0\n')
                f.write('		In= ,Null, \n')
                f.write('		Para= ' + str(j + 1) + ',' + ho_name[j] + ',1,100,1,1,0,-99999.9,8,-99999.9,8,-99999.9,8,-99999.9,8,-99999.9,8,-99999.9,8,-99999.9,8,-99999.9,8,-99999.9,8,-99999.9,8,-99999.9,8,-99999.9,8,0,0,0,8,0,0,\n')
                f.write('		Out= ,\n')
                f.write('Page=' + str(i +ao_page+ 1) + ',AO-UP-DOWN\n')
                f.write('	FuncEnd\n')
            f.write('	EndDesc=\n')
            f.write('	SubProfile=BE3C32B\n')
            f.write('PageEnd\n')


        #DO
        for k in range(0,do_page):
            f.write('Page, ' + str(k + ao_page + ho_page+1) + ':' + str((k + ao_page +ho_page+ 1)* 2) + ', 100 x10ms 6 0 0\n')
            f.write('	Description=DO-UP-DOWN\n')
            f.write('	RevTime=2019-01-16 18:15:16\n')
            f.write('	Sub=\n')
            for i in range(0,do_len):
                val1 = int(i/16)
                val2 = i%16
                a1 = 80
                b1 = 80
                f.write('	Func, NetDO, ' + str(i+ 1) + ':' + str(10 * (i + 1)) + ', (' + str(a1 + 150 * val1) + ',' + str(b1 + 30 * val2) + '), 1, 0\n')
                f.write('		In= ,Null, \n')
                f.write('		Para= '+str(i+1)+','+do_name[i]+',1,100,257,1,0,0,0,0,0,0,\n')
                f.write('		Out= ,\n')
                f.write('Page='+str(k+ao_page+ho_page+1)+',DO-UP-DOWN\n')
                f.write('	FuncEnd\n')
            f.write('	EndDesc=\n')
            f.write('	SubProfile=AB82F975\n')
            f.write('PageEnd\n')


        # AO数据源
        f.write('[POINT_DIR INFO]\n')
        f.write('BEGIN_AX\n')
        for i in range(0,ao_len):
            f.write(ao_name[i]+'=1000,,--------,--------,,7.2,100,0,1,'+str(80*i)+','+str(32+80*i)+','+str(i)+',1,'+str(1+3*i)+'\n')
        for i in range(0,ho_len):
            f.write(ho_name[i]+'=1000,,--------,--------,,7.2,100,0,1,'+str(80*i)+','+str(32+80*i)+','+str(i)+',1,'+str(1+3*i)+'\n')
        f.write('END_AX\n')

        #DO数据源
        f.write('BEGIN_DX\n')
        for i in range(0,do_len):
            f.write(do_name[i]+'=1000,,--------,--------,false,true,1,'+str(32*i)+','+str(32*i+1)+','+str(i+2)+',2,'+str(i+3)+'\n')
        f.write('END_DX\n')
CU_Generate()
# win32api.MessageBox(0,'！！！搞定！！！','提示')