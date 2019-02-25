import xlrd
import win32api

cu_id = '43'
# id_excel = '26'
# print(sys.getdefaultencoding())

def p_calc(num):
    if num % 100 != 0:
        page_cst = int(num / 100) + 1
        return (page_cst)
    elif num % 100 == 0:
        page_cst = int(num / 100)
        # print(page_cst)
        return (page_cst)

with open('./update/cu_source/CU0%s.txt' %cu_id,'r',encoding= 'gb18030',errors='ignore') as f:
    cu_lines = f.readlines()
    #print(cu_lines)
    lens = len(cu_lines)
    print(lens)
    # print(cu_lines)

    index = 0
    for i in range(0,lens):
        if cu_lines[i] != '[POINT_DIR_INFO]\n':
            index += 1
        elif cu_lines[i] == '[POINT_DIR_INFO]\n':
            #print('第%d行是：[POINT_DIR INFO]' % n)
            break
print(index)

# for item in cu_lines:
#     print(item)

data = xlrd.open_workbook('./update/IOlist_source/CU%s_201901.xlsx' %cu_id)
sheet1_text = data.sheets()[0]
nrows = sheet1_text.nrows
print(nrows)
ao_name = []
do_name = []
ho_name = []
for i in range(1, nrows):
    if sheet1_text.cell_value(i, 6) == 'AO':
        new_name_ao = 'net-' + sheet1_text.cell_value(i, 0)
        ao_name.append(new_name_ao)

    if sheet1_text.cell_value(i, 6) == 'DOA':
        new_name_do = 'net-' + sheet1_text.cell_value(i, 0)
        do_name.append(new_name_do)

    if sheet1_text.cell_value(i, 6) == 'HO':
        new_name_ho = 'net-' + sheet1_text.cell_value(i, 0)
        ho_name.append(new_name_ho)
# print(ao_name)
ao_lens = len(ao_name)
print('ao_counts=%d' % ao_lens)
do_lens = len(do_name)
print('do_counts=%d' % do_lens)
ho_lens = len(ho_name)

# print(do_name)
ao_page = p_calc(ao_lens)
do_page = p_calc(do_lens)
ho_page = p_calc(ho_lens)
print('AO页数 =', ao_page)
print('DOA页数=', do_page)
print('HO页数=',ho_page)


with open('./update/cu_need/CU%s-test.txt' %cu_id,'w',encoding='gb18030') as f:
    for i in range(0,index-1):
        f.write(cu_lines[i])
    #f.write('-----------我是分割线-------------\n')
    f.write('Class1OutputTimestamp=1547357956\n')
    f.write('Class1OutputExchange,1,1,1,1547357956,0,40000000\n\n')

    #AO
    if ao_page==0:
        pass
    else :
        for i in range(0,ao_page):
            f.write('Page, '+str(i+1)+':'+str(i+2)+', 100 x10ms 6 0 0\n')
            f.write('	Description=AI\n')
            f.write('	RevTime=2019-01-22 21:06:08\n')
            f.write('	Sub=\n')
            for j in range(0,ao_lens):
                temp1 = int(j/16)
                temp2 = j%16
                f.write('	Func, NetAI, ' + str(2*j + 1) + ':' + str(10 * (2*j + 1)) + ', (' + str(45 + 150 * temp1) + ',' + str(50 + 30 * temp2) + '), 1, 0\n')
                f.write('		In= ,\n')
                f.write('		Para= '+ao_name[j]+',65535,500,\n')
                f.write('		Out= ,0, \n')
                f.write('Page='+str(i+1)+',AI\n')
                f.write('	FuncEnd\n')
                f.write('	Func, PgAO, ' + str(2*j + 2) + ':' + str(10 * (2*j + 2)) + ', (' + str(150 + 150 * temp1) + ',' + str(42 + 30 * temp2) + '), 1, 0\n')
                f.write('		In= ,B'+str(2*j+1)+'-0, \n')
                f.write('		Para= \n')
                f.write('		Out= ,\n')
                f.write('Page=' + str(i + 1) + ',AI\n')
                f.write('	FuncEnd\n')
            f.write('	EndDesc=\n')
            f.write('	SubProfile=92992D8\n')
            f.write('PageEnd\n')


    #HO
    if ho_page==0:
        pass
    else :
        for i in range(0,ho_page):
            f.write('Page, '+str(i+ao_page+1)+':'+str(i+ao_page+2)+', 100 x10ms 6 0 0\n')
            f.write('	Description=HI\n')
            f.write('	RevTime=2019-01-22 21:06:08\n')
            f.write('	Sub=\n')
            for j in range(0,ho_lens):
                temp1 = int(j/16)
                temp2 = j%16
                f.write('	Func, NetAI, ' + str(2*j + 1) + ':' + str(10 * (2*j + 1)) + ', (' + str(45 + 150 * temp1) + ',' + str(50 + 30 * temp2) + '), 1, 0\n')
                f.write('		In= ,\n')
                f.write('		Para= '+ho_name[j]+',65535,500,\n')
                f.write('		Out= ,0, \n')
                f.write('Page='+str(i+1+ao_page)+',AI\n')
                f.write('	FuncEnd\n')
                f.write('	Func, PgAO, ' + str(2*j + 2) + ':' + str(10 * (2*j + 2)) + ', (' + str(150 + 150 * temp1) + ',' + str(42 + 30 * temp2) + '), 1, 0\n')
                f.write('		In= ,B'+str(2*j+1)+'-0, \n')
                f.write('		Para= \n')
                f.write('		Out= ,\n')
                f.write('Page=' + str(i + 1+ ao_page) + ',AI\n')
                f.write('	FuncEnd\n')
            f.write('	EndDesc=\n')
            f.write('	SubProfile=92992D8\n')
            f.write('PageEnd\n')


    #DO
    if do_page==0:
        pass
    else:
        for i in range(0,do_page):
            f.write('Page, '+str(i+ao_page+ho_page+1)+':'+str((i+ao_page+ho_page+1)*2)+', 100 x10ms 6 0 0\n')
            f.write('	Description=DI\n')
            f.write('	RevTime=2019-01-22 21:06:08\n')
            f.write('	Sub=\n')
            for j in range(0,do_lens):
                temp1 = int(j/16)
                temp2 = j%16
                f.write('	Func, NetDI, ' + str(2*j + 1) + ':' + str(10 * (2*j + 1)) + ', (' + str(45 + 150 * temp1) + ',' + str(50 + 30 * temp2) + '), 1, 0\n')
                f.write('		In= ,\n')
                f.write('		Para= '+do_name[j]+',65535,500,\n')
                f.write('		Out= ,0, \n')
                f.write('Page='+str(i+ao_page+do_page+1)+',DI\n')
                f.write('	FuncEnd\n')
                f.write('	Func, PgDO, ' + str(2*j + 2) + ':' + str(10 * (2*j + 2)) + ', (' + str(150 + 150 * temp1) + ',' + str(42 + 30 * temp2) + '), 1, 0\n')
                f.write('		In= ,B'+str(2*j+1)+'-0, \n')
                f.write('		Para= \n')
                f.write('		Out= ,\n')
                f.write('Page=' + str(i +ao_page+ho_page+1) + ',DI\n')
                f.write('	FuncEnd\n')

            f.write('	EndDesc=\n')
            f.write('	SubProfile=E5CE5DFD\n')
            f.write('PageEnd\n')

    for i in range(index,lens):
        f.write(cu_lines[i])
#win32api.MessageBox(0,'!!！搞定！！！','提示！')