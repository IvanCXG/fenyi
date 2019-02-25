import xlrd
import os

cu_id='18'
data = xlrd.open_workbook('./update/IOlist_source/CU%s_201901.xlsx' %cu_id)
sheet1_text = data.sheets()[0]
nrows = sheet1_text.nrows
print(nrows)
ao_name = []
do_name = []
ho_name = []
for i in range(1, nrows):
    if sheet1_text.cell_value(i, 6) == 'AO':
        new_name_ao = sheet1_text.cell_value(i, 0)
        ao_name.append(new_name_ao)

    if sheet1_text.cell_value(i, 6) == 'DOA':
        new_name_do = sheet1_text.cell_value(i, 0)
        do_name.append(new_name_do)
# print(ao_name)
# print(do_name)

index_ao_list = []
index_do_list = []
p_index_do_list = []
str1 = 'Func, AO'
str2 = 'Func, NetDO'
str3 = 'LanModule,uDOA'
k = 0
s = 0
with open('./update/cu_use/CU%s-test.txt' %cu_id,'r')as f:
    cu_text = f.readlines()
    print(cu_text)
    lens = len(cu_text)
    print('txt总行数：',lens)

    for i in range(0,lens):
        m = cu_text[i].find(str1)
        n = cu_text[i].find(str2)
        # l = cu_text[i].find(str3)
        if m != -1:
            index_ao = i+2
            index_ao_list.append(index_ao)
            #print(index_ao)

        if n != -1:
            for item in do_name:

                if cu_text[i+2].find(item) != -1:
                    #print(cu_text[i + 2])
                    index_do = i+2
                    # print(index_do)
                    p_index_do_list.append(index_do)
index_do_list = list(set(p_index_do_list))
index_do_list.sort()
# print(index_do_list)
print('do个数：',len(index_do_list))

# print(cu_text)
# print(index_ao_list)
print('ao个数：',len(index_ao_list))
index_ao_list.sort()
with open('./update/cu_test/CU%s-test-%s.txt' %(cu_id,cu_id),'w')as f:
    if len(index_ao_list) ==0:
        for i in range(0,len(cu_text)):
            t = i +1
            if t in index_do_list:
                s = s + 1
                str4 = '		In= ,B' + str(s * 2) + '-1,\n'
                cu_text[i] = cu_text[i].replace(cu_text[i], str4)
                f.write(cu_text[i])
            else:
                f.write(cu_text[i])

    if len(index_do_list) ==0:
        for i in range(0, len(cu_text)):
            t = i + 1
            if t in index_ao_list:
                k = k + 1
                str3 = '		In= ,B' + str(k * 1) + '-1,\n'
                cu_text[i] = cu_text[i].replace(cu_text[i], str3)
                f.write(cu_text[i])
            else:
                f.write(cu_text[i])
    if len(index_ao_list) !=0 and len(index_do_list) != 0:
        for i in range(0,len(cu_text)):
            t = i+1
            if t in index_ao_list:
                k = k+1
                str3 ='		In= ,B'+str(k*2)+'-1,\n'
                cu_text[i]=cu_text[i].replace(cu_text[i],str3)
                f.write(cu_text[i])
                # print(cu_text[i])
                # print(str3)
            elif t in index_do_list:
                s = s + 1
                str4 = '		In= ,B' + str(s * 2) + '-2,\n'
                cu_text[i] = cu_text[i].replace(cu_text[i], str4)
                f.write(cu_text[i])
            else:
                f.write(cu_text[i])


