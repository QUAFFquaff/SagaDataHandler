import csv
import re
import os
import pandas as pd
import xlrd
import numpy as np
import xlsxwriter
import xlwt
from xlwt import Workbook


def read_excel(file):
    data = xlrd.open_workbook(file)
    table = data.sheets()[0]

    list_values = []
    end = len(table.col_values(0))  # from the first line to last line
    for x in range(end):
        values = []
        row = table.row_values(x)
        for i in [0, 1, 2]:
            values.append(row[i])
        list_values.append(values)
    return list_values

CCS = ['Time', 'CSHPSTS1_ChilledWaterSupplyTemp', 'CSHPRTS1_ChilledWaterReturnTemp','CSHPSFS1_ChilledWaterFlow',
                        'CSHPSTS2_CoolingWaterSupplyTemp','CSHPRTS2_CoolingWaterReturnTemp','CSHPSFS2_CoolingWaterFlow',
                        'CCS_LastHourAccCool','CCS_LastHourAccE','CCS_AccCool']
CTW = ['Time','deviceUID','CTW_PunctualActiveP','CTW_FanFreq','CTW_LastHourAccE']  # CTW
CWP = ['Time','deviceUID','CWP_PunctualActiveP','CWP_Freq','CWP_LastHourAccE']  # CWP
CHWP = ['Time','deviceUID','CHWP_PunctualActiveP','CHWP_Freq','CHWP_LastHourAccE']  # CHWP
our_dic = ['Time','deviceUID','CHU_ChilledWaterOutTemp','CHU_ChilledWaterInTemp',
           'CHU_CoolingWaterOutTemp','CHU_CoolingWaterInTemp',
           'CHU_PunctualActiveP','CHU_LastHourAccE','CHU_ChillerLoadRate']  # CHU

sheet_list = []
def read_txt(filename,data_table=[]):
    with open(filename, 'r') as file_to_read:
        while True:
            data_dic = {}
            lines = file_to_read.readline()  # 整行读取数据
            # print(lines)
            if not lines:
                break
                pass
            E_tmp = re.findall(r'["](.*?)["]', lines)
            # print(E_tmp)
            data_dic['Time']=E_tmp[-7]
            # for i in range(1,9):
            #     data_line.append(E_tmp[2*i +2])
            for i in range(1,len(our_dic)):
                if our_dic[i] in E_tmp: data_dic[our_dic[i]] = E_tmp[E_tmp.index(our_dic[i])+1]
                else :data_dic[our_dic[i]]=-999
            deviceUID = data_dic['deviceUID']
            if deviceUID in sheet_list:
                data_table[sheet_list.index(deviceUID)].append(data_dic)
            else:
                sheet_list.append(deviceUID)
                data_table.append([data_dic])
            # data_table.append(data_dic)
            # print(data_line)
            pass
    return data_table

# def write_data(dataTemp, table):
#     print('write data')
#     [h, l] = dataTemp.shape  # h为行数，l为列数
#     for i in range(h):
#         for j in range(l):
#             table.write(i,j, dataTemp[i][j])

def write_cvs(data,file_name):
    # out = open(file_name, 'a', newline='')
    with open(file_name,'a', newline='') as csvfile:
    # csv_write = csv.writer(out, dialect='excel')
        writer = csv.writer(csvfile)
        writer.writerow(our_dic)
        csv_write = csv.DictWriter(csvfile,fieldnames=our_dic ,dialect='excel' )

        for line in data:
            csv_write.writerow(line)
    print('write out')
    pass

def write_excel(data, file_name):
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook(file_name)
    for deviceUID in range(len(sheet_list)):
        worksheet = workbook.add_worksheet(sheet_list[deviceUID])

        bold = workbook.add_format({'bold': True})

        # Expand the first columns so that the dates are visible.
        worksheet.set_column('A:H', 30)

        # Write the column headers.
        for i in range(len(our_dic)):
            temp_char =chr(ord('A')+i)+'1'
            worksheet.write(temp_char, our_dic[i], bold)
        # worksheet.write('A1', our_dic[0], bold)

        # Start from the first cell. Rows and columns are zero indexed.
        row = 1

        # Iterate over the data and write it out row by row.
        for row in range(len(data[deviceUID])):
            temp_row = []
            col = 0
            for column in range(len(our_dic)):
                temp_row.append(data[deviceUID][row][our_dic[column]])
            for item in (temp_row):
                worksheet.write(row+1, col, item)
                col += 1
            # row += 1

    print('save data')
    workbook.close()


def searchDic():
    dir = 'D:\FIles\homework\intership\dataHandler\丰科万达工程\丰科万达'
    data_table=[]
    if os.path.exists(dir):
        dirs = os.listdir(dir)
        for dirc in dirs:
            temp_dic = dir + '\\' + dirc + '\\CHU.txt'
            print(temp_dic)
            data_table = read_txt(temp_dic,data_table)
        else:
            print("dir not exists")
    # write_cvs(data_table,"excel.csv")
    write_excel(data_table,"CHU.xlsx")

# read_txt('CCS.txt')
searchDic()
# write_cvs(np.array(read_txt('CCS.txt')),"excel.csv")