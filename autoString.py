# 读写excel工作表
import  xdrlib ,sys
import xlrd
import json


filename='tanslation.xlsx'

def open_excel(file=filename):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception as e:
        print("文件打开失败,str(e)是",str(e))

#根据索引获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的所以  ，by_index：表的索引
def excel_table_byindex(file=filename,):
    data = open_excel(file)
    table = data.sheets()[0]
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    colnames =  table.row_values(0) #某一行数据
    list =[]

    for index in range(nrows):
        oneRow=table.row_values(index)
        if oneRow[0]!="":
         list.append(oneRow)

    unFindList=[]

    myJson = json.load(open("intl_en.json"))
    for findindex in myJson:
        for value in list:
            if value[0] == myJson[findindex]:
                myJson[findindex]=value[1]
            else:
                unFindList.append(findindex)
    with open("newJson.json", 'w') as f:
        json.dump(myJson, f, ensure_ascii=False)

excel_table_byindex("tanslation.xlsx")

