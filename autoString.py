import xlrd
import json

filename='tanslation.xlsx'
def open_excel(file=filename):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception as e:
        print("文件打开失败,str(e)是",str(e))

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
    findList=[]
    myJson = json.load(open("intl_en.json"))
    rawJson=myJson
    for findindex in myJson:
        for value in list:
            if str(value[0]).strip()== str(myJson[findindex]).strip():
                myJson[findindex]=value[1]
                findList.append(findindex)

    with open("intl_tl.json", 'w') as f:
        json.dump(myJson, f, ensure_ascii=False)

    for value in findList:
        rawJson.pop(value)
    print("未匹配的字符  "+str(rawJson))
    with open("unfind.json", 'w') as f:
        json.dump(rawJson, f, ensure_ascii=False)



excel_table_byindex("tanslation.xlsx")

