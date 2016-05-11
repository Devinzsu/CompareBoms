import xlrd
import chardet
import codecs
import re
from lib2to3.fixer_util import String
from _codecs import decode
import xlsxwriter

def ExtractInfoFromExcel(location):
    excfile=xlrd.open_workbook(filename=location,encoding_override='utf-8')
    worksheet=excfile.sheet_by_index(0)
    nrows=worksheet.nrows
    ncols=worksheet.ncols
    list=[]
    for rownum in range(1,nrows):
        row=worksheet.row_values(rownum)
        if row:
           data=[]
           for each in range(len(row)):
               try:
                   #print(row[each])
                   data.append(row[each])
               except UnicodeEncodeError:
                   slashUStr=row[each]
                   decodedstr=codecs.decode(slashUStr,"unicode-escape" )
                   decodedstrtoGBK=decodedstr.encode("GBK","ignore")
                   #print(decodedstrtoGBK)
                   data.append(decodedstrtoGBK)
           list.append(data)
    return list

def BomStruct(riskbuyreport): 
    excfile=xlrd.open_workbook(filename=riskbuyreport,encoding_override='utf-8')
    worksheet=excfile.sheet_by_index(0)
    ncols=worksheet.ncols
    nrows=worksheet.nrows
    list=[]
    for rownum in range(1,nrows):
        row=worksheet.row_values(rownum)
        
        if row:
           data=[]
           for each in range(len(row)):
               try:
                   #print(row[each])
                   data.append(row[each])
               except UnicodeEncodeError:
                   slashUStr=row[each]
                   decodedstr=codecs.decode(slashUStr,"unicode-escape" )
                   decodedstrtoGBK=decodedstr.encode("GBK","ignore")
                   #print(decodedstrtoGBK)
                   data.append(decodedstrtoGBK)
           list.append(data)
    bomstruct={}
    for each in range(len(list)):
        #print(list[each][0])
        #print(re.search("\d+-\d+-\d+-\d+", str(list[each][0])))
        if re.search("\d+-\d+-\d+-\d+", str(list[each][0])):
            if str(list[each][0]) not in bomstruct:
                bomstruct[str(list[each][0])]=[str(list[each][2])]
            else:
                bomstruct[str(list[each][0])].append(str(list[each][2]))
    #print(bomstruct)
    return [bomstruct,list]
              
            