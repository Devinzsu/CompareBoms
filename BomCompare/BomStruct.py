import xlrd
import chardet
import codecs
import re
from lib2to3.fixer_util import String
from _codecs import decode
import xlsxwriter

class compareTwoBOMs():
    def __init__(self,primaryBomPath,secondBomPath,primaryRiskbuyBomPath):
        self.primaryBomExcel=primaryBomPath
        self.secondBOMExcel=secondBomPath
        self.primaryRiskbuyExcel=primaryRiskbuyBomPath
        self.primaryBom=[]
        self.secondBom=[]
        self.primaryRiskBuyBom=[]
    def getDatafromBom(self):
        self.primaryBom=self.ExtractInfoFromExcel(self.primaryBomExcel)
        self.secondBom=self.ExtractInfoFromExcel(self.secondBOMExcel)
        self.primaryRiskBuyBom=self.BomStruct(self.primaryRiskbuyExcel)
    def compareBoms(self,first,second):
        refDes1=first[4].index("Ref Des")
        partNum=second[4].index("Number")
        changelist=[]
        addedlist=[]
        for index in range(len(first)):
            components=first[index][partNum]
            if re.search("\d+-\d+-\d+", str(components)) and len(first[index][refDes1])>=1:
                temp=self.findCompont(first[index], second)
                changelist.append(temp[0])
                addedlist.append(temp[1])            
        return [changelist,addedlist]#print(components)
    def writeExcel(self,excel):
        print("start")
        #list = BomStruct.ExtractInfoFromExcel(oldBom)
        #list310 = BomStruct.ExtractInfoFromExcel(newBom)
        list=self.secondBom
        list310=self.primaryBom
        print("finish")
        #print(list[4])
        excelfile=xlsxwriter.Workbook("test.xls")
        sheet1=excelfile.add_worksheet()
        for row in range(len(list)):        
            for col in range(len(list[row])):
                #print(str(list[row][col]))
                sheet1.write(row,col,str(list[row][col]))
        excelfile.close()
        excelfile1=xlsxwriter.Workbook(excel)
        sheet1_excel1=excelfile1.add_worksheet()
        
        difflist=self.compareBoms(list310, list)
        difflist1=self.compareBoms(list,list310)  
        #print(difflist)
        sheet1_excel1.write_row(0, 0, ["Updated components"])  
        sheet1_excel1.write_row(1, 0, ["Ref Des","New NVPN","Description","Old NVPN","Description"])
        count=2
        #print(len(difflist))
        for each in difflist[0]:        
            #sheet1_excel1.write(count,0,k)
            for (k,v) in each.items():
                #print(k,v)
                
                sheet1_excel1.write(count,0,k)
                sheet1_excel1.write(count,1,v[0])
                sheet1_excel1.write(count,2,v[1])
                sheet1_excel1.write(count,3,v[2])
                sheet1_excel1.write(count,4,v[3])
                count=count+1
        count=0
        sheet1_add=excelfile1.add_worksheet("added")
        sheet1_add.write_row(count+1, 0, ["added components"])  
        sheet1_add.write_row(count+2, 0,  ["Ref Des","New NVPN","Description"])
        count=count+3
        for each in difflist[1]:        
            #sheet1_excel1.write(count,0,k)
            for (k,v) in each.items():
                #print(k,v)
                
                sheet1_add.write(count,0,k)
                sheet1_add.write(count,1,v[0])
                sheet1_add.write(count,2,v[1])
                count=count+1
        count=0
        sheet1_remove=excelfile1.add_worksheet("removed")
        sheet1_remove.write_row(count+1, 0, ["removed components"])  
        sheet1_remove.write_row(count+2, 0,  ["Ref Des","New NVPN","Description"])
        count=count+3
        for each in difflist1[1]:        
            #sheet1_excel1.write(count,0,k)
            for (k,v) in each.items():
                #print(k,v)
                
                sheet1_remove.write(count,0,k)
                sheet1_remove.write(count,1,v[0])
                sheet1_remove.write(count,2,v[1])
    
                count=count+1
        
    def ExtractInfoFromExcel(self,excelPath):
        excfile=xlrd.open_workbook(filename=excelPath,encoding_override='utf-8')
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
    def BomStruct(self,riskbuyreport):
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
            if re.search("\d+-\d+-\d+-\d+", str(list[each][0])):
                if str(list[each][0]) not in bomstruct:
                    bomstruct[str(list[each][0])]=[str(list[each][2])]
                else:
                    bomstruct[str(list[each][0])].append(str(list[each][2]))
        return [bomstruct,list]
    def compareBoms(self,firstBom,secondBom):
        refDes1=firstBom[4].index("Ref Des")
        partNum=secondBom[4].index("Number")
        changelist=[]
        addedlist=[]
        for index in range(len(firstBom)):
            components=firstBom[index][partNum]
            if re.search("\d+-\d+-\d+", str(components)) and len(firstBom[index][refDes1])>=1:
                temp=self.findCompont(firstBom[index], secondBom)
                changelist.append(temp[0])
                addedlist.append(temp[1])
                #print(difflist)
                #then, find the component in second BOM
                #print(firstBom[index])
            
        return [changelist,addedlist]#print(components)
    def findCompont(self,component, secondBom):
        refDes=secondBom[4].index("Ref Des")
        partNum=secondBom[4].index("Number")
        indexOfdesc=secondBom[4].index("Description")
        sourceComponent=component[refDes].replace("-",'').replace(" ",'')
        sourceComponent=sourceComponent.split(",")
        list=[]
        partlist={}
        addedlist={}
        #print(component[refDes])
        #print(sourceComponent)
        for eachCompoent in sourceComponent:
            #print(eachCompoent)
            list.append(eachCompoent)
            if eachCompoent is not None and len(eachCompoent)>1:
                #print(eachCompoent)
    #             if component[partNum]=="195-3223-000":
    #                 print(component)
                for index in range(len(secondBom)):
                    targetComponent=secondBom[index][refDes]
                    targetlist=targetComponent.replace(" ",'').split(",")
                    if str(eachCompoent) in targetlist:
                        if eachCompoent in list:
                            list.remove(eachCompoent)
                        if(component[partNum]==secondBom[index][partNum]):
                            #if eachCompoent=="Q1":
                                #print("test",secondBom[index],targetlist)
                            #if eachCompoent in list:
                                #print("Exit",secondBom[index])
                                #print(sourceComponent)
                                #print(str(eachCompoent),targetComponent)
                            #list.append(eachCompoent)
                            if eachCompoent in list:
                                list.remove(eachCompoent)
                        elif (component[partNum]!=secondBom[index][partNum]):
                            partlist[eachCompoent]=[component[partNum],component[indexOfdesc],secondBom[index][partNum],secondBom[index][indexOfdesc]]
                if(len(list)>=1 and list[0]!=''):
                    addedlist[eachCompoent]=[component[partNum],component[indexOfdesc]]
        if(len(list)>=1 and list[0]!=''):
            print("%s not exist"%list)
        return [partlist,addedlist]
        

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
              
            