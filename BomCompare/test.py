import BomStruct
import xlsxwriter
import compareBOM
from xlsxwriter.workbook import Workbook

def compareManufatoryBOMs(newBom,oldBom,excel):
    print("start")
    list = BomStruct.ExtractInfoFromExcel(oldBom)
    list310 = BomStruct.ExtractInfoFromExcel(newBom)
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
    
    difflist=compareBOM.compareBoms(list310, list)
    difflist1=compareBOM.compareBoms(list,list310)  
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
    
if __name__=="__main__":
    compareManufatoryBOMs("BOMReport202.xls","BOMReport302.xls","diff.xlsx")
    #print(difflist)    
        