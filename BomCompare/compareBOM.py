import re
from macpath import split
def compareBoms(firstBom,secondBom):
    refDes1=firstBom[4].index("Ref Des")
    partNum=secondBom[4].index("Number")
    changelist=[]
    addedlist=[]
    for index in range(len(firstBom)):
        components=firstBom[index][partNum]
        if re.search("\d+-\d+-\d+", str(components)) and len(firstBom[index][refDes1])>=1:
            temp=findCompont(firstBom[index], secondBom)
            changelist.append(temp[0])
            addedlist.append(temp[1])
            #print(difflist)
            #then, find the component in second BOM
            #print(firstBom[index])
        
    return [changelist,addedlist]#print(components)
def findCompont(component, secondBom):
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

#     if len(partlist)>0:
        #print(partlist.keys())
#         print(partlist)
#     if(len(partlist)>0):
#         print((partlist))
    if(len(list)>=1 and list[0]!=''):
        print("%s not exist"%list)
    return [partlist,addedlist]
        
            
            
         