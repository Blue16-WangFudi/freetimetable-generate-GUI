
import re
#需要替换成pandas
import openpyxl
def parse_range(pair_str):
    temp=pair_str.split("-")
    ret=[]
    for i in temp:
        ret.append(int(i))
    return ret

#对于单个课程明细（一个String），返回正则表达式匹配到的文本块，类似于“周数：A-B”，也就是提取课程周次，EG：['9-10', '11']
def parse_weekrange(str):
    section=re.findall("(?<=周数: ).*?(?=周)",str)
    return section

#返回一个课程项目的list套list：一个课程项目：[String course_section, [String week_section1,String week_section2 ……]]
#V3.1开始，最后一条记录为['divide',1,2,3,……]用于分隔每一周
def parse_courselist_xlsx(xlsx_path):
    ret=[]

    #xlsx操作
    coursebook=openpyxl.load_workbook(xlsx_path)
    coursebook_sheet=coursebook['Sheet1']

    #获取最大行数
    max_row=coursebook_sheet.max_row-1
    
    courserange=""
    weekrange=[]
    cell_id=""
    cell_id_before=""
    for i in range(1,max_row+1):
        cell_id="A"+str(i)
        cell_id_before="A"+str(i-1)
        if(coursebook_sheet[cell_id].value!=None):
            #添加上一个
            temp=[courserange,weekrange]
            ret.append(temp)
            weekrange=[]
            #更新courserange
            courserange=coursebook_sheet[cell_id].value
        cell_id="B"+str(i)
        #print("#",cell_id,get_weekrange(coursebook_sheet[cell_id].value))
        for tmp in parse_weekrange(coursebook_sheet[cell_id].value):
            weekrange.append(tmp) 
         
    
    #添加最后一个
    temp=[courserange,weekrange]
    ret.append(temp)
    weekrange=[]
    #删除第一个空的
    ret.pop(0)
    #print("xlsx",ret)
    #添加最后一行，也就是divide。读取最多7个
    list_cellID=['A','B','C','D','E','F','G']
    list_divide=[]
    count=1
    for tmp in list_cellID:
        cell_id=list_cellID[count-1]+str(max_row+1)
        if(coursebook_sheet[cell_id].value!=None):
            list_divide.append(int(coursebook_sheet[cell_id].value))
        else:
            list_divide.append(0)
        
        count=count+1
    
    temp=["divide",list_divide]
    ret.append(temp)
    return ret