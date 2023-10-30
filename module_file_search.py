import os

import module_parse
import module_dataop as dp
def get_department(path):
    if(path==''):
        return
    list_department=os.listdir(path)
    if(list_department!=None):
        return list_department
    return ["Error"]

#搜索对应部门成员然后直接写入
def set_department_member(wb_sheet,path,department):
    member_file_list=os.listdir(path+"//"+department)
    member_list=[]
    #查找出人员数量
    for temp in member_file_list:
        temp2=temp.split('.')
        if(temp2[1]=='xlsx'):
            member_list.append(temp)
    #获取到了member_list，下面开始
    print('查找到的成员文件:',member_list)
    for temp in member_list:
        temp2=temp.split('(')
        print('当前写入部门：',department,"当前写入人员：",temp2[0])
        #准备工作
        sourcepath=path+'//'+department+'//'+temp
        #destpath=path+'//'+department+'//'+temp2[0]+'.csv'
        temp3=temp.split('.')
        type=temp3[1]
        if(type=="xlsx"):
            dp.output_member_empty(wb_sheet,department,temp2[0],module_parse.parse_courselist_xlsx(sourcepath))
