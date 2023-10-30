import module_parse

#数据库操作：写一条记录到对应Cell，week_range分两种，如果为空(请勿提交null)，则无，如果有范围，则记作：办公室 张三(5-7/8-9)
def set_record(sheet,department,name,week_num,section_num,week_range):
    #num均从1开始，方便观察;索引从0开始递增
    week_list=['C','D','E','F','G','H','I']
    section_list=list(range(2,16))#左开右闭一定注意
    msg_week_range=""
    if(week_range==""):
        msg_week_range=""
    else:
        msg_week_range="("
        count=0
        for temp in week_range:
            count=count+1
            if(count==len(week_range)):
                msg_week_range=msg_week_range+temp
            else:
                msg_week_range=msg_week_range+temp+"/"
            
        msg_week_range=msg_week_range+")"
    
    cell_id=week_list[week_num-1]+str(section_list[section_num-1])
    if sheet[cell_id].value==None:
        sheet[cell_id]=department+' '+name+msg_week_range+'，'
    else:
        sheet[cell_id]=sheet[cell_id].value+department+' '+name+msg_week_range+'，'

def output_member_empty(sheet,department,name,courselist):
    pre=courselist[0]
    week=1
    mark=[0]*14
    if_divide=courselist[len(courselist)-1][0]=="divide"
    if(if_divide):
        list_divide=courselist[len(courselist)-1][1]
        courselist.pop(len(courselist)-1)#删去divide记录
        accumulate=list_divide[0]
    currentrecord=1
    #第一轮循环，先找出无课的，先写
    while(True):
        #拿到一条课程，先判断是否是一定有课，如果一定有课，则直接填
        #这里无需遍历课程周次范围，因为如果是一直有课，那么就有且只有一个周次段，也就是6-16或者6-17，只需要读取list第一个元素index=0
        if(currentrecord<=len(courselist)):
            temp=courselist[currentrecord-1]
        #这里是读取当前的(tempA)和前一个(tempB)的数据

        
        tempA=module_parse.parse_range(temp[0])
        tempB=module_parse.parse_range(pre[0])


        #判断是否切换到下一天，切换的同时写入全空的记录
        if(if_divide and accumulate<currentrecord):
            count=0
            for i in mark:
                count=count+1
                if(i==0): 
                    set_record(sheet,department,name,week,count,"")
            week=week+1
            mark=[0]*14#为下一天做准备
            if(week<=7):
                accumulate=accumulate+list_divide[week-1]
        else: 
            if(if_divide==False and tempA[1]<tempB[1]):#后面的比前面的小，那就是跳天了,跳天直接开始写当日的空
                #这里本来是0的，可是课表有改变，所以这个方法有风险
                count=0
                for i in mark:
                    count=count+1
                    if(i==0): 
                        set_record(sheet,department,name,week,count,"")
                week=week+1
                mark=[0]*14#为下一天做准备
        
        #判断什么时候跳出循环,多一次进行收尾
        if(currentrecord>len(courselist)):
            break
        
        
        temp_range=module_parse.parse_range(temp[0])#解析出对应课程覆盖时间段
        #把有课的置为1
        for i in range(temp_range[0],temp_range[1]+1):
            mark[i-1]=1
        #顺便可以分析哪些是间断性的有课
        if(temp[1][0]!="6-16" and temp[1][0]!="6-17"):
            temp_range=module_parse.parse_range(temp[0])#解析出对应课程覆盖时间段
            for i in range(temp_range[0],temp_range[1]+1):
                set_record(sheet,department,name,week,i,temp[1])
        pre=temp#保存当前课程条目，供下一个比对确定是第几周
        currentrecord=currentrecord+1