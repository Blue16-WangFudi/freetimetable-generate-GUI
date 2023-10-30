#encoding=utf-8
#author: blue16（Index:blue16.cn）
#date: 2023-10-29
#summary: 一个带GUI且分模块的无课表制作程序。只有一个小时，这么大的代码量，我觉得如果三个高手一起都不可能完成，还是算了。

#GUI库文件
import PySimpleGUI as sg
import openpyxl
import os
#自写模块文件
import module_file_search as gd
import module_parse as mp
import module_file_search as fs
 
# 布局

layout=[[sg.Text("选择输入文件夹：      "),sg.In(key = '-Input-',enable_events=True),sg.FolderBrowse(button_text = "选择目录",target = '-Input-')],
        [sg.Text("找到的部门：",key='-Prompt_Department-')],
        [sg.Text("选择输出文件路径：    "),sg.In(key = '-Output-'),sg.FileSaveAs(button_text = "选择路径",target = '-Output-')],
        [sg.Button("开始生成",key = '-Start-')],
        [sg.Text("解析输出：")],
        [sg.Text("Output",key='-Prompt_Output-')]

] 
# 创建窗口
window = sg.Window('无课表生成工具', layout)
 
# 事件循环
while True:
    event, values = window.read()

    #这里有一个问题：反复点击编辑框会导致反复查找
    if(event=="-Input-"):
        list_department=gd.get_department(values['-Input-'])
        print("list_department=",list_department)
        str=""
        count=1
        for temp in list_department:
            if(count==len(list_department)):
                str=str+temp
            else:
                str=str+temp+","
            count=count+1
        window['-Prompt_Department-'].update("找到的部门："+str)
    if(event=="-Start-"):
        
        wb=openpyxl.load_workbook('output.xlsx')
        wb_sheet=wb['Sheet']
        department_list=os.listdir(values['-Input-'])
        print('查找到的部门：',department_list)
        for temp in department_list:
            fs.set_department_member(wb_sheet,values['-Input-'],temp)
        wb.save(values['-Output-'])
        window['-Prompt_Output-'].update("Finished.")

    if event == sg.WIN_CLOSED or event == 'Submit':
        break

# 关闭窗口
window.close()
