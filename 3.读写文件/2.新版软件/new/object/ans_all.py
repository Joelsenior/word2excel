
import pandas as pd
def xlsx_single_ans(ans,writer,index,length = 6,anal = False):
    """
    鉴于答案和解析写入的复杂程度，将单选、多选、判断三种题型答案解析提取方法总结在这个包里。
    注意这个包只包含所有答案和解析写入的方法，不包括题目和选项提取。
    """
    raw_index = [5,8,11,14,17,20]
    using_index = raw_index[:length]
    if anal == False:
        right_ans = pd.DataFrame([1]) 
        wrong_ans = pd.DataFrame([2]) 
        #ans是["A"]需要取值，因此选择ans[0]
        if ans[0] == 'A':
            #如果答案是A,则将1写入5行，2写入8,1,1,14行
            right_ans.to_excel(writer, sheet_name='Sheet1', startrow=index,startcol=5, index=False, header=False)  #head = flase 就是不用写表格头
            using_index.pop(0)  #删除正确选项，留下错误选项
            for i in using_index:  #若length = 6 ,则形成1~5
                wrong_ans.to_excel(writer, sheet_name='Sheet1', startrow=index,startcol=i, index=False, header=False)                   
        if ans[0] == 'B':
            right_ans.to_excel(writer, sheet_name='Sheet1', startrow=index,startcol=8, index=False, header=False)  #head = flase 就是不用写表格头
            using_index.pop(1)  #删除正确选项，留下错误选项
            for i in using_index:  #若length = 6 ,则形成1~5
                wrong_ans.to_excel(writer, sheet_name='Sheet1', startrow=index,startcol=i, index=False, header=False)     
        if ans[0] == 'C':
            right_ans.to_excel(writer, sheet_name='Sheet1', startrow=index,startcol=11, index=False, header=False)  #head = flase 就是不用写表格头
            using_index.pop(2)  #删除正确选项，留下错误选项
            for i in using_index:  #若length = 6 ,则形成1~5
                wrong_ans.to_excel(writer, sheet_name='Sheet1', startrow=index,startcol=i, index=False, header=False)     
        if ans[0] == 'D':    
            right_ans.to_excel(writer, sheet_name='Sheet1', startrow=index,startcol=14, index=False, header=False)  #head = flase 就是不用写表格头
            using_index.pop(3)  #删除正确选项，留下错误选项
            for i in using_index:  #若length = 6 ,则形成1~5
                wrong_ans.to_excel(writer, sheet_name='Sheet1', startrow=index,startcol=i, index=False, header=False)      
        #配伍题 
        if ans[0] == 'E':    
            right_ans.to_excel(writer, sheet_name='Sheet1', startrow=index,startcol=17, index=False, header=False)  #head = flase 就是不用写表格头
            using_index.pop(4)  #删除正确选项，留下错误选项
            for i in using_index:  #若length = 6 ,则形成1~5
                wrong_ans.to_excel(writer, sheet_name='Sheet1', startrow=index,startcol=i, index=False, header=False)       
        if ans[0] == 'F':    
            right_ans.to_excel(writer, sheet_name='Sheet1', startrow=index,startcol=20, index=False, header=False)  #head = flase 就是不用写表格头
            using_index.pop(5)  #删除正确选项，留下错误选项
            for i in using_index:  #若length = 6 ,则形成1~5
                wrong_ans.to_excel(writer, sheet_name='Sheet1', startrow=index,startcol=i, index=False, header=False)   
        #这里将写入解析，也可能会因为答案不是上述内容，而导致写入到解析的错误
    if anal == True:
        anal_data = pd.DataFrame(ans)
        anal_data.to_excel(writer, sheet_name='Sheet1', startrow=index,startcol=23, index=False, header=False)

def xlsx_multi_ans(ans,writer,index,length = 6,anal = False):
    if anal == False:
        right_ans = pd.DataFrame([1]) 
        wrong_ans = pd.DataFrame([2]) 
        raw_all_choice = {'A': 5, 'B': 8, 'C': 11, 'D': 14, 'E': 17, 'F': 20} 
        #将字典切片，随着选项长度缩短而对相应列填充2
        all_choice = {k: v for i, (k, v) in enumerate(raw_all_choice.items()) if i < length}
        list1 = []  #正确选项的对应字典中的value
        list2 = []  #不正确选项的对应字典中的value
        #获得匹配list
        for choice in all_choice:
            if choice in ans[0]:
                list1.append(all_choice[choice])
            else:
                list2.append(all_choice[choice]) 
        # print(list1,list2)                      
        for i in list1:    
            right_ans.to_excel(writer, sheet_name='Sheet1', startrow=index,startcol=i, index=False, header=False)
        for j in list2:   
            wrong_ans.to_excel(writer, sheet_name='Sheet1', startrow=index,startcol=j, index=False, header=False)                 
    
    #有解析就写入解析
    if anal==True:
        anal_data = pd.DataFrame(ans)
        anal_data.to_excel(writer, sheet_name='Sheet1', startrow=index,startcol=23, index=False, header=False)

def xlsx_judge_ans(ans,writer,index,anal=False):
    right_ans = pd.DataFrame([1]) 
    wrong_ans = pd.DataFrame([2]) 
    #ans是["A"]需要取值，因此选择ans[0]\
    #如果正确
    if anal == False:
        if ans[0] == '√' or ans[0] == 'v' or ans[0] == 'V'or ans[0] == '对':
            right_ans.to_excel(writer, sheet_name='Sheet1', startrow=index,startcol=5, index=False, header=False)  #head = flase 就是不用写表格头
            wrong_ans.to_excel(writer, sheet_name='Sheet1', startrow=index,startcol=8, index=False, header=False) 
        #如果错误               
        if ans[0] == '×'or ans[0] == 'x' or ans[0] == 'X'or ans[0] == '错':
            right_ans.to_excel(writer, sheet_name='Sheet1', startrow=index,startcol=8, index=False, header=False)  #head = flase 就是不用写表格头
            wrong_ans.to_excel(writer, sheet_name='Sheet1', startrow=index,startcol=5, index=False, header=False) 

    
    #有解析就写入解析
    if anal == True:
        anal_data = pd.DataFrame(ans)
        anal_data.to_excel(writer, sheet_name='Sheet1', startrow=index,startcol=23, index=False, header=False)