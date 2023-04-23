"""
此算法将多选题题目和选项包裹在一个result中,形式为[["题目1","A.xx"],["题目2","A.xx"]]
"""
import re
# 定义正则表达式
#题目提取
digit_pattern = re.compile(r'^\d+[．.].*')
#解析
infer_pattern = re.compile(r'^解析.*\n')   
#答案
ans_pattern = re.compile(r'^答案.*\n') 


#result = [[[题目1，选项1，2，3，4][答案1]]...[[题目2，选项1，2，3，4][答案2]]]
#result[i]--第i题目
#result[i][0]--第i题目的“题目和选项”部分，利用遍历可以获得题目和选项
#result[i][1]--第i题目的“答案和解析”部分，利用遍历可以获得答案和解析

def judge_block(line):
    result = []
    for i,item in enumerate(line):
        if digit_pattern.match(item):
            #形成一个二维空数组
            current_list = [[],[]]  
            #向二维数组中加入题目
            current_list[0].append(item)
            current_list[0].append("正确")
            current_list[0].append("错误")
            #将二维空数组加载到result
            result.append(current_list)
            #向二维数组中加入答案
        elif ans_pattern.match(item):
            current_list[1].append(item)
            #向二维数组中加入解析
        elif infer_pattern.match(item):
            current_list[1].append(item)
        else:
            index = i
            break

    return result,index
