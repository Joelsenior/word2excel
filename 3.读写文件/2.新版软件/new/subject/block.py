from operator import index
import re
# 定义正则表达式
digit_pattern = re.compile(r'^\d+[．.].*')
#解析
infer_pattern = re.compile(r'^解析.*\n')   
#答案
ans_pattern = re.compile(r'^答案.*\n') 
#非文字题目前很难处理，比如题目是非文字，答案是文字
#非文字匹配
# none_word = "[非文字]"


def explan(line):
    result = []
    for i,item in enumerate(line):
        if digit_pattern.match(item):
            #形成一个二维空数组
            current_list = [[],[]]  
            #向二维数组中加入题目
            current_list[0].append(item)
            #将二维空数组加载到result
            result.append(current_list)
            #向二维数组中加入答案
        elif ans_pattern.match(item):
            current_list[1].append(item)
            #向二维数组中加入解析
        elif infer_pattern.match(item):
            current_list[1].append(item)
        else:
            #此时跳出的因素一般认为是碰到了非题目、非解析、非答案。一般认为是“XX题型”，此时记住这个索引号，让下一个循环从这个index开始。
            index = i
            break

    return result,index

#简答题
def simple(line):
    result = []
    for i,item in enumerate(line):
        if digit_pattern.match(item):
            #形成一个二维空数组
            current_list = [[],[]]  
            #向二维数组中加入题目
            current_list[0].append(item)
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


#填空题
def fill(line):
    result = []
    for i,item in enumerate(line):
        if digit_pattern.match(item):
            #形成一个二维空数组
            current_list = [[],[]]  
            #向二维数组中加入题目
            current_list[0].append(item)
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



#画图题
def drawing(line):
    result = []
    for i,item in enumerate(line):
        if digit_pattern.match(item):
            #形成一个二维空数组
            current_list = [[],[]]  
            #向二维数组中加入题目
            current_list[0].append(item)
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


#计算题
def caculate(line):
    result = []
    for i,item in enumerate(line):
        # if none_word in item:
        #     print("这是非文字题")
        if digit_pattern.match(item):
            #形成一个二维空数组
            current_list = [[],[]]  
            #向二维数组中加入题目
            current_list[0].append(item)
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

#其他题型
# def others(line):
#     result = []
#     for i,item in enumerate(line):
#         if digit_pattern.match(item):
#             #形成一个二维空数组
#             current_list = [[],[]]  
#             #向二维数组中加入题目
#             current_list[0].append(item)
#             #将二维空数组加载到result
#             result.append(current_list)
#             #向二维数组中加入答案
#         elif ans_pattern.match(item):
#             current_list[1].append(item)
#             #向二维数组中加入解析
#         elif infer_pattern.match(item):
#             current_list[1].append(item)
#         else:
#             index = i
#             break

    # return result,index