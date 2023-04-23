import re
# def single_choice_block(line):
#     result = ''
#     i = 0
#     while i < len(line):
#         if re.match(r'^\d+[．.].*', line[i]) or re.match(r'[A-D].*\n',line[i]):  #匹配题目和选项
#             result += line[i]
#             i += 1
#         else:
#             break
#     return result

# 定义正则表达式
digit_pattern = re.compile(r'^\d+[．.].*')
letter_pattern = re.compile(r'[A-F].*\n')  #有可能有配伍题，因此需要包含E和F
current_list = []  #搬运工
result = []
# 遍历列表，根据正则表达式匹配情况将元素存入不同的列表
def single_choice_block(line):
    for item in line:
        if digit_pattern.match(item):
            current_list = [item]
            result.append(current_list)
        elif letter_pattern.match(item):
            current_list.append(item)
        else:
            break

    return result