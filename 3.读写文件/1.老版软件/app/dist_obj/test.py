import re

# 原始列表
lst = [ "123","A", "B", "C","456",'\n',  "789", "A", "B", "C", '\n']

# 定义正则表达式
digit_pattern = re.compile(r"\d+")
letter_pattern = re.compile(r"[A-E]")

# 遍历列表，根据正则表达式匹配情况将元素存入不同的列表
current_list = []  #搬运工
result = []
for item in lst:
    if digit_pattern.match(item):
        current_list = [item]
        result.append(current_list)
    elif letter_pattern.match(item):
        current_list.append(item)
        
print(result)
