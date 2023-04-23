import re
def single_choice_block(line):
    result = ''
    i = 0
    while i < len(line):
        if re.match(r'\d+.*\n', line[i]):  #匹配题目和选项
            result += line[i]
            i += 1
        else:
            break
    return result