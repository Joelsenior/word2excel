import re
def explan_of_nouns(line):
    result = ''
    i = 0
    while i < len(line):
        if re.match(r'^\d+[．.].*', line[i]):  #匹配题目和选项
            result += line[i]
            i += 1
        else:
            break
    return result