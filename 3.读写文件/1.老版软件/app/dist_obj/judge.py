import re
def judge_block(line):
    result = ''
    i = 0
    while i < len(line):
        if re.match(r'^\d+[．.].*', line[i]):  #匹配1.XX内容
            result += line[i]
            i += 1
        else:
            break
    return result