import re
import docx

from fill_in_blanks_ans import fill_block
from explan_of_nouns_ans import explan_of_nouns
from short_ans_ques_ans import short_ans_ques
def extract_block(path_name):
    """将.docx中包含有所有客观题的内容,提取出来,并返回字符串。

    Args:
        path_name (str): path_name为路径名+文件名
    """
    # with open(path_name, 'r', encoding='utf-8') as f:
    doc = docx.Document(path_name)
    lines = []
    for paragraph in doc.paragraphs:
        text = paragraph.text + '\n' 
        lines.append(text)
    print('这是逐行读取的内容：',lines)
    start_flag_1 = '填空题'
    start_flag_2 = '名词解释'
    start_flag_3 = '简答题'
    start_flag_4 = '计算题'

    result1 = '' 
    result2 = '' 
    result3 = '' 
    for i,line in enumerate(lines):
        #如果是填空题，则返回判断题标签3
        if start_flag_1 in line:
            result1 = fill_block(lines[i+1:])  
            print('填空题:',result1)
        elif start_flag_2 in line: 
            result2 = explan_of_nouns(lines[i+1:])  
            print('\n名词解释',result2)
        elif start_flag_3 in line: 
            result3 = short_ans_ques(lines[i+1:])  
            print('\n简答题:',result3)
        else:
            continue

    return result1,result2,result3

if __name__ == '__main__':
    extract_block(r'主观题答案.docx')


