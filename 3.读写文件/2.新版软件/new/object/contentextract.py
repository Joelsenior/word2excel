import re
from blockextract_from_docx import extract_block
import write_to_xlsx_all
def extract_content(aimed_block):
    """形参接受字符串，通过正则表达式，最终print出来

    Args:
        aimed_block (str): 匹配"1．内皮："中"内皮"并打印。
    # """

    #如果有判断题
    if aimed_block[0]: 
        for each in aimed_block[0]:
           write_to_xlsx_all.judge_write(each)
    #单选题
    if aimed_block[1]:
         for each in aimed_block[1]:
           write_to_xlsx_all.single_write(each)

    #多选题
    if aimed_block[2]:
        for each in aimed_block[2]:
           write_to_xlsx_all.multi_write(each)



##测试函数
if __name__ == '__main__': 
    A = extract_content(extract_block('客观题题目.docx'))   
