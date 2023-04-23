import re
from blockextract_from_docx_ans import extract_block
from write_to_xlsx_ans import write_to_xlsx
def extract_content(aimed_block):
    """形参接受字符串，通过正则表达式，最终print出来

    Args:
        aimed_block (str): 匹配"1．内皮："中"内皮"并打印。
    # """
    matches_of_list_quest = []
    input_type = []
    pattern2_for_ans = r'\d+[．.](.*?)\n'      # pattern2_for_ans题目提取，根据回车作为结尾符，读取。

    #填空题
    if aimed_block[0]:
        pattern2_for_ans = r'\d+[．.](.*?)\n'      # pattern2_for_ans题目提取，根据回车作为结尾符，读取。

        matches_of_quest = re.findall(pattern2_for_ans, aimed_block[0])
        matches_of_list_quest = matches_of_quest
        input_type = ['1']*len(matches_of_quest)
        print('这是填空题答案提取：',matches_of_quest)
        ques_choice_list = (matches_of_list_quest,input_type)
        write_to_xlsx(ques_choice_list)    
    #清空列表
        matches_of_list_quest.clear()
        input_type.clear()

    #名词解释
    if aimed_block[1]:
        pattern2_for_ans = r'\d+[．.](.*?)\n'      # pattern2_for_ans题目提取，根据回车作为结尾符，读取。

        matches_of_quest = re.findall(pattern2_for_ans, aimed_block[1])
        matches_of_list_quest = matches_of_quest
        input_type = ['1']*len(matches_of_quest)
        print('这是名词解释答案提取：',matches_of_quest)
        ques_choice_list = (matches_of_list_quest,input_type)
        write_to_xlsx(ques_choice_list)
        #清空list
        matches_of_list_quest.clear()
        input_type.clear()

    #简答题
    if aimed_block[2]: 
        matches_of_quest = re.findall(pattern2_for_ans, aimed_block[2])
        matches_of_list_quest = matches_of_quest
        input_type = ['1']*len(matches_of_quest)
        print('这是简答题答案提取：',matches_of_quest)
        ques_choice_list = (matches_of_list_quest,input_type)
        write_to_xlsx(ques_choice_list)
        #清空list
        matches_of_list_quest.clear()
        input_type.clear()

    # return matches_of_quest,matches_of_A,matches_of_B,matches_of_C, matches_of_D ,matches_of_E,matches_of_F#返回的是1*5 的 cell
    # return matches_of_list_quest,matches_of_list_A,matches_of_list_B,matches_of_list_C,matches_of_list_D,matches_of_list_E,matches_of_list_F
##测试函数
if __name__ == '__main__': 
    A = extract_content(extract_block('test.docx'))   
