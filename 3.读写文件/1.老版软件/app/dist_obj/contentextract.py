import re
from blockextract_from_docx import extract_block
from write_to_xlsx import write_to_xlsx
import write_to_xlsx_multiiiii,write_to_xlsx_singleeee
def extract_content(aimed_block):
    """形参接受字符串，通过正则表达式，最终print出来

    Args:
        aimed_block (str): 匹配"1．内皮："中"内皮"并打印。
    # """
    matches_of_list_quest = []
    matches_of_list_A = []
    matches_of_list_B = []

    pattern2_for_ans = r'\d+[．.](.*?)\n'      # pattern2_for_ans题目提取，根据回车作为结尾符，读取。

    #单选题
    if aimed_block[1]:
         for each in aimed_block[1]:
        #将单题的题目和选项分开
            write_to_xlsx_singleeee.write_to_xlsx(each)

    #多选题
    if aimed_block[2]:
        #将单题提出，并用（）封装
        for each in aimed_block[2]:
        #将单题的题目和选项分开
            write_to_xlsx_multiiiii.write_to_xlsx(each)

    #如果有判断题
    if aimed_block[0]: 
        matches_of_quest = re.findall(pattern2_for_ans, aimed_block[0])
        matches_of_list_quest = matches_of_quest
        matches_of_A = ['正确']*len(matches_of_quest)
        matches_of_list_A = matches_of_A
        matches_of_B = ['错误']*len(matches_of_quest)
        matches_of_list_B = matches_of_B
        print('这是判断题题目提取：',matches_of_quest)
        # print('这是选项A提取：',matches_of_A)
        # print('这是选项B提取：',matches_of_B)
        ques_choice_list = (matches_of_list_quest, 
                            matches_of_list_A,
                            matches_of_list_B,

                            )
        write_to_xlsx(ques_choice_list)
        #清空list
        matches_of_list_quest.clear()
        matches_of_list_A.clear()
        matches_of_list_B.clear()


##测试函数
if __name__ == '__main__': 
    A = extract_content(extract_block('客观题题目.docx'))   
