import re
from blockextract_from_docx_ans import extract_block
import write_to_xlsx_multi_ans,write_to_xlsx_single_ans,write_to_xlsx_judge_ans
def extract_content(aimed_block):
    """形参接受字符串，通过正则表达式，最终print出来

    Args:
        aimed_block (str): 匹配"1．内皮："中"内皮"并打印。
    # """
    matches_of_list_quest = []


    #单选题
    if aimed_block[1]:
        pattern2_for_ans = r'[A-F]'    #如果有配伍题，此时配伍题就存在选项E和F
        matches_for_ans  = re.findall(pattern2_for_ans,aimed_block[1])
        matches_of_list_quest = matches_for_ans
        type_of_choice =  ['1']*len(matches_for_ans)

        print('\n这是单选题题目提取：',matches_for_ans)

        ques_choice_list = (matches_of_list_quest, 
                            type_of_choice
                            )
        write_to_xlsx_single_ans.write_to_xlsx(ques_choice_list)    
    #清空列表
        matches_of_list_quest.clear()

    #多选题
    if aimed_block[2]:
        pattern2_for_ans = r'\d+[\.\．](.*?)\s'    
     

        matches_of_quest = re.findall(pattern2_for_ans, aimed_block[2])
        matches_of_list_quest = matches_of_quest
        type_of_choice =  ['2']*len(matches_of_quest)

        print('\n这是多选题题目提取：',matches_of_quest)

        ques_choice_list = (matches_of_list_quest, 
                            type_of_choice
                            )
        write_to_xlsx_multi_ans.write_to_xlsx(ques_choice_list)
        #清空list
        matches_of_list_quest.clear()

    # #如果有判断题
    # if aimed_block[0]: 
    #     pattern2_for_ans = r'[XV]|[×√]'
    #     matches_for_ans  = re.findall(pattern2_for_ans,aimed_block[0])
    #     matches_of_list_quest = matches_for_ans
    #     type_of_choice =  ['1']*len(matches_for_ans)

    #     print('这是判断题答案提取：',matches_for_ans)

    #     ques_choice_list = (matches_of_list_quest, 
    #                         type_of_choice
    #                         )
    #     write_to_xlsx_judge_ans.write_to_xlsx(ques_choice_list)    
    # #清空列表
    #     matches_of_list_quest.clear()


##测试函数
if __name__ == '__main__': 
    A = extract_content(extract_block('客观题答案2.docx'))   
