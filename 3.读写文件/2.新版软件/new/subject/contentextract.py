from blockextract_from_docx import extract_block
import write_to_xlsx_all_sub
def extract_content(aimed_block):
    """形参接受字符串，通过正则表达式，最终print出来

    Args:
        aimed_block (str): 匹配"1．内皮："中"内皮"并打印。
    # """
    #填空题
    if aimed_block[0]:
         for each in aimed_block[0]:
           write_to_xlsx_all_sub.fill(each)
    #名词解释
    if aimed_block[1]:
         for each in aimed_block[1]:
           write_to_xlsx_all_sub.explan(each)
    #简答题
    if aimed_block[2]: 
        for each in aimed_block[2]:
            write_to_xlsx_all_sub.simple(each)
    #计算题
    if aimed_block[3]: 
        for each in aimed_block[3]:
            write_to_xlsx_all_sub.caculate(each)
    #画图题
    if aimed_block[4]: 
        for each in aimed_block[4]:
            write_to_xlsx_all_sub.drawing(each)
    #其他题型
    if aimed_block[5]: 
        for each in aimed_block[5]:
            write_to_xlsx_all_sub.others(each)


##测试函数
if __name__ == '__main__': 
    A = extract_content(extract_block('主观题题目.docx'))   
