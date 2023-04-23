import docx

from judge import judge_block
from single_choice import single_choice_block
from multi_choice import multi_choice_block
def extract_block(path_name,logger):
    """将.docx中包含有所有客观题的内容,提取出来,并返回字符串。

    Args:
        path_name (str): path_name为路径名+文件名
    """
    doc = docx.Document(path_name)
    lines = []
    for paragraph in doc.paragraphs:
        text = paragraph.text + '\n' 
        lines.append(text)
    print('这是逐行读取的内容：',lines)
    logger.info("逐行读取的内容:")
    logger.info(lines)
    start_flag_1 = '判断题'
    start_flag_2 = '单选题'
    start_flag_3 = '多选题'

    result1 = '' 
    result2 = '' 
    result3 = '' 
    #默认index = 0 保证从0开始可以执行。
    #随着index不断增大。
    index = 0
    for i,line in enumerate(lines):
        if i < index:
            continue
        #如果是判断题，则返回判断题标签3
        if start_flag_1 in line:
            result1 = judge_block(lines[i+1:])[0]
            index = judge_block(lines[i+1:])[1] 
            print('判断题:',result1)
            logger.info("填空题：")
            logger.info(result1)
        elif start_flag_2 in line: 
            result2 = single_choice_block(lines[i+1:])[0] 
            index = single_choice_block(lines[i+1:])[1]  
            print('\n单选题:',result2)
            logger.info("单选题：")
            logger.info(result2)
        elif start_flag_3 in line: 
            result3 = multi_choice_block(lines[i+1:])[0] 
            index = multi_choice_block(lines[i+1:])[1]   
            print('\n多选题:',result3)
            logger.info("多选题：")
            logger.info(result3)
        else:
            continue

    return result1,result2,result3

if __name__ == '__main__':
    extract_block(r'客观题题目.docx')


