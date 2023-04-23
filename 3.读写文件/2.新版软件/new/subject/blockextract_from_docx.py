import docx

from set_logs import setup_logger
import block


def extract_block(path_name,logger):
    """将.docx中包含有所有客观题的内容,提取出来,并返回字符串。

    Args:
        path_name (str): path_name为路径名+文件名
    """
    #日志
    # logger = setup_logger()
    doc = docx.Document(path_name)
    lines = []
    for paragraph in doc.paragraphs:
        text = paragraph.text + '\n' 
        lines.append(text)
    print('这是逐行读取的内容：',lines)
    logger.info("逐行读取的内容:")
    logger.info(lines)
    start_flag_1 = '填空题'
    start_flag_2 = '名词解释'
    start_flag_3 = '简答题'
    start_flag_4 = '计算题'
    start_flag_5 = '画图题'
    #增加一个其他题类型，要求其他题题目必须包含“题”字
    # start_flag_6 = '题'
    result1 = '' 
    result2 = '' 
    result3 = '' 
    result4 = ''
    result5 = ''
    result_other = ''
    #默认index = 0 保证从0开始可以执行。
    #随着index不断增大。
    index = 0
    for i,line in enumerate(lines):
        if i < index:
            continue
        #############################存在问题，会把名词解释啥的写进去
        ########################存在問題，其他题目，这个地方可能会读到解析，如果解析中带有该题字样
        ####另外，这个不是跳跃读取，应该进入到某个block后，出来直接跳到该索引后。
        #如果是填空题，则返回判断题标签3
        if start_flag_1 in line:
            result1 = block.fill(lines[i+1:])[0] 
            #返回未匹配的索引，实现跳跃式循环
            index = block.fill(lines[i+1:])[1] 
            print('填空题:',result1)
            logger.info(result1)
        elif start_flag_2 in line: 
            #返回匹配值
            result2 = block.explan(lines[i+1:])[0] 
            #返回未匹配的索引，实现跳跃式循环
            index = block.explan(lines[i+1:])[1] 
            print('\n名词解释:',result2)
            logger.info("名词解释：")
            logger.info(result2)
        elif start_flag_3 in line: 
            result3 = block.simple(lines[i+1:])[0]
            index = block.simple(lines[i+1:])[1] 
            print('\n简答题:',result3)
            logger.info("简答题：")
            logger.info(result3)
        elif start_flag_4 in line:
            result4 = block.caculate(lines[i+1:])[0]  
            index = block.caculate(lines[i+1:])[1] 
            print('\n计算题:',result4)
            logger.info("计算题：")
            logger.info(result4)
        elif start_flag_5 in line:
            result5 = block.drawing(lines[i+1:])[0]
            index = block.drawing(lines[i+1:])[1] 
            print('\n画图题:',result5)
            logger.info("画图题：")
            logger.info(result5)
        # elif start_flag_6 in line:
        #     result_other = block.others(lines[i+1:])   
        else:
            continue

    return result1,result2,result3,result4,result5,result_other

if __name__ == '__main__':
    extract_block(r'主观题题目.docx')


