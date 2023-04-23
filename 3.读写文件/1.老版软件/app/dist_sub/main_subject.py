import blockextract_from_docx 
import contentextract
import blockextract_from_docx_ans 
import contentextract_ans


while(True):
    typeofinput = input('输入1，将录入主观题题目；输入2，将录入主观题答案：')
    if int(typeofinput) == 1:
        path = r'主观题题目.docx' # 读取文本的路径
        aimed_block_string = blockextract_from_docx.extract_block(path)  #提取目标块，此处为名词解释块，返回字符串
        contentextract.extract_content(aimed_block_string)  #提取内容，此处从名词解释块提取每段文字。
        break
    elif int(typeofinput) == 2:
        path = r'主观题答案.docx' # 读取文本的路径
        aimed_block_string = blockextract_from_docx_ans.extract_block(path)  #提取目标块，此处为名词解释块，返回字符串
        contentextract_ans.extract_content(aimed_block_string)  #提取内容，此处从名词解释块提取每段文字。
        break
    else:
        print('输入出错啦，请重新输入。')