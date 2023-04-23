import pandas as pd
from existed_data import ifexistedlist, ifexistedlist2
def explan(explan_of_block,excel_file_path = '主观题.xlsx'):
    """将列表读取，并写入到已经存在的列表zgt的E列。

    Args:
        ques_list (_type_): 列表，列表包含题目、选项、答案、解析。
        excel_file_path:相对路径
        type 暂定为题型名，不同题型都可以传入,传入参数为字符串。
    """
    try:
        df = pd.read_excel(excel_file_path, engine='openpyxl')
    except FileNotFoundError:
        print(f"文件路径 {excel_file_path} 不存在")
        return
    
    #读取存在的数据，确定写入的行索引
    existing_data = pd.read_excel(excel_file_path, sheet_name='Sheet1',usecols=[1])  #读取题型名字占用的索引
    last_row_index = existing_data.iloc[:, 0].last_valid_index()
    # 如果最后一个非空单元格所在的行索引为None，则说明该列所有单元格都为空，下一行索引为0；否则，下一行索引为最后一个非空单元格所在的行索引加1。
    if last_row_index is None:
        num_existing_rows = 0
    else:
        num_existing_rows = last_row_index + 1
    #将题目和答案块拆开
    ques_list = explan_of_block[0]
    ans_anal_list = explan_of_block[1]
    #将答案解析洗出来
    # 原始列表
    lst = ans_anal_list
    lst2 = []
    # 去掉开头和结尾的空格和换行符，并将关键词替换为空字符串
    answer = lst[0].strip().replace('答案 ', '')
    try:
        explanation = lst[1].strip().replace('解析 ', '') 
    except:
        explanation = ""
    # 将答案和解析写入到原始列表中
    lst2.append(answer)
    lst2.append(explanation)
    ans_anal_list = lst2   
    # print("答案和解析清理后：",ans_anal_list)

    #题目
    new_data_ques = pd.DataFrame(ifexistedlist(list=ques_list,index=0)) 
    
    #答案
    ans_data = ifexistedlist(list=ans_anal_list,index=0)
    #解析需要拼装成特别的形式  
    anal_data = ifexistedlist2(list=ans_anal_list,index=1) 

    #题干类型
    type_of_quest = pd.DataFrame(["名词解释"]) 
    #输入类型，即三种形式，文字、图片、音频。
    #根据label判断类型，然后传入进来。默认都为1-文字。
    #当某个地方为图片，这时候更新列表。
    type_tigan = pd.DataFrame([1])
    type_xiaowen1 = pd.DataFrame([1])
 
    #参考答案默认为文字类型，如果为图片则为2，如果是语音则为3
    type_xiaowen1_ans = pd.DataFrame([1])
    voice_time = pd.DataFrame([0])
    keywords = pd.DataFrame(["some keywords"])
    # 将新数据写入Excel文件的E列
    try:
        with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a',if_sheet_exists = 'overlay') as writer: #续写Mode改为'a'
        
            #写入题干输入形式，startcol  = 0
            type_tigan.to_excel(writer, sheet_name='Sheet1', startcol=0 ,startrow=num_existing_rows+1, index=False, header=False) 
            #写入题干种类，文字输入简答题，图片，语音输入地址。startcol  = 1
            type_of_quest.to_excel(writer, sheet_name='Sheet1', startcol=1 ,startrow=num_existing_rows+1, index=False, header=False) 
            # #写入题目音频秒数，startcol = 2
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=2 ,startrow=num_existing_rows+1, index=False, header=False) 
            #小问类型，startcol = 3  
            type_xiaowen1.to_excel(writer, sheet_name='Sheet1', startcol=3 ,startrow=num_existing_rows+1, index=False, header=False) 
            #小问标题为语音，写入值为语音秒数，startcol = 5
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=7 ,startrow=num_existing_rows+1, index=False, header=False) 
            #关键字,startcol = 6
            # keywords.to_excel(writer, sheet_name='Sheet1', startcol=6 ,startrow=num_existing_rows+1, index=False, header=False)
            # #写入语音秒数
            # voice_index = [2,5,9,12,16,19,23,26,30]
            # for index_v in voice_index:
            #     voice_time.to_excel(writer, sheet_name='Sheet1', startcol=index_v,startrow=num_existing_rows+1, index=False, header=False) 
            #写入小问标题，startcol = 4
            new_data_ques.to_excel(writer, sheet_name='Sheet1', startcol=4,startrow=num_existing_rows+1, index=False, header=False)  #head = flase 就是不用写表格头
            """
            因为答案写入非常复杂,需要对应不同列写入值1和2.
            因此将答案和解析通过函数single_ans写入。
            函数接受，数据，写入器，写入行索引。
            """
            #主观题，一般没有解析，因此就看看有没有答案。
            try:
                ans = pd.DataFrame(ans_data)
                ans.to_excel(writer, sheet_name='Sheet1', startcol=8,startrow=num_existing_rows+1, index=False, header=False) 
            except:
                print("没有读到答案")
            #解析应该没有，假如有，就写入到关键字。
            try:
                ans = pd.DataFrame(anal_data)
                ans.to_excel(writer, sheet_name='Sheet1', startcol=6,startrow=num_existing_rows+1, index=False, header=False) 
            except Exception as e:
                print("没有读到解析")


    except Exception as e:
        print(f"写入Excel文件时出错: {e}")
        return
    print("数据已成功写入Excel文件的目标列!")


def fill(explan_of_block,excel_file_path = '主观题.xlsx'):
    """将列表读取，并写入到已经存在的列表zgt的E列。

    Args:
        ques_list (_type_): 列表，列表包含题目、选项、答案、解析。
        excel_file_path:相对路径
        type 暂定为题型名，不同题型都可以传入,传入参数为字符串。
    """
    try:
        df = pd.read_excel(excel_file_path, engine='openpyxl')
    except FileNotFoundError:
        print(f"文件路径 {excel_file_path} 不存在")
        return
    
    #读取存在的数据，确定写入的行索引
    existing_data = pd.read_excel(excel_file_path, sheet_name='Sheet1',usecols=[1])  #读取题型名字占用的索引
    last_row_index = existing_data.iloc[:, 0].last_valid_index()
    # 如果最后一个非空单元格所在的行索引为None，则说明该列所有单元格都为空，下一行索引为0；否则，下一行索引为最后一个非空单元格所在的行索引加1。
    if last_row_index is None:
        num_existing_rows = 0
    else:
        num_existing_rows = last_row_index + 1
    #将题目和答案块拆开
    ques_list = explan_of_block[0]
    ans_anal_list = explan_of_block[1]
    #将答案解析洗出来
    # 原始列表
    lst = ans_anal_list
    lst2 = []
    # 去掉开头和结尾的空格和换行符，并将关键词替换为空字符串
    answer = lst[0].strip().replace('答案 ', '')
    try:
        explanation = lst[1].strip().replace('解析 ', '') 
    except:
        explanation = ""
    # 将答案和解析写入到原始列表中
    lst2.append(answer)
    lst2.append(explanation)
    ans_anal_list = lst2     
    # print("答案和解析清理后：",ans_anal_list)

    #题目
    new_data_ques = pd.DataFrame(ifexistedlist(list=ques_list,index=0)) 
    
    #答案
    ans_data = ifexistedlist(list=ans_anal_list,index=0)
    #解析需要拼装成特别的形式  
    anal_data = ifexistedlist2(list=ans_anal_list,index=1) 

    #题干类型
    type_of_quest = pd.DataFrame(["填空题"]) 
    #输入类型，即三种形式，文字、图片、音频。
    #根据label判断类型，然后传入进来。默认都为1-文字。
    #当某个地方为图片，这时候更新列表。
    type_tigan = pd.DataFrame([1])
    type_xiaowen1 = pd.DataFrame([1])
 
    #参考答案默认为文字类型，如果为图片则为2，如果是语音则为3
    type_xiaowen1_ans = pd.DataFrame([1])
    voice_time = pd.DataFrame([0])
    keywords = pd.DataFrame(["some keywords"])
    # 将新数据写入Excel文件的E列
    try:
        with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a',if_sheet_exists = 'overlay') as writer: #续写Mode改为'a'
        
            #写入题干输入形式，startcol  = 0
            type_tigan.to_excel(writer, sheet_name='Sheet1', startcol=0 ,startrow=num_existing_rows+1, index=False, header=False) 
            #写入题干种类，文字输入简答题，图片，语音输入地址。startcol  = 1
            type_of_quest.to_excel(writer, sheet_name='Sheet1', startcol=1 ,startrow=num_existing_rows+1, index=False, header=False) 
            # #写入题目音频秒数，startcol = 2
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=2 ,startrow=num_existing_rows+1, index=False, header=False) 
            #小问类型，startcol = 3  
            type_xiaowen1.to_excel(writer, sheet_name='Sheet1', startcol=3 ,startrow=num_existing_rows+1, index=False, header=False) 
            #小问标题为语音，写入值为语音秒数，startcol = 5
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=7 ,startrow=num_existing_rows+1, index=False, header=False) 
            #关键字,startcol = 6
            # keywords.to_excel(writer, sheet_name='Sheet1', startcol=6 ,startrow=num_existing_rows+1, index=False, header=False)
            # #写入语音秒数
            # voice_index = [2,5,9,12,16,19,23,26,30]
            # for index_v in voice_index:
            #     voice_time.to_excel(writer, sheet_name='Sheet1', startcol=index_v,startrow=num_existing_rows+1, index=False, header=False) 
            #写入小问标题，startcol = 4
            new_data_ques.to_excel(writer, sheet_name='Sheet1', startcol=4,startrow=num_existing_rows+1, index=False, header=False)  #head = flase 就是不用写表格头
            """
            因为答案写入非常复杂,需要对应不同列写入值1和2.
            因此将答案和解析通过函数single_ans写入。
            函数接受，数据，写入器，写入行索引。
            """
            #主观题，一般没有解析，因此就看看有没有答案。
            try:
                ans = pd.DataFrame(ans_data)
                ans.to_excel(writer, sheet_name='Sheet1', startcol=8,startrow=num_existing_rows+1, index=False, header=False) 
            except:
                print("没有读到答案",e)
            #解析应该没有，假如有，就写入到关键字。
            try:
                ans = pd.DataFrame(anal_data)
                ans.to_excel(writer, sheet_name='Sheet1', startcol=6,startrow=num_existing_rows+1, index=False, header=False) 
            except:
                print("没有读到解析",e)


    except Exception as e:
        print(f"写入Excel文件时出错: {e}")
        return
    print("数据已成功写入Excel文件的目标列!")


def simple(explan_of_block,excel_file_path = '主观题.xlsx'):
    """将列表读取，并写入到已经存在的列表zgt的E列。

    Args:
        ques_list (_type_): 列表，列表包含题目、选项、答案、解析。
        excel_file_path:相对路径
        type 暂定为题型名，不同题型都可以传入,传入参数为字符串。
    """
    try:
        df = pd.read_excel(excel_file_path, engine='openpyxl')
    except FileNotFoundError:
        print(f"文件路径 {excel_file_path} 不存在")
        return
    
    #读取存在的数据，确定写入的行索引
    existing_data = pd.read_excel(excel_file_path, sheet_name='Sheet1',usecols=[1])  #读取题型名字占用的索引
    last_row_index = existing_data.iloc[:, 0].last_valid_index()
    # 如果最后一个非空单元格所在的行索引为None，则说明该列所有单元格都为空，下一行索引为0；否则，下一行索引为最后一个非空单元格所在的行索引加1。
    if last_row_index is None:
        num_existing_rows = 0
    else:
        num_existing_rows = last_row_index + 1
    #将题目和答案块拆开
    ques_list = explan_of_block[0]
    ans_anal_list = explan_of_block[1]
    #将答案解析洗出来
    # 原始列表
    lst = ans_anal_list
    lst2=[]
    # 去掉开头和结尾的空格和换行符，并将关键词替换为空字符串
    answer = lst[0].strip().replace('答案 ', '')
    try:
        explanation = lst[1].strip().replace('解析 ', '') 
    except:
        explanation = ""
    # 将答案和解析写入到原始列表中
    lst2.append(answer)
    lst2.append(explanation)
    ans_anal_list = lst2    
    # print("答案和解析清理后：",ans_anal_list)

    #题目
    new_data_ques = pd.DataFrame(ifexistedlist(list=ques_list,index=0)) 
    
    #答案
    ans_data = ifexistedlist(list=ans_anal_list,index=0)
    #解析需要拼装成特别的形式  
    anal_data = ifexistedlist2(list=ans_anal_list,index=1) 

    #题干类型
    type_of_quest = pd.DataFrame(["简答题"]) 
    #输入类型，即三种形式，文字、图片、音频。
    #根据label判断类型，然后传入进来。默认都为1-文字。
    #当某个地方为图片，这时候更新列表。
    type_tigan = pd.DataFrame([1])
    type_xiaowen1 = pd.DataFrame([1])
 
    #参考答案默认为文字类型，如果为图片则为2，如果是语音则为3
    type_xiaowen1_ans = pd.DataFrame([1])
    voice_time = pd.DataFrame([0])
    keywords = pd.DataFrame(["some keywords"])
    # 将新数据写入Excel文件的E列
    try:
        with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a',if_sheet_exists = 'overlay') as writer: #续写Mode改为'a'
        
            #写入题干输入形式，startcol  = 0
            type_tigan.to_excel(writer, sheet_name='Sheet1', startcol=0 ,startrow=num_existing_rows+1, index=False, header=False) 
            #写入题干种类，文字输入简答题，图片，语音输入地址。startcol  = 1
            type_of_quest.to_excel(writer, sheet_name='Sheet1', startcol=1 ,startrow=num_existing_rows+1, index=False, header=False) 
            # #写入题目音频秒数，startcol = 2
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=2 ,startrow=num_existing_rows+1, index=False, header=False) 
            #小问类型，startcol = 3  
            type_xiaowen1.to_excel(writer, sheet_name='Sheet1', startcol=3 ,startrow=num_existing_rows+1, index=False, header=False) 
            #小问标题为语音，写入值为语音秒数，startcol = 5
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=7 ,startrow=num_existing_rows+1, index=False, header=False) 
            #关键字,startcol = 6
            # keywords.to_excel(writer, sheet_name='Sheet1', startcol=6 ,startrow=num_existing_rows+1, index=False, header=False)
            # #写入语音秒数
            # voice_index = [2,5,9,12,16,19,23,26,30]
            # for index_v in voice_index:
            #     voice_time.to_excel(writer, sheet_name='Sheet1', startcol=index_v,startrow=num_existing_rows+1, index=False, header=False) 
            #写入小问标题，startcol = 4
            new_data_ques.to_excel(writer, sheet_name='Sheet1', startcol=4,startrow=num_existing_rows+1, index=False, header=False)  #head = flase 就是不用写表格头
            """
            因为答案写入非常复杂,需要对应不同列写入值1和2.
            因此将答案和解析通过函数single_ans写入。
            函数接受，数据，写入器，写入行索引。
            """
            #主观题，一般没有解析，因此就看看有没有答案。
            try:
                ans = pd.DataFrame(ans_data)
                ans.to_excel(writer, sheet_name='Sheet1', startcol=8,startrow=num_existing_rows+1, index=False, header=False) 
            except:
                print("没有读到答案",e)
            #解析应该没有，假如有，就写入到关键字。
            try:
                ans = pd.DataFrame(anal_data)
                ans.to_excel(writer, sheet_name='Sheet1', startcol=6,startrow=num_existing_rows+1, index=False, header=False) 
            except:
                print("没有读到解析",e)
    except Exception as e:
        print(f"写入Excel文件时出错: {e}")
        return
    print("数据已成功写入Excel文件的目标列!")

def drawing(explan_of_block,excel_file_path = '主观题.xlsx'):
    """将列表读取，并写入到已经存在的列表zgt的E列。

    Args:
        ques_list (_type_): 列表，列表包含题目、选项、答案、解析。
        excel_file_path:相对路径
        type 暂定为题型名，不同题型都可以传入,传入参数为字符串。
    """
    try:
        df = pd.read_excel(excel_file_path, engine='openpyxl')
    except FileNotFoundError:
        print(f"文件路径 {excel_file_path} 不存在")
        return
    
    #读取存在的数据，确定写入的行索引
    existing_data = pd.read_excel(excel_file_path, sheet_name='Sheet1',usecols=[1])  #读取题型名字占用的索引
    last_row_index = existing_data.iloc[:, 0].last_valid_index()
    # 如果最后一个非空单元格所在的行索引为None，则说明该列所有单元格都为空，下一行索引为0；否则，下一行索引为最后一个非空单元格所在的行索引加1。
    if last_row_index is None:
        num_existing_rows = 0
    else:
        num_existing_rows = last_row_index + 1
    #将题目和答案块拆开
    ques_list = explan_of_block[0]
    ans_anal_list = explan_of_block[1]
    #将答案解析洗出来
    # 原始列表
    lst = ans_anal_list
    lst2 = []
    # 去掉开头和结尾的空格和换行符，并将关键词替换为空字符串
    answer = lst[0].strip().replace('答案 ', '')
    try:
        explanation = lst[1].strip().replace('解析 ', '') 
    except:
        explanation = ""
    # 将答案和解析写入到原始列表中
    lst2.append(answer)
    lst2.append(explanation)
    ans_anal_list = lst2   
    # print("答案和解析清理后：",ans_anal_list)

    #题目
    new_data_ques = pd.DataFrame(ifexistedlist(list=ques_list,index=0)) 
    
    #答案
    ans_data = ifexistedlist(list=ans_anal_list,index=0)
    #解析需要拼装成特别的形式  
    anal_data = ifexistedlist2(list=ans_anal_list,index=1) 

    #题干类型
    type_of_quest = pd.DataFrame(["画图题"]) 
    #输入类型，即三种形式，文字、图片、音频。
    #根据label判断类型，然后传入进来。默认都为1-文字。
    #当某个地方为图片，这时候更新列表。
    type_tigan = pd.DataFrame([1])
    type_xiaowen1 = pd.DataFrame([1])
 
    #参考答案默认为文字类型，如果为图片则为2，如果是语音则为3
    type_xiaowen1_ans = pd.DataFrame([1])
    voice_time = pd.DataFrame([0])
    keywords = pd.DataFrame(["some keywords"])
    # 将新数据写入Excel文件的E列
    try:
        with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a',if_sheet_exists = 'overlay') as writer: #续写Mode改为'a'
        
            #写入题干输入形式，startcol  = 0
            type_tigan.to_excel(writer, sheet_name='Sheet1', startcol=0 ,startrow=num_existing_rows+1, index=False, header=False) 
            #写入题干种类，文字输入简答题，图片，语音输入地址。startcol  = 1
            type_of_quest.to_excel(writer, sheet_name='Sheet1', startcol=1 ,startrow=num_existing_rows+1, index=False, header=False) 
            # #写入题目音频秒数，startcol = 2
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=2 ,startrow=num_existing_rows+1, index=False, header=False) 
            #小问类型，startcol = 3  
            type_xiaowen1.to_excel(writer, sheet_name='Sheet1', startcol=3 ,startrow=num_existing_rows+1, index=False, header=False) 
            #小问标题为语音，写入值为语音秒数，startcol = 5
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=7 ,startrow=num_existing_rows+1, index=False, header=False) 
            #关键字,startcol = 6
            # keywords.to_excel(writer, sheet_name='Sheet1', startcol=6 ,startrow=num_existing_rows+1, index=False, header=False)
            # #写入语音秒数
            # voice_index = [2,5,9,12,16,19,23,26,30]
            # for index_v in voice_index:
            #     voice_time.to_excel(writer, sheet_name='Sheet1', startcol=index_v,startrow=num_existing_rows+1, index=False, header=False) 
            #写入小问标题，startcol = 4
            new_data_ques.to_excel(writer, sheet_name='Sheet1', startcol=4,startrow=num_existing_rows+1, index=False, header=False)  #head = flase 就是不用写表格头
            """
            因为答案写入非常复杂,需要对应不同列写入值1和2.
            因此将答案和解析通过函数single_ans写入。
            函数接受，数据，写入器，写入行索引。
            """
            #主观题，一般没有解析，因此就看看有没有答案。
            try:
                ans = pd.DataFrame(ans_data)
                ans.to_excel(writer, sheet_name='Sheet1', startcol=8,startrow=num_existing_rows+1, index=False, header=False) 
            except:
                print("没有读到答案",e)
            #解析应该没有，假如有，就写入到关键字。
            try:
                ans = pd.DataFrame(anal_data)
                ans.to_excel(writer, sheet_name='Sheet1', startcol=6,startrow=num_existing_rows+1, index=False, header=False) 
            except:
                print("没有读到解析",e)
    except Exception as e:
        print(f"写入Excel文件时出错: {e}")
        return
    print("数据已成功写入Excel文件的目标列!")


def caculate(explan_of_block,excel_file_path = '主观题.xlsx'):
    """将列表读取，并写入到已经存在的列表zgt的E列。

    Args:
        ques_list (_type_): 列表，列表包含题目、选项、答案、解析。
        excel_file_path:相对路径
        type 暂定为题型名，不同题型都可以传入,传入参数为字符串。
    """
    try:
        df = pd.read_excel(excel_file_path, engine='openpyxl')
    except FileNotFoundError:
        print(f"文件路径 {excel_file_path} 不存在")
        return
    
    #读取存在的数据，确定写入的行索引
    existing_data = pd.read_excel(excel_file_path, sheet_name='Sheet1',usecols=[1])  #读取题型名字占用的索引
    last_row_index = existing_data.iloc[:, 0].last_valid_index()
    # 如果最后一个非空单元格所在的行索引为None，则说明该列所有单元格都为空，下一行索引为0；否则，下一行索引为最后一个非空单元格所在的行索引加1。
    if last_row_index is None:
        num_existing_rows = 0
    else:
        num_existing_rows = last_row_index + 1
    #将题目和答案块拆开
    ques_list = explan_of_block[0]
    ans_anal_list = explan_of_block[1]
    #将答案解析洗出来
    # 原始列表
    lst = ans_anal_list
    lst2 = []
    # 去掉开头和结尾的空格和换行符，并将关键词替换为空字符串
    answer = lst[0].strip().replace('答案 ', '')
    try:
        explanation = lst[1].strip().replace('解析 ', '') 
    except:
        explanation = ""
    # 将答案和解析写入到原始列表中
    lst2.append(answer)
    lst2.append(explanation)
    ans_anal_list = lst2   
    # print("答案和解析清理后：",ans_anal_list)

    #题目
    new_data_ques = pd.DataFrame(ifexistedlist(list=ques_list,index=0)) 
    
    #答案
    ans_data = ifexistedlist(list=ans_anal_list,index=0)
    #解析需要拼装成特别的形式  
    anal_data = ifexistedlist2(list=ans_anal_list,index=1) 

    #题干类型
    type_of_quest = pd.DataFrame(["计算题"]) 
    #输入类型，即三种形式，文字、图片、音频。
    #根据label判断类型，然后传入进来。默认都为1-文字。
    #当某个地方为图片，这时候更新列表。
    type_tigan = pd.DataFrame([1])
    type_xiaowen1 = pd.DataFrame([1])
 
    #参考答案默认为文字类型，如果为图片则为2，如果是语音则为3
    type_xiaowen1_ans = pd.DataFrame([1])
    voice_time = pd.DataFrame([0])
    keywords = pd.DataFrame(["some keywords"])
    # 将新数据写入Excel文件的E列
    try:
        with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a',if_sheet_exists = 'overlay') as writer: #续写Mode改为'a'
        
            #写入题干输入形式，startcol  = 0
            type_tigan.to_excel(writer, sheet_name='Sheet1', startcol=0 ,startrow=num_existing_rows+1, index=False, header=False) 
            #写入题干种类，文字输入简答题，图片，语音输入地址。startcol  = 1
            type_of_quest.to_excel(writer, sheet_name='Sheet1', startcol=1 ,startrow=num_existing_rows+1, index=False, header=False) 
            # #写入题目音频秒数，startcol = 2
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=2 ,startrow=num_existing_rows+1, index=False, header=False) 
            #小问类型，startcol = 3  
            type_xiaowen1.to_excel(writer, sheet_name='Sheet1', startcol=3 ,startrow=num_existing_rows+1, index=False, header=False) 
            #小问标题为语音，写入值为语音秒数，startcol = 5
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=7 ,startrow=num_existing_rows+1, index=False, header=False) 
            #关键字,startcol = 6
            # keywords.to_excel(writer, sheet_name='Sheet1', startcol=6 ,startrow=num_existing_rows+1, index=False, header=False)
            # #写入语音秒数
            # voice_index = [2,5,9,12,16,19,23,26,30]
            # for index_v in voice_index:
            #     voice_time.to_excel(writer, sheet_name='Sheet1', startcol=index_v,startrow=num_existing_rows+1, index=False, header=False) 
            #写入小问标题，startcol = 4
            new_data_ques.to_excel(writer, sheet_name='Sheet1', startcol=4,startrow=num_existing_rows+1, index=False, header=False)  #head = flase 就是不用写表格头
            """
            因为答案写入非常复杂,需要对应不同列写入值1和2.
            因此将答案和解析通过函数single_ans写入。
            函数接受，数据，写入器，写入行索引。
            """
            #主观题，一般没有解析，因此就看看有没有答案。
            try:
                ans = pd.DataFrame(ans_data)
                ans.to_excel(writer, sheet_name='Sheet1', startcol=8,startrow=num_existing_rows+1, index=False, header=False) 
            except:
                print("没有读到答案",e)
            #解析应该没有，假如有，就写入到关键字。
            try:
                ans = pd.DataFrame(anal_data)
                ans.to_excel(writer, sheet_name='Sheet1', startcol=6,startrow=num_existing_rows+1, index=False, header=False) 
            except:
                print("没有读到解析",e)
    except Exception as e:
        print(f"写入Excel文件时出错: {e}")
        return
    print("数据已成功写入Excel文件的目标列!")


def others(explan_of_block,excel_file_path = '主观题.xlsx'):
    """将列表读取，并写入到已经存在的列表zgt的E列。

    Args:
        ques_list (_type_): 列表，列表包含题目、选项、答案、解析。
        excel_file_path:相对路径
        type 暂定为题型名，不同题型都可以传入,传入参数为字符串。
    """
    try:
        df = pd.read_excel(excel_file_path, engine='openpyxl')
    except FileNotFoundError:
        print(f"文件路径 {excel_file_path} 不存在")
        return
    
    #读取存在的数据，确定写入的行索引
    existing_data = pd.read_excel(excel_file_path, sheet_name='Sheet1',usecols=[1])  #读取题型名字占用的索引
    last_row_index = existing_data.iloc[:, 0].last_valid_index()
    # 如果最后一个非空单元格所在的行索引为None，则说明该列所有单元格都为空，下一行索引为0；否则，下一行索引为最后一个非空单元格所在的行索引加1。
    if last_row_index is None:
        num_existing_rows = 0
    else:
        num_existing_rows = last_row_index + 1
    #将题目和答案块拆开
    ques_list = explan_of_block[0]
    ans_anal_list = explan_of_block[1]
    #将答案解析洗出来
    # 原始列表
    lst = ans_anal_list
    lst2 = []
    # 去掉开头和结尾的空格和换行符，并将关键词替换为空字符串
    answer = lst[0].strip().replace('答案 ', '')
    try:
        explanation = lst[1].strip().replace('解析 ', '') 
    except:
        explanation = ""
    # 将答案和解析写入到原始列表中
    lst2.append(answer)
    lst2.append(explanation)
    ans_anal_list = lst2   
    # print("答案和解析清理后：",ans_anal_list)

    #题目
    new_data_ques = pd.DataFrame(ifexistedlist(list=ques_list,index=0)) 
    
    #答案
    ans_data = ifexistedlist(list=ans_anal_list,index=0)
    #解析需要拼装成特别的形式  
    anal_data = ifexistedlist2(list=ans_anal_list,index=1) 

    #题干类型
    type_of_quest = pd.DataFrame(["其他题型"]) 
    #输入类型，即三种形式，文字、图片、音频。
    #根据label判断类型，然后传入进来。默认都为1-文字。
    #当某个地方为图片，这时候更新列表。
    type_tigan = pd.DataFrame([1])
    type_xiaowen1 = pd.DataFrame([1])
 
    #参考答案默认为文字类型，如果为图片则为2，如果是语音则为3
    type_xiaowen1_ans = pd.DataFrame([1])
    voice_time = pd.DataFrame([0])
    keywords = pd.DataFrame(["some keywords"])
    # 将新数据写入Excel文件的E列
    try:
        with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a',if_sheet_exists = 'overlay') as writer: #续写Mode改为'a'
        
            #写入题干输入形式，startcol  = 0
            type_tigan.to_excel(writer, sheet_name='Sheet1', startcol=0 ,startrow=num_existing_rows+1, index=False, header=False) 
            #写入题干种类，文字输入简答题，图片，语音输入地址。startcol  = 1
            type_of_quest.to_excel(writer, sheet_name='Sheet1', startcol=1 ,startrow=num_existing_rows+1, index=False, header=False) 
            # #写入题目音频秒数，startcol = 2
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=2 ,startrow=num_existing_rows+1, index=False, header=False) 
            #小问类型，startcol = 3  
            type_xiaowen1.to_excel(writer, sheet_name='Sheet1', startcol=3 ,startrow=num_existing_rows+1, index=False, header=False) 
            #小问标题为语音，写入值为语音秒数，startcol = 5
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=7 ,startrow=num_existing_rows+1, index=False, header=False) 
            #关键字,startcol = 6
            # keywords.to_excel(writer, sheet_name='Sheet1', startcol=6 ,startrow=num_existing_rows+1, index=False, header=False)
            # #写入语音秒数
            # voice_index = [2,5,9,12,16,19,23,26,30]
            # for index_v in voice_index:
            #     voice_time.to_excel(writer, sheet_name='Sheet1', startcol=index_v,startrow=num_existing_rows+1, index=False, header=False) 
            #写入小问标题，startcol = 4
            new_data_ques.to_excel(writer, sheet_name='Sheet1', startcol=4,startrow=num_existing_rows+1, index=False, header=False)  #head = flase 就是不用写表格头
            """
            因为答案写入非常复杂,需要对应不同列写入值1和2.
            因此将答案和解析通过函数single_ans写入。
            函数接受，数据，写入器，写入行索引。
            """
            #主观题，一般没有解析，因此就看看有没有答案。
            try:
                ans = pd.DataFrame(ans_data)
                ans.to_excel(writer, sheet_name='Sheet1', startcol=8,startrow=num_existing_rows+1, index=False, header=False) 
            except:
                print("没有读到答案",e)
            #解析应该没有，假如有，就写入到关键字。
            try:
                ans = pd.DataFrame(anal_data)
                ans.to_excel(writer, sheet_name='Sheet1', startcol=6,startrow=num_existing_rows+1, index=False, header=False) 
            except:
                print("没有读到解析",e)
    except Exception as e:
        print(f"写入Excel文件时出错: {e}")
        return
    print("数据已成功写入Excel文件的目标列!")