import pandas as pd
from  existed_data import ifexistedlist, ifexistedlist2
import ans_all
def single_write(single_choice_block,excel_file_path = '客观题.xlsx'):
    """将列表读取,并写入到已经存在的列表zgt的E列。

    Args:
        ques_list (_type_): 列表，列表包含题目、选项、答案、解析。
        excel_file_path:相对路径
    """
    try:
        df = pd.read_excel(excel_file_path, engine='openpyxl')
    except FileNotFoundError:
        print(f"文件路径 {excel_file_path} 不存在")
        return
    
    #读取存在的数据，确定写入的行索引
    existing_data = pd.read_excel(excel_file_path, sheet_name='Sheet1',usecols=[2])  #题目、选项A,B,C,D
    last_row_index = existing_data.iloc[:, 0].last_valid_index()
    # 如果最后一个非空单元格所在的行索引为None，则说明该列所有单元格都为空，下一行索引为0；否则，下一行索引为最后一个非空单元格所在的行索引加1。
    if last_row_index is None:
        num_existing_rows = 0
    else:
        num_existing_rows = last_row_index + 1
    #将题目和答案块拆开
    ques_choice_list = single_choice_block[0]
    length_choice = len(ques_choice_list)-1  #获得单选题的选项长度
    ans_anal_list = single_choice_block[1]
    #将答案解析洗出来
    # 原始列表
    lst = ans_anal_list
    # 去掉开头和结尾的空格和换行符，并将关键词替换为空字符串
    answer = lst[0].strip().replace('答案 ', '')
    try:
        explanation = lst[1].strip().replace('解析 ', '') 
    except:
        explanation = ""
    # 将答案和解析写入到原始列表中
    lst[0] = answer
    lst.append(explanation)
    ans_anal_list = lst    
    # print("答案和解析清理后：",ans_anal_list)

    #题目
    new_data_ques = pd.DataFrame(ifexistedlist(list=ques_choice_list,index=0)) 
    #选项
    new_data_A = pd.DataFrame(ifexistedlist(list=ques_choice_list,index=1)) 
    new_data_B = pd.DataFrame(ifexistedlist(list=ques_choice_list,index=2)) 
    new_data_C = pd.DataFrame(ifexistedlist(list=ques_choice_list,index=3)) 
    new_data_D = pd.DataFrame(ifexistedlist(list=ques_choice_list,index=4)) 

    
    new_data_E = pd.DataFrame(ifexistedlist(list=ques_choice_list,index=5)) #有些配伍题有E
    new_data_F = pd.DataFrame(ifexistedlist(list=ques_choice_list,index=6))  #有些配伍题有E

    #答案
    ans_data = ifexistedlist(list=ans_anal_list,index=0)
    #解析需要拼装成特别的形式  
    anal_data = ifexistedlist2(list=ans_anal_list,index=1) 


    #题目类型
    type_of_choice = pd.DataFrame([1]) 
    #输入类型
    type_of_input = pd.DataFrame([1]) 
    voice_time = pd.DataFrame([0])
    
    # 将新数据写入Excel文件的E列
    try:
        with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a',if_sheet_exists = 'overlay') as writer: #续写Mode改为'a'
            #写入题目类型，startcol  = 0
            type_of_choice.to_excel(writer, sheet_name='Sheet1', startcol=0 ,startrow=num_existing_rows+1, index=False, header=False) 
            #写入题干输入类型，startcol  = 1
            type_of_input.to_excel(writer, sheet_name='Sheet1', startcol=1 ,startrow=num_existing_rows+1, index=False, header=False) 
            # #写入题目音频秒数，startcol = 3
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=3 ,startrow=num_existing_rows+1, index=False, header=False) 
            #写入选项输入类型，startcol = 4
            type_of_input.to_excel(writer, sheet_name='Sheet1', startcol=4 ,startrow=num_existing_rows+1, index=False, header=False) 
            #选项1为语音是语音秒数，startcol = 7
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=7 ,startrow=num_existing_rows+1, index=False, header=False) 
            # #选项2为语音是语音秒数，startcol = 10
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=10 ,startrow=num_existing_rows+1, index=False, header=False) 
            # #选项3为语音是语音秒数，startcol = 13
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=13 ,startrow=num_existing_rows+1, index=False, header=False) 
            # #选项4为语音是语音秒数，startcol = 16
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=16 ,startrow=num_existing_rows+1, index=False, header=False) 
            # #选项5为语音是语音秒数，startcol = 19
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=19 ,startrow=num_existing_rows+1, index=False, header=False) 
            # #选项6为语音是语音秒数，startcol = 22
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=22 ,startrow=num_existing_rows+1, index=False, header=False)  

            #题目写入
            new_data_ques.to_excel(writer, sheet_name='Sheet1', startcol=2,startrow=num_existing_rows+1, index=False, header=False)  #head = flase 就是不用写表格头
            #选项abcd写入
            new_data_A.to_excel(writer, sheet_name='Sheet1', startcol=6  ,startrow=num_existing_rows+1, index=False, header=False)  
            new_data_B.to_excel(writer, sheet_name='Sheet1', startcol=9  ,startrow=num_existing_rows+1, index=False, header=False)  
            new_data_C.to_excel(writer, sheet_name='Sheet1', startcol=12  ,startrow=num_existing_rows+1, index=False, header=False)  
            new_data_D.to_excel(writer, sheet_name='Sheet1', startcol=15  ,startrow=num_existing_rows+1, index=False, header=False)  
            #选项ef写入
            new_data_E.to_excel(writer, sheet_name='Sheet1', startcol=18  ,startrow=num_existing_rows+1, index=False, header=False)  #单选配伍题
            new_data_F.to_excel(writer, sheet_name='Sheet1', startcol=21  ,startrow=num_existing_rows+1, index=False, header=False)  #单选配伍题
            """
            因为答案写入非常复杂,需要对应不同列写入值1和2.
            因此将答案和解析通过函数single_ans写入。
            函数接受，数据，写入器，写入行索引。
            """
            #答案写入,需要将答案剥离
            ans_all.xlsx_single_ans(ans_data,writer,num_existing_rows+1,length_choice)
            #解析写入，需要把解析剥离
            ans_all.xlsx_single_ans(anal_data,writer,num_existing_rows+1,anal = True)

    except Exception as e:
        print(f"写入Excel文件时出错: {e}")
        return
    print("数据已成功写入Excel文件的目标列!")


def multi_write(multi_choice_block,excel_file_path = '客观题.xlsx'):
    """将列表读取，并写入到已经存在的列表zgt的E列。

    Args:
        ques_list (_type_): 列表，列表包含题目、选项、答案、解析。
        excel_file_path:相对路径
    """
    try:
        df = pd.read_excel(excel_file_path, engine='openpyxl')
    except FileNotFoundError:
        print(f"文件路径 {excel_file_path} 不存在")
        return
    
    #读取存在的数据，确定写入的行索引
    existing_data = pd.read_excel(excel_file_path, sheet_name='Sheet1',usecols=[2])  #仅仅读取题目
    last_row_index = existing_data.iloc[:, 0].last_valid_index()
    # 如果最后一个非空单元格所在的行索引为None，则说明该列所有单元格都为空，下一行索引为0；否则，下一行索引为最后一个非空单元格所在的行索引加1。
    if last_row_index is None:
        num_existing_rows = 0
    else:
        num_existing_rows = last_row_index + 1
    #将题目和答案块拆开
    ques_choice_list = multi_choice_block[0]
    length_choice = len(ques_choice_list)-1  #获得多选题的选项长度
    ans_anal_list = multi_choice_block[1]
    #将答案解析洗出来
    # 原始列表
    lst = ans_anal_list
    # 去掉开头和结尾的空格和换行符，并将关键词替换为空字符串
    answer = lst[0].strip().replace('答案 ', '')
    try:
        explanation = lst[1].strip().replace('解析 ', '') 
    except:
        explanation = ""
    # 将答案和解析写入到原始列表中
    lst[0] = answer
    lst.append(explanation)
    ans_anal_list = lst    
    # print("答案和解析清理后：",ans_anal_list)

    #题目
    new_data_ques = pd.DataFrame(ifexistedlist(list=ques_choice_list,index=0)) 
    #选项
    new_data_A = pd.DataFrame(ifexistedlist(list=ques_choice_list,index=1)) 
    new_data_B = pd.DataFrame(ifexistedlist(list=ques_choice_list,index=2)) 
    new_data_C = pd.DataFrame(ifexistedlist(list=ques_choice_list,index=3)) 
    new_data_D = pd.DataFrame(ifexistedlist(list=ques_choice_list,index=4)) 

    
    new_data_E = pd.DataFrame(ifexistedlist(list=ques_choice_list,index=5)) #有些配伍题有E
    new_data_F = pd.DataFrame(ifexistedlist(list=ques_choice_list,index=6))  #有些配伍题有E

    #答案
    ans_data = ifexistedlist(list=ans_anal_list,index=0)
    #解析需要拼装成特别的形式  
    anal_data = ifexistedlist2(list=ans_anal_list,index=1) 


    #题目类型
    type_of_choice = pd.DataFrame([2]) 
    #输入类型
    type_of_input = pd.DataFrame([1]) 
    voice_time = pd.DataFrame([0])
    

    try:
        with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a',if_sheet_exists = 'overlay') as writer: #续写Mode改为'a'
            #写入题目类型，startcol  = 0
            type_of_choice.to_excel(writer, sheet_name='Sheet1', startcol=0 ,startrow=num_existing_rows+1, index=False, header=False) 
            #写入题干类型输入类型，startcol  = 1
            type_of_input.to_excel(writer, sheet_name='Sheet1', startcol=1 ,startrow=num_existing_rows+1, index=False, header=False) 
            # #写入题目音频秒数，startcol = 3
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=3 ,startrow=num_existing_rows+1, index=False, header=False) 
            #写入选项类型输入类型，startcol = 4
            type_of_input.to_excel(writer, sheet_name='Sheet1', startcol=4 ,startrow=num_existing_rows+1, index=False, header=False) 
            #选项1为语音是语音秒数，startcol = 7
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=7 ,startrow=num_existing_rows+1, index=False, header=False) 
            # #选项2为语音是语音秒数，startcol = 10
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=10 ,startrow=num_existing_rows+1, index=False, header=False) 
            # #选项3为语音是语音秒数，startcol = 13
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=13 ,startrow=num_existing_rows+1, index=False, header=False) 
            # #选项4为语音是语音秒数，startcol = 16
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=16 ,startrow=num_existing_rows+1, index=False, header=False) 
            # #选项5为语音是语音秒数，startcol = 19
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=19 ,startrow=num_existing_rows+1, index=False, header=False) 
            # #选项6为语音是语音秒数，startcol = 22
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=22 ,startrow=num_existing_rows+1, index=False, header=False)  

            #题目写入
            new_data_ques.to_excel(writer, sheet_name='Sheet1', startcol=2,startrow=num_existing_rows+1, index=False, header=False)  #head = flase 就是不用写表格头
            #选项abcd写入
            new_data_A.to_excel(writer, sheet_name='Sheet1', startcol=6  ,startrow=num_existing_rows+1, index=False, header=False)  
            new_data_B.to_excel(writer, sheet_name='Sheet1', startcol=9  ,startrow=num_existing_rows+1, index=False, header=False)  
            new_data_C.to_excel(writer, sheet_name='Sheet1', startcol=12  ,startrow=num_existing_rows+1, index=False, header=False)  
            new_data_D.to_excel(writer, sheet_name='Sheet1', startcol=15  ,startrow=num_existing_rows+1, index=False, header=False)  
            #选项ef写入
            new_data_E.to_excel(writer, sheet_name='Sheet1', startcol=18  ,startrow=num_existing_rows+1, index=False, header=False)  #单选配伍题
            new_data_F.to_excel(writer, sheet_name='Sheet1', startcol=21  ,startrow=num_existing_rows+1, index=False, header=False)  #单选配伍题
            """
            因为答案写入非常复杂,需要对应不同列写入值1和2.
            因此将答案和解析通过函数single_ans写入。
            函数接受，数据，写入器，写入行索引。
            """
            #答案写入,需要将答案剥离
            ans_all.xlsx_multi_ans(ans_data,writer,num_existing_rows+1,length_choice)
            #解析写入，需要把解析剥离
            ans_all.xlsx_multi_ans(anal_data,writer,num_existing_rows+1,anal=True)

    except Exception as e:
        print(f"写入Excel文件时出错: {e}")
        return
    print("数据已成功写入Excel文件的目标列!")


#判断题
def judge_write(judge_block,excel_file_path = '客观题.xlsx'):
    """将列表读取，并写入到已经存在的列表zgt的E列。

    Args:
        ques_list (_type_): 列表，列表包含题目、选项、答案、解析。
        excel_file_path:相对路径
    """
    try:
        df = pd.read_excel(excel_file_path, engine='openpyxl')
    except FileNotFoundError:
        print(f"文件路径 {excel_file_path} 不存在")
        return
    
    #读取存在的数据，确定写入的行索引
    existing_data = pd.read_excel(excel_file_path, sheet_name='Sheet1',usecols=[2])  #题目、选项A,B,C,D
    last_row_index = existing_data.iloc[:, 0].last_valid_index()
    # 如果最后一个非空单元格所在的行索引为None，则说明该列所有单元格都为空，下一行索引为0；否则，下一行索引为最后一个非空单元格所在的行索引加1。
    if last_row_index is None:
        num_existing_rows = 0
    else:
        num_existing_rows = last_row_index + 1
    #将题目和答案块拆开
    ques_choice_list = judge_block[0]
    ans_anal_list = judge_block[1]
    #将答案解析洗出来
    # 原始列表
    lst = ans_anal_list
    # 去掉开头和结尾的空格和换行符，并将关键词替换为空字符串
    answer = lst[0].strip().replace('答案 ', '')
    try:
        explanation = lst[1].strip().replace('解析 ', '') 
    except:
        explanation = ""
    # 将答案和解析写入到原始列表中
    lst[0] = answer
    lst.append(explanation)
    ans_anal_list = lst    
    # print("答案和解析清理后：",ans_anal_list)

    #题目
    new_data_ques = pd.DataFrame(ifexistedlist(list=ques_choice_list,index=0)) 
    #选项
    new_data_A = pd.DataFrame(ifexistedlist(list=ques_choice_list,index=1)) 
    new_data_B = pd.DataFrame(ifexistedlist(list=ques_choice_list,index=2)) 


    #答案
    ans_data = ifexistedlist(list=ans_anal_list,index=0)
    #解析需要拼装成特别的形式  
    anal_data = ifexistedlist2(list=ans_anal_list,index=1) 

    #题目类型
    type_of_choice = pd.DataFrame([1]) 
    #输入类型
    type_of_input = pd.DataFrame([1]) 
    voice_time = pd.DataFrame([0])
    
    # 将新数据写入Excel文件的E列
    try:
        with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a',if_sheet_exists = 'overlay') as writer: #续写Mode改为'a'
            #写入题目类型，startcol  = 0
            type_of_choice.to_excel(writer, sheet_name='Sheet1', startcol=0 ,startrow=num_existing_rows+1, index=False, header=False) 
            #写入题干输入类型，startcol  = 1
            type_of_input.to_excel(writer, sheet_name='Sheet1', startcol=1 ,startrow=num_existing_rows+1, index=False, header=False) 
            # #写入题目音频秒数，startcol = 3
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=3 ,startrow=num_existing_rows+1, index=False, header=False) 
            #写入选项输入类型，startcol = 4
            type_of_input.to_excel(writer, sheet_name='Sheet1', startcol=4 ,startrow=num_existing_rows+1, index=False, header=False) 
            #选项1为语音是语音秒数，startcol = 7
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=7 ,startrow=num_existing_rows+1, index=False, header=False) 
            # #选项2为语音是语音秒数，startcol = 10
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=10 ,startrow=num_existing_rows+1, index=False, header=False) 
            # #选项3为语音是语音秒数，startcol = 13
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=13 ,startrow=num_existing_rows+1, index=False, header=False) 
            # #选项4为语音是语音秒数，startcol = 16
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=16 ,startrow=num_existing_rows+1, index=False, header=False) 
            # #选项5为语音是语音秒数，startcol = 19
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=19 ,startrow=num_existing_rows+1, index=False, header=False) 
            # #选项6为语音是语音秒数，startcol = 22
            # voice_time.to_excel(writer, sheet_name='Sheet1', startcol=22 ,startrow=num_existing_rows+1, index=False, header=False)  

            #题目写入
            new_data_ques.to_excel(writer, sheet_name='Sheet1', startcol=2,startrow=num_existing_rows+1, index=False, header=False)  #head = flase 就是不用写表格头
            #选项abcd写入
            new_data_A.to_excel(writer, sheet_name='Sheet1', startcol=6  ,startrow=num_existing_rows+1, index=False, header=False)  
            new_data_B.to_excel(writer, sheet_name='Sheet1', startcol=9  ,startrow=num_existing_rows+1, index=False, header=False)  

            """
            因为答案写入非常复杂,需要对应不同列写入值1和2.
            因此将答案和解析通过函数single_ans写入。
            函数接受，数据，写入器，写入行索引。
            """
            #答案写入,需要将答案剥离
            ans_all.xlsx_judge_ans(ans_data,writer,num_existing_rows+1)
            #解析写入，需要把解析剥离
            ans_all.xlsx_judge_ans(anal_data,writer,num_existing_rows+1,anal=True)

    except Exception as e:
        print(f"写入Excel文件时出错: {e}")
        return
    print("数据已成功写入Excel文件的目标列!")
    
# 测试函数
if __name__ == '__main__': 
    data_list = [1, 2, 3, 4, 5]
    excel_file_path = '客观题.xlsx'
    write_to_xlsx(data_list, excel_file_path)