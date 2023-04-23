import pandas as pd
def write_to_xlsx(ques_choice_list,excel_file_path = '主观题.xlsx'):
    """将列表读取，并写入到已经存在的列表zgt的E列。

    Args:
        ques_list (_type_): 列表
        excel_file_path:相对路径
    """
    try:
        df = pd.read_excel(excel_file_path, engine='openpyxl')
    except FileNotFoundError:
        print(f"文件路径 {excel_file_path} 不存在")
        return
    #读取存在的数据
    existing_data = pd.read_excel(excel_file_path, sheet_name='Sheet1',usecols=[7,8])
    last_row_index = existing_data.iloc[:, 1].last_valid_index()
    # 如果最后一个非空单元格所在的行索引为None，则说明该列所有单元格都为空，下一行索引为0；否则，下一行索引为最后一个非空单元格所在的行索引加1。
    if last_row_index is None:
        num_existing_rows = 0
    else:
        num_existing_rows = last_row_index + 1

    new_data_ques = pd.DataFrame(ques_choice_list[0]) 
    input_type  = pd.DataFrame(ques_choice_list[1]) #输入类型


    
    # 将新数据写入Excel文件的E列
    try:
        with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a',if_sheet_exists = 'overlay') as writer: #续写Mode改为'a'
            # #从E(4列)列写入题目
            new_data_ques.to_excel(writer, sheet_name='Sheet1', startcol=8,startrow=num_existing_rows+1, index=False, header=False)  #head = flase 就是不用写表格头
            input_type.to_excel(writer, sheet_name='Sheet1', startcol=7  ,startrow=num_existing_rows+1, index=False, header=False)  #改为切片，每四次一个切片

            #从I(8列)列写入题目
            # new_data.to_excel(writer, sheet_name='Sheet1', startcol=8,startrow=1, index=False, header=False)  #head = flase 就是不用写表格头
    except Exception as e:
        print(f"写入Excel文件时出错: {e}")
        return
    print("数据已成功写入Excel文件的目标列!")
    
# 测试函数
if __name__ == '__main__': 
    data_list = [1, 2, 3, 4, 5]
    excel_file_path = '客观题.xlsx'
    write_to_xlsx(data_list, excel_file_path)