import pandas as pd
def write_to_xlsx(ques_choice_list,excel_file_path = '客观题.xlsx'):
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
    existing_data = pd.read_excel(excel_file_path, sheet_name='Sheet1',usecols=[0,5,8])
    last_row_index = existing_data.iloc[:, 0].last_valid_index()
    # 如果最后一个非空单元格所在的行索引为None，则说明该列所有单元格都为空，下一行索引为0；否则，下一行索引为最后一个非空单元格所在的行索引加1。
    if last_row_index is None:
        num_existing_rows = 0
    else:
        num_existing_rows = last_row_index + 1 
    type_of_choice = pd.DataFrame(ques_choice_list[1])  #获得类型值
    right_ans = pd.DataFrame([1]) 
    wrong_ans = pd.DataFrame([2]) 
    # 将新数据写入Excel文件的E列
    try:
        with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a',if_sheet_exists = 'overlay') as writer: #续写Mode改为'a'
                type_of_choice.to_excel(writer, sheet_name='Sheet1', startcol=0 ,startrow=num_existing_rows+1, index=False, header=False) 
                for index,ans in enumerate (ques_choice_list[0],start=1):
                    if ans == 'X' or '×':
                        #如果答案是A,则将1写入5行，2写入8,1,1,14行
                        right_ans.to_excel(writer, sheet_name='Sheet1', startrow=index+num_existing_rows,startcol=5, index=False, header=False)  #head = flase 就是不用写表格头
                        wrong_ans.to_excel(writer, sheet_name='Sheet1', startrow=index+num_existing_rows,startcol=8, index=False, header=False) 
                    
                    if ans == 'V' or '√':
                        right_ans.to_excel(writer, sheet_name='Sheet1', startrow=index+num_existing_rows,startcol=8, index=False, header=False)  #head = flase 就是不用写表格头
                        wrong_ans.to_excel(writer, sheet_name='Sheet1', startrow=index+num_existing_rows,startcol=5, index=False, header=False) 

    except Exception as e:
        print(f"写入Excel文件时出错: {e}")
        return
    print("数据已成功写入Excel文件的目标列！")
    
# 测试函数
if __name__ == '__main__': 
    data_list = [1, 2, 3, 4, 5]
    excel_file_path = '客观题.xlsx'
    write_to_xlsx(data_list, excel_file_path)