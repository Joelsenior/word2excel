import pandas as pd

def clean_excel_data(filename):
    # 读取 Excel 数据
    df = pd.read_excel(filename)

    # 清理第 3 列，题干列，索引为2
    df.iloc[:, 2] = df.iloc[:, 2].apply(lambda x: x.split(".", 1)[-1].strip() if isinstance(x, str) and x[0].isdigit() and x.find(".") != -1 else x)

    # 清理第 7 列
    df.iloc[:, 6] = df.iloc[:, 6].apply(lambda x: x.split(".", 1)[-1].strip() if isinstance(x, str) and x[0] in ['A', 'B', 'C', 'D', 'E', 'F'] and x.find(".") != -1 else x)
    df.iloc[:, 9] = df.iloc[:, 9].apply(lambda x: x.split(".", 1)[-1].strip() if isinstance(x, str) and x[0] in ['A', 'B', 'C', 'D', 'E', 'F'] and x.find(".") != -1 else x)
    df.iloc[:, 12] = df.iloc[:, 12].apply(lambda x: x.split(".", 1)[-1].strip() if isinstance(x, str) and x[0] in ['A', 'B', 'C', 'D', 'E', 'F'] and x.find(".") != -1 else x)
    df.iloc[:, 15] = df.iloc[:, 15].apply(lambda x: x.split(".", 1)[-1].strip() if isinstance(x, str) and x[0] in ['A', 'B', 'C', 'D', 'E', 'F'] and x.find(".") != -1 else x)
    
    df.iloc[:, 18] = df.iloc[:, 18].apply(lambda x: x.split(".", 1)[-1].strip() if isinstance(x, str) and x[0] in ['A', 'B', 'C', 'D', 'E', 'F'] and x.find(".") != -1 else x)
    df.iloc[:, 21] = df.iloc[:, 21].apply(lambda x: x.split(".", 1)[-1].strip() if isinstance(x, str) and x[0] in ['A', 'B', 'C', 'D', 'E', 'F'] and x.find(".") != -1 else x)
    # 写回 Excel 表格
    writer = pd.ExcelWriter(filename)
    df.to_excel(writer, index=False)
    writer.save()
    print("excel清洗完成!")

# 测试函数
filename = "701-2016年客观题.xlsx"
clean_excel_data(filename)
