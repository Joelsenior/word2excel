import docx
import os, re

def get_pictures(word_path, result_path):
    """
    图片提取
    :param word_path: word路径
    :return: 
    """
    try:
        doc = docx.Document(word_path)   # 使用docx库打开word文档
        dict_rel = doc.part._rels        # 获取文档中的关系字典
        for rel in dict_rel:             # 遍历关系字典
            rel = dict_rel[rel]
            if "image" in rel.target_ref:    # 如果是图片关系
                if not os.path.exists(result_path):   # 如果结果文件夹不存在，则创建
                    os.makedirs(result_path)
                img_name = re.findall("/(.*)", rel.target_ref)[0]   # 通过正则表达式获取图片文件名
                word_name = os.path.splitext(word_path)[0]     # 获取word文档名称
                if os.sep in word_name:
                    new_name = word_name.split('\\')[-1]    # 分离文件名
                else:
                    new_name = word_name.split('/')[-1]    # 分离文件名
                img_name = f'{new_name}-'+'.'+f'{img_name}'   # 设置新的图片名称
                with open(f'{result_path}/{img_name}', "wb") as f:   # 以二进制写入文件
                    f.write(rel.target_part.blob)    # 写入图片数据
    except:
        pass

if __name__ == '__main__':

    #获取文件夹下的word文档列表,路径自定义
    get_pictures("example.docx", "Demo")
