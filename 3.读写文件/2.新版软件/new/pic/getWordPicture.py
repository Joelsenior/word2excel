import zipfile
import time
import os

def getPicture(doc_path,image_path):
    start_time = time.time()
    doc = zipfile.ZipFile(doc_path)     
    #doc是个ZipFile对象，在创建ZipFile对象时，默认参数是r，表示读已经存在的zip文件，还有w和a，分别为新建和覆盖一个zip文件

    for info in doc.infolist():         #infolist()是获取zip文件中所有文件的信息，返回一个列表
    #print(info.filename.endswith)
        if (info.filename.endswith((".jpeg",".png",".gif"))):       #找到文档中图片并用extract方法解压出来，即提取文档中的图片
            print(info)
            doc.extract(info,image_path)
    doc.close()
    end_time = time.time()
    print(f"总共耗时：{end_time-start_time} 秒")

def getDirName():
    t1 = time.localtime(time.time())
    dir_name = f"{t1.tm_year}-{t1.tm_mon}-{t1.tm_mday} {t1.tm_hour}_{t1.tm_min}"
    print(dir_name)
    return dir_name


if __name__ == "__main__":
    getPicture("Demo\example.docx",r"D:\桌面\new\pic\Demo")
    
