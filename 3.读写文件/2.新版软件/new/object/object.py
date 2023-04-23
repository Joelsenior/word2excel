from contentextract import extract_content
from blockextract_from_docx import extract_block
from set_logs import setup_logger

def main():
    # 主函数代码
    logger = setup_logger()
    logger.info('-----------------------程序开始运行-----------------------')
    extract_content(extract_block('客观题题目.docx',logger))
    logger.info('-----------------------程序结束运行-----------------------\n\n')

if __name__ == "__main__":
    main()



