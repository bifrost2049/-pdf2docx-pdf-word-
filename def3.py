from pdf2docx import Converter  # 导入Converter类用于PDF到Word的转换
import os  # 操作系统接口库，用于检查文件是否存在


def pdf_to_word(pdf_path, word_path):
    # 检查PDF文件是否存在
    if not os.path.exists(pdf_path):  # 检查指定路径的文件是否存在
        print(f"文件不存在: {pdf_path}")  # 如果文件不存在，打印错误信息
        return  # 返回函数，不继续执行

    # 检查Word文件是否被占用
    if os.path.exists(word_path):
        try:
            with open(word_path, 'a'):  # 尝试以追加模式打开文件
                pass
        except IOError as e:
            print(f"无法访问文件 {word_path}: {e}")
            return

    # 创建Converter对象
    cv = Converter(pdf_path)  # 使用Converter打开指定路径的PDF文件

    try:
        # 将PDF转换为Word文档
        cv.convert(word_path, start=0, end=None)  # 将整个PDF转换为Word文档
    except Exception as e:
        print(f"转换过程中发生错误: {e}")
    finally:
        # 关闭Converter对象
        cv.close()  # 关闭Converter对象


# 使用函数
pdf_to_word(
    r"C:\Users\confinement\Desktop\翻译.pdf",
    'output1.docx')  # 调用pdf_to_word函数，将指定PDF转换为'output.docx'





