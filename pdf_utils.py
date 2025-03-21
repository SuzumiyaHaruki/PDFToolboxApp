from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from pdf2image import convert_from_path
from PIL import Image
from docx import Document
import pdfkit
import comtypes.client
import os

# --------------------------
# 文件操作模块
# --------------------------

def merge_pdfs(pdf_list, output_path):
    """合并多个PDF文件
    Args:
        pdf_list (list): 要合并的PDF文件路径列表
        output_path (str): 合并后的输出文件路径
    """
    merger = PdfMerger()
    for pdf in pdf_list:
        merger.append(pdf)  # 按顺序追加文件
    merger.write(output_path)
    merger.close()

def split_pdf(input_pdf, start_page, end_page, output_pdf):
    """拆分PDF文件
    Args:
        input_pdf (str): 输入PDF路径
        start_page (int): 起始页码（从1开始）
        end_page (int): 结束页码（包含该页）
        output_pdf (str): 输出文件路径
    Note: PyPDF2的页码从0开始，需要做-1转换
    """
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    # 注意：PyPDF2的页面索引从0开始
    for page in range(start_page - 1, end_page):
        writer.add_page(reader.pages[page])

    with open(output_pdf, "wb") as output_file:
        writer.write(output_file)

# --------------------------
# 安全模块
# --------------------------

def encrypt_pdf(input_pdf, password, output_pdf):
    """给PDF文件添加密码保护
    Args:
        input_pdf (str): 输入文件路径
        password (str): 加密密码
        output_pdf (str): 加密后的输出路径
    """
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    writer.encrypt(password)  # 设置密码

    with open(output_pdf, "wb") as output_file:
        writer.write(output_file)

def decrypt_pdf(input_pdf, password, output_pdf):
    """解密受保护的PDF文件
    Args:
        input_pdf (str): 加密的PDF文件路径
        password (str): 解密密码
        output_pdf (str): 解密后的输出路径
    """
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    if reader.is_encrypted:
        reader.decrypt(password)  # 尝试解密

    for page in reader.pages:
        writer.add_page(page)

    with open(output_pdf, "wb") as output_file:
        writer.write(output_file)

# --------------------------
# 内容处理模块
# --------------------------

def extract_text_from_pdf(input_pdf, output_doc):
    """从PDF提取文本到Word文档
    Args:
        input_pdf (str): PDF文件路径
        output_doc (str): 输出的Word文档路径
    Note: 提取的文本可能受PDF格式影响，保留换行符
    """
    reader = PdfReader(input_pdf)
    text = ""
    for page in reader.pages:
        text += page.extract_text() + "\n"  # 添加换行分隔不同页面

    word_doc = Document()
    word_doc.add_paragraph(text)
    word_doc.save(output_doc)

# --------------------------
# 格式转换模块
# --------------------------

def pdf_to_images(input_pdf, output_folder):
    """将PDF每页转换为PNG图片
    Args:
        input_pdf (str): PDF文件路径
        output_folder (str): 图片输出目录
    Note: 依赖poppler-utils，Windows需安装poppler
    """
    images = convert_from_path(input_pdf)
    for i, img in enumerate(images):
        img_path = os.path.join(output_folder, f"page_{i+1}.png")
        img.save(img_path, "PNG")  # 保存为PNG格式

def images_to_pdf(image_list, output_pdf):
    """将多张图片合并为PDF文件
    Args:
        image_list (list): 图片路径列表
        output_pdf (str): 输出PDF路径
    Note: 所有图片将合并为多页PDF，默认使用RGB模式
    """
    # 转换为RGB模式避免颜色问题
    images = [Image.open(img).convert("RGB") for img in image_list]
    images[0].save(output_pdf, save_all=True, append_images=images[1:])

def word_to_pdf(input_docx, output_pdf):
    """将Word文档转换为PDF（仅限Windows）
    Args:
        input_docx (str): Word文档路径
        output_pdf (str): 输出PDF路径
    Warning: 依赖Microsoft Word客户端，必须安装Office
    """
    word = comtypes.client.CreateObject("Word.Application")
    doc = word.Documents.Open(input_docx)
    doc.SaveAs(output_pdf, FileFormat=17)  # 17对应PDF格式
    doc.Close()
    word.Quit()

def html_to_pdf(input_html, output_pdf):
    """将HTML文件转换为PDF
    Args:
        input_html (str): HTML文件路径
        output_pdf (str): 输出PDF路径
    Note: 需要提前安装wkhtmltopdf并配置路径
    """
    # Windows默认安装路径，其他系统需要修改
    config = pdfkit.configuration(
        wkhtmltopdf="C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe"
    ) 
    pdfkit.from_file(input_html, output_pdf, configuration=config)