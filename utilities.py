import PyPDF2
import qrcode
# pip install python-barcode
# pip install pillow
import barcode
from barcode.writer import ImageWriter  # 引入一位条码库写模块


# 生成二维码
def get_qrcode(code, output):
    img = qrcode.make(code, version=4, border=1, box_size=2)
    img.save(output)  # 保存图片


# 生成条码
def get_barcode(code, output):
    bar_code = barcode.generate('code128', code,
                                writer=ImageWriter(),
                                output=output,
                                writer_options={"module_width": 0.3, "module_height": 5, "font_size": 10})


def split_pdf(filename, output_folder):
    pdf_file = open(filename, 'rb')
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    total_pages = len(pdf_reader.pages)
    for i in range(total_pages):
        output_pdf = PyPDF2.PdfWriter()
        output_pdf.add_page(pdf_reader.pages[i])
        output_filename = output_folder + str(i) + '.pdf'
        with open(output_filename, 'wb') as output:
            output_pdf.write(output)
    pdf_file.close()


# split_pdf('D:/大数据管理部/询证函/20230831供应商往来对账函-天津-已盖章.pdf', 'D:/大数据管理部/询证函/split/天津/')
# get_qrcode('hello', 'd:/')
# get_barcode('JDVE08431183145', 'd:/')
