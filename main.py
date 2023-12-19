from Bills_PDF import generate_word

ExcelFileName = 'D:/Code/python/2023年半年度客户函证制函（研发）0717/2023半年度天职客户函证0717副本.xlsx'
TemplateFileName = 'D:/Code/python/2023年半年度客户函证制函（研发）0717/20230630客户函证文字模板/20230630天职客户函证模板.docx'
merge_pdf_name = 'merge.pdf'
temp_path = 'D:/Code/python/2023年半年度客户函证制函（研发）0717/三个机构回函码/天职/天职寄件二维码1-100.pdf'

if __name__ == '__main__':
    # 根据Excel客户信息文档ExcelFileName、及对账单模板TemplateFileName生成对账单Word文档
    generate_word(ExcelFileName, TemplateFileName)
    # # 将对账单Word文档转成PDF文档
    # word_to_pdf(word_path)
    # # 合并PDF文档
    # merge_pdf(OutputFolder, merge_pdf_name)

    # get_qrcode('hello', qrCodeOutPutFolder)
    # get_barcode('JDVE08431183145', OutputFolder)
    # split_pdf(temp_path,OutputFolder)


