from Bills_PDF import generate_word



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


