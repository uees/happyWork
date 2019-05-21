# -*- coding: utf-8 -*-
import os

from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from PyPDF2 import PdfFileWriter, PdfFileReader


def create_watermark(text, path=None):
    if path:
        f_pdf = os.path.join(path, 'mark.pdf')
    else:
        f_pdf = 'mark.pdf'

    w_pdf = 20 * cm
    h_pdf = 20 * cm

    c = canvas.Canvas(f_pdf, pagesize=(w_pdf, h_pdf))
    c.setFillAlpha(0.6)  # 设置透明度
    c.drawString(3.5 * cm, 7 * cm, text)
    c.showPage()
    c.save()

    return f_pdf


def add_watermark(pdf_file_mark, pdf_file_in, pdf_file_out):
    with open(pdf_file_in, 'rb') as fp:
        pdf_input = PdfFileReader(fp)

        # PDF文件被加密了
        if pdf_input.getIsEncrypted():
            print('该PDF文件被加密了.')
            # 尝试用空密码解密
            try:
                pdf_input.decrypt('')
            except Exception:
                print('尝试用空密码解密失败.')
                return False
            else:
                print('用空密码解密成功.')

        # 获取PDF文件的页数
        pageNum = pdf_input.getNumPages()

        with open(pdf_file_mark, 'rb') as mfp:
            pdf_output = PdfFileWriter()
            # 读入水印pdf文件
            pdf_watermark = PdfFileReader(mfp)

            # 给每一页打水印
            for i in range(pageNum):
                page = pdf_input.getPage(i)
                page.mergePage(pdf_watermark.getPage(0))
                page.compressContentStreams()  # 压缩内容
                pdf_output.addPage(page)

            with open(pdf_file_out, 'wb') as wfp:
                pdf_output.write(wfp)
