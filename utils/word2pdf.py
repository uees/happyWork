import os

from win32com.client import constants, gencache


def word2pdf(wordFile, pdfFile):
    word = gencache.EnsureDispatch("Word.Application")
    doc = word.Documents.Open(wordFile, ReadOnly=1)
    doc.ExportAsFixedFormat(pdfFile, constants.wdExportFormatPDF,
                            Item=constants.wdExportDocumentWithMarkup,
                            CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
    word.Quit(constants.wdDoNotSaveChanges)


def word2pdfbydir(dir_path):
    wordFiles = [fn for fn in os.listdir(dir_path) if fn.endswith(('.doc', '.docx'))]

    for wordFile in wordFiles:
        pdf_path = os.path.join(dir_path, 'pdf')
        if not os.path.isdir(pdf_path):
            os.mkdir(pdf_path)

        fname, ext = os.path.splitext(wordFile)
        pdfFile = os.path.join(pdf_path, f"{fname}.pdf")
        wordFile = os.path.join(dir_path, wordFile)
        word2pdf(wordFile, pdfFile)


if __name__ == "__main__":
    dir_path = "C:\\Users\\pzb\\Desktop\\COC"
    word2pdfbydir(dir_path)
