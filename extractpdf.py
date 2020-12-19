from PyPDF2 import PdfFileReader

pdfFileObj = open("20201204_historische-schiessen_daten_2021_inkl_im_felde_website_ssv.pdf", "rb")
pdfReader  = PdfFileReader(pdfFileObj,strict = False)
[print(page.extractText()) for page in pdfReader.pages]