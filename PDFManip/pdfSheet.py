import PyPDF2, os

# pdfFileObj = open('automate.pdf', 'rb')
# pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

##############Si esta encriptado el PDF#############
# pdfReader.decrypt('contraseña')

# print(pdfReader.numPages)

# pageObj = pdfReader.getPage(0)

# text = pageObj.extractText()
# print(text)

#################################################
############## Creando PDFs #####################
#################################################


##########Copiando Paginas##############
# pdfWriter = PyPDF2.PdfFileWriter()
# for pageNum in range(pdfReader.numPages):
#     pageObj = pdfReader.getPage(pageNum)
#     pdfWriter.addPage(pageObj)

# pdfOutputFile = open('combinedminutes.pdf', 'wb')
# pdfWriter.write(pdfOutputFile)
# pdfOutputFile.close()


#########Rotando Paginas###############
# minutesFile = open('automate.pdf', 'rb')
# pdfReader = PyPDF2.PdfFileReader(minutesFile)
# page = pdfReader.getPage(0)
# page.rotateClockwise(90)
# pdfWriter = PyPDF2.PdfFileWriter()
# pdfWriter.addPage(page)
# resultPdfFile = open('rotatedPage.pdf', 'wb')
# pdfWriter.write(resultPdfFile)
# resultPdfFile.close()
# minutesFile.close()


#############Overlaying Pages###############
# minutesFile = open('automate.pdf', 'rb')
# pdfReader = PyPDF2.PdfFileReader(minutesFile)
# minutesFirstPage = pdfReader.getPage(0)
# pdfWatermarkReader = PyPDF2.PdfFileReader(open('watermark.pdf', 'rb'))
# minutesFirstPage.mergePage(pdfWatermarkReader.getPage(0))
# pdfWriter = PyPDF2.PdfFileWriter()
# pdfWriter.addPage(minutesFirstPage)

# for pageNum in range(1, pdfReader.numPages):
#     pageObj = pdfReader.getPage(pageNum)
#     pdfWriter.addPage(pageObj)

# resultPdfFile = open('watermarkedCover.pdf', 'wb')
# pdfWriter.write(resultPdfFile)
# minutesFile.close()
# resultPdfFile.close()


################Encriptando PDFs####################
# pdfFile = open('automate.pdf', 'rb')
# pdfReader = PyPDF2.PdfFileReader(pdfFile)
# pdfWriter = PyPDF2.PdfFileWriter()
# for pageNum in range(pdfReader.numPages):
#     pdfWriter.addPage(pdfReader.getPage(pageNum))
# pdfWriter.encrypt('swordfish')
# resultPdf = open('encryptedminutes.pdf', 'wb')
# pdfWriter.write(resultPdf)
# resultPdf.close()


#########################################################
## Project: Combining Select Pages from Many PDFs #######
#########################################################

pdfWriter = PyPDF2.PdfFileWriter()
pdfFiles = []
for filename in os.listdir('.'):
    if filename.endswith('.pdf'):
        pdfFiles.append(filename)

pdfFiles.sort(key = str.lower)

# Loop through all the PDF files.
for pdfFile in pdfFiles:
    pdfFileObj = open(pdfFile, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
 ## Loop through all the pages (except the first) and add them.
    for pageNum in range(1,pdfReader.numPages):
        pageObj = pdfReader.getPage(pageNum)
        pdfWriter.addPage(pageObj)


# #Save the resulting PDF to a file.
pdfOutput = open('TodosPDFSCombinados.pdf', 'wb')
pdfWriter.write(pdfOutput)
pdfOutput.close()


# pdfFileObj.close()