import PyPDF2

fileObj = open('./guidelines.pdf', 'rb')
pdfReader = PyPDF2.PdfFileReader(fileObj)
text = ""

for i in range(0, pdfReader.numPages - 1):
    page = pdfReader.getPage(i)
    text += page.extractText()

with open('./guidelines.txt', 'w') as outfile:
    outfile.write(text)
    
