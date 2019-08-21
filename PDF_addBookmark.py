import PyPDF2
 #word ppt等轉成的PDF檔無法翻轉
pdf_in = open('F1101F1102F1103F1104F1105F1106F1107F1108.pdf', 'rb')
pdf_reader = PyPDF2.PdfFileReader(pdf_in)
pdf_writer = PyPDF2.PdfFileWriter()#創一個虛擬空白的PDF檔
 
for pagenum in range(pdf_reader.numPages):
    page = pdf_reader.getPage(pagenum)#將正本一頁一頁加入虛擬的PDF內
#    if pagenum % 2:
    #page.rotateClockwise(180)
    pdf_writer.addPage(page)



pdf_writer.addBookmark('Hello, World Bookmark', 0, parent=None)
#pdf_writer.addBookmark('bbbbbbb', 2, parent=None)
pdf_out = open('XE35-D1124-001_RevA.1(SIGNED).pdf', 'wb')
pdf_writer.write(pdf_out)
pdf_out.close()
pdf_in.close()