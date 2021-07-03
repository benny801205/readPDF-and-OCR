import os
import pyPdf  
import subprocess
import time
import cv2
from PIL import Image
import pytesseract
from xlutils.copy import copy
import xlrd
import xlwt
import shutil

def getPDFContent(path):
    content = ""
    p = file(path, "rb")
    pdf = pyPdf.PdfFileReader(p)

    content += pdf.getPage(0).extractText() + "\n"
  #  content = " ".join(content.replace(u"\xa0", " ").strip().split())     
    return content 

def ORC(url,i):
    fromfile=url
    i=str(i)
    GHOSTSCRIPTCMD=r'C:\Program Files\gs\gs9.20\bin\gswin64.exe'
    arglist = [GHOSTSCRIPTCMD,
    "-dBATCH",
    "-dNOPAUSE",
	"-dLastPage=1",
	"-q",
	"-sDEVICE=jpeg",
	"-r350",
    "-sOutputFile=C:\\Users\\6601\\.spyder\\Mgroup\\process\\" + i +'.jpg',
    fromfile]
    print(subprocess.Popen(
            arglist,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,shell=True))
    while not os.path.exists("C:\\Users\\6601\\.spyder\\Mgroup\\process\\" + i +'.jpg'):
        print 'False'
        time.sleep(1) 
    print os.path.exists("C:\\Users\\6601\\.spyder\\Mgroup\\process\\" + i +'.jpg')
    img = cv2.imread("C:\\Users\\6601\\.spyder\\Mgroup\\process\\" + i +'.jpg')
    cv2.imwrite("C:\\Users\\6601\\.spyder\\Mgroup\\tesseract\\" + i +'.jpg', img[0:img.shape[0]/2,0:img.shape[1]])
    img_1=Image.open("C:\\Users\\6601\\.spyder\\Mgroup\\tesseract\\" + i +'.jpg')
    a=pytesseract.image_to_string(img_1).split()
    b=str(a[-3])+' '+str(a[-2])+' '+str(a[-1])
    output=filter(str.isalnum, b)
    print output
    cv2.imwrite("C:\\Users\\6601\\.spyder\\Mgroup\\finial\\" + c +'.jpg',img[0:img.shape[0]/2,0:img.shape[1]])
    return output


#########EXCEL prepare########################
rb = xlrd.open_workbook('D:\\1.xls')
rs = rb.sheet_by_index(0) 
wb = copy(rb)
#通过get_sheet()获取的sheet有write()方法
ws = wb.get_sheet(0)
#########################





List=[]
ListName=[]
rows=0
for dirPath, dirNames, fileNames in os.walk("C:\\Users\\6601\\.spyder\\Mgroup\\2.5.5\\"):
    for f in fileNames:
        fileurl=os.path.join(dirPath, f)
        List.append(fileurl)
        a=getPDFContent(fileurl).encode('utf8').strip()
        if a is '':
            a=ORC(fileurl,rows)
        else:
            a=str(a.split()[-3])+str(a.split()[-2])+str(a.split()[-1])
            a=filter(str.isalnum, a)
        print a,rows
        ws.write(rows, 0, a)
        print f
        ws.write(rows, 1, filter(str.isalnum,str(f)))
        shutil.copyfile("C:\\Users\\6601\\.spyder\\Mgroup\\2.5.5\\"+f,"C:\\Users\\6601\\.spyder\\Mgroup\\2.5.5.1\\"+a+'.pdf')
        rows=rows+1
        wb.save('D:\\1.xls')
#        print f
wb.save('D:\\1.xls')
print 'end'
