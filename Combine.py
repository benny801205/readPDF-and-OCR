import xlrd
import os
import subprocess
import time


def Merge(filelist,outpathname):
    GHOSTSCRIPTCMD=r'C:\Program Files\gs\gs9.20\bin\gswin64.exe'
    arglist = [GHOSTSCRIPTCMD,
    "-dBATCH",
    "-dNOPAUSE",
	"-q",
	"-sDEVICE=pdfwrite",
	"-dPDFSETTINGS=/prepress",
    "-sOutputFile=" + outpathname,
	]+filelist
    #print arglist
    print(subprocess.Popen(
           arglist,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,shell=True))
    while not os.path.exists(outpathname):
        print 'False'
        time.sleep(1) 


#######################


data = xlrd.open_workbook('Equipment Number List.xlsx')
table = data.sheet_by_name(u'Sheet2 (1)')
cell_A2=table.cell(1,0).value
nrows = table.nrows
ncols = table.ncols
c=table.row_values(0)


for i in range(2):
    name_1=table.cell(1+i,0).value
    name_2=table.cell(1+i,1).value.replace('/','')
    filelist=[]
    name_3=filter(str.isalnum, str(name_2))
    dir_name2 = 'D:\\Mgroup\\' + name_1 + '\\' + name_2 + '\\'#根目錄資料夾位置
    outpathname=dir_name2 + name_3 + '.pdf'
    print outpathname,i
   # print outpathname
    for dirPath, dirNames, fileNames in os.walk(dir_name2):
#        print dirPath,dirNames
        for f in fileNames:
            #print os.path.join(dirPath, f)
            filelist.append(os.path.join(dirPath, f))
    #print filelist
    if filelist != []:
        Merge(filelist,outpathname)
    
    
    
    
#print filelist[0] + ',' + filelist[1]
