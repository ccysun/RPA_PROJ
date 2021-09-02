
    import win32com.client
    pptApp= win32com.client.gencache.EnsureDispatch ("Powerpoint.Application")
    presentation = pptApp.Presentations.Add()
    slide = presentation.Slides.Add(1, 12)
    myDiamond = slide.Shapes.AddShape(4, Top=100,Left=100, Width=20, Height=20)
    presentation.SaveAs("C:\\RPA\\Test\\myPowerPoint.pptx",1)
    
    myDiamond
    




















import openpyxl as op
import pptx
import os
import win32com.client

import smtplib

os.chdir(r'C:\Users\aju.mathew.thomas\Desktop\PBC\Pepsi\PBC\Performance Reports\2019\PPT')
path= r'C:\Users\aju.mathew.thomas\Desktop\PBC\Pepsi\PBC\Performance Reports\2019\PPT\Summary2.xlsx'
wb = op.load_workbook(path)
ExcelApp = win32com.client.Dispatch("Excel.Application")
ExcelApp.Visible = False
workbook = ExcelApp.Workbooks.open(r'C:\Users\aju.mathew.thomas\Desktop\PBC\Pepsi\PBC\Performance Reports\2019\PPT\Summary2.xlsx')
worksheet = workbook.Worksheets("Summary")
excelrange = worksheet.Range("A2:R24")

PptApp = win32com.client.Dispatch("Powerpoint.Application")
PptApp.Visible = True
z= excelrange.Copy()
PPtPresentation = PptApp.Presentations.Open(r'C:\Users\aju.mathew.thomas\Desktop\PBC\Pepsi\PBC\Performance Reports\2019\PPT\PBC Performance Update.pptx')
pptSlide = PPtPresentation.Slides.Add(1,11)
#pptSlide.Title.Characters.Text ='Metrics'

#title = pptSlide.Shapes.Title
#title.Text ='Metrics Summary'
pptSlide.Shapes.PasteSpecial(z)
PPtPresentation.Save()

