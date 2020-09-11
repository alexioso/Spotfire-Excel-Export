from System.IO import StreamWriter, FileMode
from Spotfire.Dxp.Application.Visuals import CrossTablePlot, TablePlot
from  Spotfire.Dxp.Data.Export import DataWriterTypeIdentifiers
import tempfile
import os
#Common Language Runtime - manages execution of .NET programs
import clr
#Interop allows program to communicate with other MS office products,
# like Excel
clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel
import csv
from datetime import datetime


clr.AddReference('System.Drawing')
from System.Drawing import Color, ColorTranslator
   

#translate colors to a format that Excel understands
def rgbForExcel(r, g, b):
  return ColorTranslator.ToOle(Color.FromArgb(r, g, b))



tmp = os.path.join(tempfile.gettempdir(), str(table.Title) +".csv")
#tmp = "C:/Users/abraks/Documents/DDO/whoopesh.csv"
stream = StreamWriter(tmp)

#export text

try:
	#cast input variable as Spotfire TablePlot object to variable t
    t = table.As[TablePlot]()
    #export the data table
    t.ExportText(stream)
except Exception as e:
    print(e)
finally:
    stream.Close()

#write file to csv so Excel can open it (assuming default delimiter for Excel is comma)
reader = list(csv.reader(open(tmp, "r"), delimiter='\t'))
with open(tmp,'wb') as outfile:
    writer = csv.writer(outfile, delimiter=',')
    writer.writerows(row for row in reader)




#setup Excel session
ex = Excel.ApplicationClass()   
#make it visible (open)
ex.Visible = True
ex.DisplayAlerts = False   
#IMPORTANT: everything starts with the workbook object
workbook = ex.Workbooks.Open(tmp)


#if you want to start workbook from scratch, use this line instead:
#workbook = ex.Workbooks.Add()




#from there can add tabs (aka sheets)
for k in ["a","b","c"]:
    sh = workbook.Sheets.Add()
    sh.Name = k

#here is how you can remove a sheet
#workbook.Sheets["Sheet1"].Delete()


#after workbook level you can access sheet level

output = workbook.Sheets["iris"]


#after choosing sheet, all there is left to choose from is Range
#Various ways to do this including .Columns, .Rows, .Range

#details about horizontal alignment at 
#https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.style.horizontalalignment?view=excel-pia
output.Cells[1,1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
output.Range("B2").HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
output.Range("C1:C20").HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
output.Rows("3").HorizontalAlignment = Excel.XlHAlign.xlHAlignDistributed
output.Columns("E").HorizontalAlignment = Excel.XlHAlign.xlHAlignFill





#preformat column widths and headers
output.Columns("A").ColumnWidth = 16
output.Columns("B").ColumnWidth = 16
output.Columns("C").ColumnWidth = 16
output.Columns("D").ColumnWidth = 16
output.Columns("E").ColumnWidth = 32


output.Rows(10).RowHeight = 60
output.Rows(20).RowHeight = 30


    
#color rectangular ranges
output.Range("A1:B1").Interior.Color =  rgbForExcel(244, 176, 132)
output.Range("C1:D1").Interior.Color =  rgbForExcel(248, 203, 173)
output.Range("E1").Interior.Color =  rgbForExcel(252, 228, 214)
output.Range("A3:C20").Interior.Color =  rgbForExcel(244, 15, 224)
output.Range("D20:E30").Font.Color =  rgbForExcel(255, 230, 153)

#details at
#https://docs.microsoft.com/en-us/office/vba/api/excel.xllinestyle
output.Range("A2:B2").Borders.LineStyle = Excel.XlLineStyle.xlContinuous
output.Range("C3:E3").Borders.LineStyle = Excel.XlLineStyle.xlDashDot

output.Range("A2:A12").Font.Size = 10
output.Range("B2:B12").Font.Size = 14

output.Range("A3").Font.Bold = True
output.Range("A3").Font.Italic = True


output.Range("F2").Value2 = "Hungry"
output.Range("F3").Value2 = "Angry"
output.Range("F4").Value2 = "Happy"
output.Range("F5").Value2 = "OK"


workbook.SaveAs(str(datetime.today()).split(" ")[0].replace("-",""))


workbook.Sheets["iris"].Activate()
        
#adjust the zoom of window (different from sheet)
workbook.Windows(1).Zoom = 120






workbook.Windows(1).FreezePanes = False
workbook.Windows(1).SplitColumn = 4
workbook.Windows(1).SplitRow = 2
workbook.Windows(1).FreezePanes = True



