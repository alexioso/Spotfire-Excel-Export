
#INPORTS
#filewriters and IO handlers
from System.IO import StreamWriter, FileMode
#spotfire objects
from Spotfire.Dxp.Application.Visuals import CrossTablePlot, TablePlot
from  Spotfire.Dxp.Data.Export import DataWriterTypeIdentifiers
#creates tempfiles to count rows/cols
import tempfile
import os
#NET Common Language Runtime scripting enabled
import clr
#clr allows us to directly interact with open Excel applications
clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
#import Excel applications
from Microsoft.Office.Interop import Excel
#csv and datetime data handling
import csv
import datetime

clr.AddReference('System.Drawing')
from System.Drawing import Color, ColorTranslator
   
print(Application.DocumentMetadata.LoadedFromFileName)

#translates RGC from Spotfire to format Excel can understand
def rgbForExcel(r, g, b):
  return ColorTranslator.ToOle(Color.FromArgb(r, g, b))

#returns RGB color values for number based on linear gradient rule
def gradient_value(rgb1, rgb2, max_value, min_value, value):
    rgb = [0,0,0]
    for i in range(3):
        rgb[i] = rgb1[i] + (rgb2[i] - rgb1[i])*((value-min_value)/(max_value - min_value))
    print(rgb)
    return(rgb)

#parses Spotfire color rule object to determine values needed to compute gradient_value() and return rgb
def cont_color_rule_value(color_rule, value, max_value, min_value):
  
    
    max_val = float("-inf")
    max_index = -1

    temp_list = [] #this will help find top color gradient
    
    #this loop will find bottom color gradient
    for i in range(len(color_rule[0][1])):
        v = color_rule[0][1][i]
        temp_list.append(v - value)
        if value >= v and v >= max_val:
            max_val = v
            max_index = i
    rgb_color1 = color_rule[0][0][max_index]
    v1 = color_rule[0][1][max_index]
    
    

    top_ind = -1
    top_value = float("inf")
    #this loop will find top color gradient
    for j in range(len(temp_list)):
        if temp_list[j] > 0 and temp_list[j] < top_value:
            top_ind = j
            top_value = temp_list[j]
    rgb_color2 = color_rule[0][0][top_ind]
    v2 = color_rule[0][1][top_ind]


    if v1 == float("-inf"):
        v1 = min_value
    if v2 == float("inf"):
        v2 = max_value

    return (gradient_value(rgb_color1, rgb_color2, v2, v1, value))
       
               
#parses Spotfire segment color rule object to return rgb color for given value
def segment_color_rule_value(color_rule, value):
     
    if value is None:
        return color_rule[0][2][0]

    max_val = float("-inf")
    max_index = -1
    
    for i in range(len(color_rule[0][1])):
        v = color_rule[0][1][i]
        if value >= v and v >= max_val:
            max_val = v
            max_index = i
    rgb_color = color_rule[0][0][max_index]   
    return(rgb_color)                 

#deterined rgb color of value based on Spotfire categorical color rule object
def cat_color_rule_value(color_rule, value):

    if value is None:
        return color_rule[0][1][0]

    else:
        return color_rule[0][0][value]
                


#extracts RGB from Spotfire color object to list of three values
def color_to_rgb(color):
    rgb = []
    for i in color.split("=")[1:]:
        rgb.append(int(i.replace(",","").replace("]","").replace("R","").replace("G","").replace("B","")))
    return rgb[1:]

# setup stream for export text function
filename = str(datetime.datetime.today()).split(" ")[0].replace("-","") + " " + str(table.Title) + ".csv"
#open a temp file in Spotfire default temp directory for writing
tmp = os.path.join(tempfile.gettempdir(), filename)
stream = StreamWriter(tmp)

#export text

try:
    #cast input table as Spotfire TablePlot object to variable t
    t = table.As[TablePlot]()
    #export table to temp file
    t.ExportText(stream)
except Exception as e:
    print(e)
finally:
    stream.Close()


#save color rules to dictionary objects with custom formats (see below comments)
segment_color_rules = {}
cat_color_rules = {} 
cont_color_rules = {}
#extract the columns with colorings in table t
coloring = t.Colorings
for i in coloring:
    #extract the color rule for each coloring
    colorRule = i.Item[0]
    print(i.DisplayName)
    #if empty color value present print it out
    print("Empty Color: " + str(color_to_rgb(str(i.EmptyColor))))
    #handles the categorical color rules
    if 'Categorical' in  str(type(colorRule)):
        #this list holds categorical color rules in this format: [{cat1:col1, cat2:col2, ... , catN:colN},[emptyColor]]
        cat_color_rules[i.DisplayName] = [] 
        counter = 0
        category_colors = {}
        for ck in colorRule.GetExplicitCategories():
            category_colors[str(ck)] = color_to_rgb(str(colorRule.Item[ck]))
                    
        empty_col = color_to_rgb(str(i.EmptyColor))
        
        cat_color_rules[i.DisplayName].append([category_colors,[empty_col]])
        
    #handles the segments color rules
    elif 'Segments' in str(colorRule.IntervalMode):
        #this list holds segment color rules in this format: [[bp1,bp2,bp3,bp4],[v1,v2,v3,v4],[emptyColor]]
        segment_color_rules[i.DisplayName] = [] 
        breakpoints = colorRule.Breakpoints
        bps = []
        vals = []
        emptys = []
        emptys.append(color_to_rgb(str(i.EmptyColor)))
        for bp in breakpoints:
            if str(bp.Value.Type) == "Literal" or str(bp.Value.Type) == "MinValue":
                if bp.Value.Value is None:
                    vals.append(float("-inf"))
                else:
    			          vals.append(bp.Value.Value)
                b1 = color_to_rgb(str(bp.Color))
                bps.append(b1)

        segment_color_rules[i.DisplayName].append([bps,vals,emptys])
    #handles conitnuous color rules
    elif 'Continuous' in str(type(colorRule)):
        #this holds continuous color rules in this format: [[minBp,maxBp,bp1,bp2,...],[-inf,inf,v1,v2],[emptyColor]]
        cont_color_rules[i.DisplayName] = [] 
        breakpoints = colorRule.Breakpoints
        bps = []
        vals = []
        for bp in breakpoints:
            if str(bp.Value.Type) == "MinValue":
                vals.append(float("-inf"))
            elif str(bp.Value.Type) == "MaxValue":
                vals.append(float("inf"))
            else:
                vals.append(bp.Value.Value)
            bps.append(color_to_rgb(str(bp.Color)))          

        empty_col = color_to_rgb(str(i.EmptyColor))
        cont_color_rules[i.DisplayName].append([bps,vals,[empty_col]])
        

print(segment_color_rules)
print(cat_color_rules)
print(cont_color_rules)

#use dictionary object to store colors to values
df = {}

#read table from temp file
ff = open(tmp, 'r').readlines()
#count rows and cols in table since that's not stored anywhere
row_count = 0
#with open(tmp,"wb") as file_:
lastCols = []
for line in ff:
    lastCols.append(len(line.split("\t"))_
    #line = line.replace(",","|")
    #file_.write(line.replace("\t",","))
    row_count += 1

lastCol = max(lastCols)

reader = list(csv.reader(open(tmp, "r"), delimiter='\t'))
with open(tmp,'wb') as outfile:
    writer = csv.writer(outfile, delimiter=',')
    writer.writerows(row for row in reader)




ex = Excel.ApplicationClass()   
ex.Visible = True
ex.DisplayAlerts = False   

workbook = ex.Workbooks.Open(tmp)

output = workbook.ActiveSheet

output_range = output.Range(output.Cells(1,1),output.Cells(1,lastCol))

col_count = 1
for h in output_range:
    header = str(h.Value2)

    #add logic for header coloring for Lucy report
    if "*blue*" in header:
        output.Cells(1, col_count).Value = header.replace("*blue* ","")
        output.Cells(1, col_count).Interior.Color = rgbForExcel(0,127,255)
    if "*orange*" in header:
        output.Cells(1, col_count).Value = header.replace("*orange* ","")
        output.Cells(1, col_count).Interior.Color = rgbForExcel(255,153,0)        
    if "*green*" in header:
        output.Cells(1, col_count).Value = header.replace("*green* ","")
        output.Cells(1, col_count).Interior.Color = rgbForExcel(112,219,147)
    if "*yellow*" in header:
        output.Cells(1, col_count).Value = header.replace("*yellow* ","")
        output.Cells(1, col_count).Interior.Color = rgbForExcel(255,255,0)
    if "*red*" in header:
        output.Cells(1, col_count).Value = header.replace("*red* ","")
        output.Cells(1, col_count).Font.Color = rgbForExcel(255,0,0)

    if header in list(cat_color_rules.keys()):
        col = col_count
        
        colName = header

        print("We have a categorical column match: " + colName +"," + str(col))
        col_range = output.Range(output.Cells(2,col),output.Cells(row_count,col))
        row_num = 2
        for entry in col_range:
            if (entry.Value2 is None):
                val = None
            else:
                val = entry.Value2
            #print(val)
            #print(color_rules[colName])
            rgb_color = cat_color_rule_value(cat_color_rules[colName],val)
            #print(rgb_color)
            if val is None and len(rgb_color) == 0:
                row_num += 1
                continue
            output.Cells(row_num, col).Interior.Color = rgbForExcel(rgb_color[0], rgb_color[1], rgb_color[2])
            row_num += 1
 
    elif header in list(segment_color_rules.keys()):
        col = col_count
        
        colName = header

        print("We have a segment column match: " + colName +"," + str(col))
        col_range = output.Range(output.Cells(2,col),output.Cells(row_count,col))
        row_num = 2
        for entry in col_range:
            if (entry.Value2 is None):
                val = None
            else:
                val = float(entry.Value2)
            #print(val)
            #print(color_rules[colName])
            rgb_color = segment_color_rule_value(segment_color_rules[colName],val)
            #print(rgb_color)
            if val is None and len(rgb_color) == 0:
                row_num += 1
                continue
            output.Cells(row_num, col).Interior.Color = rgbForExcel(rgb_color[0], rgb_color[1], rgb_color[2])
            row_num += 1

    elif header in list(cont_color_rules.keys()):
        col = col_count
        colName = header
        print("We have a continuous column match: " + colName + "," + str(col))
        col_range = output.Range(output.Cells(2,col),output.Cells(row_count,col))

        maximum = float("-inf")
        maximum_i = -1
        minimum = float("inf")
        minimum_i = -1
         
        max_min_list = []
        for entry0 in col_range:
            if entry0.Value2 is None:
                continue
            max_min_list.append(float(entry0.value2))
        print(max_min_list)
        max_value = max(max_min_list)
        min_value = min(max_min_list)
        print(max_value, min_value)

        

        row_num = 2
        for entry in col_range:
            if (entry.Value2 is None):
                val = None
            else:
                val = float(entry.Value2)
            #print(val)
            #print(color_rules[colName])
            rgb_color = cont_color_rule_value(cont_color_rules[colName],val,max_value,min_value)
            #print(rgb_color)
            if val is None and len(rgb_color) == 0:
                row_num += 1
                continue
            output.Cells(row_num, col).Interior.Color = rgbForExcel(rgb_color[0], rgb_color[1], rgb_color[2])
            row_num += 1


    col_count += 1
