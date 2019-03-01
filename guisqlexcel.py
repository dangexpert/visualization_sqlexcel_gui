import PySimpleGUI as sg
import pyodbc
import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
import sys
'exec(%matplotlib inline)' #The inline backend is only available in Jupyter Notebook and the Jupyter QtConsole. (Optional) 

# GUI Fields
name = [[sg.Text('Enter server name, database, query: ')],
         [sg.Text('Server', size=(15,1)), sg.InputCombo(['Insert Server Name'])],
          [sg.Text('Database', size=(15, 1)), sg.InputCombo(['Insert Database Name'])],      
          [sg.Text('SQL Query', size=(15, 1)), sg.InputText()],  
          [sg.Text('File Name', size=(15, 1)), sg.InputText()],
          [sg.Text('Data Visualization', size=(15, 1)), sg.InputCombo(['Yes', 'No'])], 
          [sg.Text('Graph Title', size=(15, 1)), sg.InputText()],
          [sg.Text('X-Axis', size=(15, 1)), sg.InputText()],      
          [sg.Text('Y-Axis', size=(15, 1)), sg.InputText()],  
          [sg.Text('X-Label', size=(15, 1)), sg.InputText()],      
          [sg.Text('Y-Label', size=(15, 1)), sg.InputText()],            
          
# Graphing GUI Fields          
          [sg.Text('Line', size=(15, 1)), sg.InputCombo(['Yes', 'No'])],
          [sg.Text('Histogram', size=(15, 1)), sg.InputCombo(['Yes', 'No'])],
          [sg.Text('Box Plot', size=(15, 1)), sg.InputCombo(['Yes', 'No'])],
          #add any other graphing methods as needed
          
#Title 
form = sg.Window("SQL-Excel Converter/Visualization").Layout(name)      

button, values = form.Read()
form.Close()
#------------------------------------------------------------------------------------------------------------
# SQL Query 
query = pyodbc.connect('Driver={SQL Server};' 'Server=' + str(values[0]) + ';' + 'Database=' + str(values[1]) + ';' + 'Trusted_Connection=yes;')
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------
# Sets SQL query value 
sql = values[2]
#Read the SQL query for the driver, server, and database connection
data = pd.read_sql(sql, query)
#Formats table into dataframe using the read_sql function 
Data = pd.DataFrame(data)
print(Data)
#----------------------------------------------------------------------------------------------------------------------------------------------------------------
#Exports database table to Excel format 
newFile = ("C:\\Users\\insert path directory here" + values[3] + ".xlsx")
export_excel = Data.to_excel(newFile, index=None, header=True) 
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------
# Formats Excel sheet in correct width 
wb = openpyxl.load_workbook(filename = newFile) #uses the openpyxl module by loading the workbook 
worksheet = wb.active #activates the worksheet -- need to make sure to activate wb.save() at the end if changing excel 

for col in worksheet.columns:
    max_length = 0
    column = col[0].column

    for cell in col:
        try: 
#based on the value of the cell, it makes sure it equals the max_length 
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        worksheet.column_dimensions[column].width = adjusted_width

wb.save(newFile)

print("File has been created!")
#---------------------------------------------------------------------------------------------------------
read = pd.read_excel(r"C:\\Users\\insert path directory" + values[3] + ".xlsx")

four = str(values[4]) #yes or no graphing
five = values[5] #title
six = values[6] #x value
seven = values[7] #y value
eight = values[8] 
nine = values[9] 
chart = values[10] 
chart1= values[11] 
chart2 = values[12] 

# Can add more charts using https://python-graph-gallery.com/ depending on data set
if four == 'Yes': 
    if chart == 'Yes': #line plot
        fig, ax = plt.subplots(1,1)
        read.plot(x = seven, y = eight, label = six, ax=ax)
        ax.set(xlabel = six, ylabel = nine) 
        plt.title(five) 
        plt.show
    elif  chart1 == 'Yes': #histogram
        sys.exit()
    elif chart2 == 'Yes': #box plot
        read.boxplot(column = six)
        read.boxplot(column = six, by = seven)
        read[five].hist(bins=50)
elif four == ("No"):
    sys.exit() 

