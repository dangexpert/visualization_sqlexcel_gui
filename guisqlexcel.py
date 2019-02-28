import PySimpleGUI as sg
import pyodbc
import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
import sys
'exec(%matplotlib inline)'

# GUI Fields
name = [[sg.Text('Enter server name, database, query: ')],
         [sg.Text('Server', size=(15,1)), sg.InputCombo(['N/A','Insert Other Values Needed'])],
          [sg.Text('Database', size=(15, 1)), sg.InputCombo(['N/A','Insert Other Values Needed'])],      
          [sg.Text('SQL Query', size=(15, 1)), sg.InputText()],  
          [sg.Text('File Name', size=(15, 1)), sg.InputText()],
          [sg.Text('Visualize Data', size=(15,1)), sg.InputCombo(['Yes', 'No'])], #set Yes as default
          [sg.Text('Data Type', size=(15,1)), sg.InputCombo(['N/A', 'Insert Other Values Needed'])],
          [sg.Text('Graph Title', size=(15, 1)), sg.InputText()],
          [sg.Text('X-Axis', size=(15, 1)), sg.InputText()],      
          [sg.Text('Y-Axis', size=(15, 1)), sg.InputText()],       
          [sg.Submit(), sg.Cancel()]]

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
data1 = pd.read_excel(r"C:\\Users\\insert path directory" + values[3] + ".xlsx")

four = str(values[4]) #yes or no graphing
five = values[5] #data type
six = values[6] #title
seven = values[7] #x value
eight = values[8] #y value
nine = values[9] 
ten = values[10] 


if four == 'Yes': 
    if five == "Database": 
        fig, ax = plt.subplots(1,1)
        data1.plot(x = seven, y = eight, label = six, ax=ax)
        ax.set(xlabel = nine, ylabel = ten) 
        plt.title(six) 
        plt.show
    elif five == "Pool Level" :
        sys.exit()
    elif five == "Loan Level" :
        sys.exit() 
elif four == ("No"):
    sys.exit() 


