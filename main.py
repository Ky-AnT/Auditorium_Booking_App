from tkinter import *
from tkmacosx import Button
import pandas as pd
import pyexcelerate as px
import xlwings as xw


app = xw.App(visible=False)

try:
    workbook = app.books.open('spreadsheet.xlsx')
except:
    print("spreadsheet.xlsx does not exist!!! Exiting program...")
    exit()

sheet = workbook.sheets['Sheet1']

root = Tk()

root.title("Auditorium Bookings")

buttonDict = {}
alphabet = "ABCDEFGHJKLMNOPQ"
matrix = []
tempList = []
def reloadData():
    global alphabet, matrix, tempList
    for n in range(8):
        for i in range(16):
            tempList.append(sheet.range(f"{alphabet[i]}{str(n+1)}").value)
        matrix.append(tempList)
        tempList = []
    return matrix
currentData = reloadData()
def toggleBooking(id):
    global correspondingDict
    buttonDict[id].configure(bg=correspondingDict[buttonDict[id].cget("bg")])
    cell = buttonDict[id].cget("text")
    print(cell)
    print(buttonDict[id])
    sheet.range(cell).value = bookedDict[buttonDict[id].cget("bg")]
    workbook.save("spreadsheet.xlsx")
    createTable(refreshTableFunction())

def refreshTableFunction():
    global workbook, sheet
    #resetFunction()
    #workbook = load_workbook(filename="spreadsheet.xlsx", data_only=True)
    #sheet = workbook.active
    print(sheet["C11"].value)

    #returnTable = [("Gold",sheet["C11"].value,sheet["E11"].value),("Silver",sheet["C12"].value,sheet["E12"].value),("Bronze",sheet["C13"].value,sheet["E13"].value)]
    # evaluate_formula(sheet["C11"])
    returnTable = [
        ("Class","No. of Tickets","Cost"),
        ("Gold", sheet.range("C11").value, sheet.range("E11").value),
        ("Silver", sheet.range("C12").value, sheet.range("E12").value),
        ("Bronze", sheet.range("C13").value, sheet.range("E13").value),
        ("Total", sheet.range("C14").value, sheet.range("E14").value)
    ]
    #print(returnTable)
    return returnTable
def secondTableRefresh():
    global workbook, sheet
    #resetFunction()
    #workbook = load_workbook(filename="spreadsheet.xlsx", data_only=True)
    #sheet = workbook.active
    print(sheet["C11"].value)

    print(sheet.range("I11").value)

    returnTable = [
        ("Refreshments", "Cost per Item", "Quantity", "Cost"),
        (sheet.range("H11").value, sheet.range("J11").value, int(sheet.range("L11").value), sheet.range("N11").value),
        (sheet.range("H12").value, sheet.range("J12").value, int(sheet.range("L12").value), sheet.range("N12").value),
        (sheet.range("H13").value, sheet.range("J13").value, int(sheet.range("L13").value), sheet.range("N13").value),
        (sheet.range("H14").value, sheet.range("J14").value, int(sheet.range("L14").value), sheet.range("N14").value),
        (sheet.range("H15").value, sheet.range("J15").value, int(sheet.range("L15").value), sheet.range("N15").value),
        (sheet.range("H16").value, sheet.range("J16").value, int(sheet.range("L16").value), sheet.range("N16").value),
        (sheet.range("H17").value, sheet.range("J17").value, int(sheet.range("L17").value), sheet.range("N17").value),
        (sheet.range("H18").value, sheet.range("J18").value, int(sheet.range("L18").value), sheet.range("N18").value)
    ]

    print(returnTable)
    return returnTable

def evaluate_formula(cell):
    formula = cell.value
    print(formula)
    if formula is None:
        print("It is None")
        print(cell.value)
        return cell.value
    
    # Create a temporary DataFrame with the formula in a single cell
    df = pd.DataFrame([[formula]], columns=[cell.coordinate])
    evaluated_value = df[cell.coordinate].values[0]
    
    return evaluated_value


'''
def letter_to_number(column_letter):
    return xw.utils.column_index_from_string(column_letter)

# Example usage
column_letter = 'A'
column_number = letter_to_number(column_letter)
print("Column number:", column_number)
'''

'''
def resetFunction():
    global sheet
    sheet['C11'] = '=COUNTIF($A$7:$H$8,"Booked")+COUNTIF($J$7:$Q$8,"Booked")'
    sheet['E11'] = '=PRODUCT(C11,20)'
    sheet['C12'] = '=COUNTIF($A$5:$H$6,"Booked")+COUNTIF($J$5:$Q$6,"Booked")'
    sheet['E12'] = '=PRODUCT(C12,15)'
    sheet['C13'] = '=COUNTIF($A$1:$H$4,"Booked")+COUNTIF($J$1:$Q$4,"Booked")'
    sheet['E13'] = '=PRODUCT(C13,10)'
    workbook.save(filename="spreadsheet.xlsx")
    #workbook.close()
'''
tableEntries = []

def createTable(table):
    global tableEntries
    count = 0
    for i in range(len(table)):
        for j in range(len(table[0])):
            if len(tableEntries) <= count:
                # Create a new Entry widget and append it to the tableEntries list
                entry = Entry(root, width=8, fg='blue', font=('Arial', 16, 'bold'))
                entry.grid(row=i + 9, column=j)
                entry.insert(END, table[i][j])
                tableEntries.append(entry)
            else:
                # Use the existing Entry widget from tableEntries
                entry = tableEntries[count]
                entry.grid(row=i + 9, column=j)
                entry.delete(0, END)
                entry.insert(END, table[i][j])
            count += 1
refresherTableEntries = []
otherTableDict = {}
def createRefreshersTable(table):
    global refresherTableEntries
    count = 0
    for i in range(len(table)):
        for j in range(len(table[0])):
            if len(refresherTableEntries) <= count:
                # Create a new Entry widget and append it to the tableEntries list
                if j == 2 and i != 0:
                    otherTableDict[count] = Spinbox(root, width=8, from_=0, to=100,command=lambda id=count: setSpinbox(id), fg='blue', font=('Arial', 16, 'bold'))
                    otherTableDict[count].grid(row=i + 9, column=j+3)
                    otherTableDict[count].insert(END, table[i][j])
                    refresherTableEntries.append(otherTableDict)
                else:
                    entry = Entry(root, width=8, fg='blue', font=('Arial', 16, 'bold'))
                    entry.grid(row=i + 9, column=j+3)
                    entry.insert(END, table[i][j])
                    refresherTableEntries.append(entry)
                
            else:
                # Use the existing Entry widget from tableEntries
                entry = tableEntries[count]
                entry.grid(row=i + 9, column=j+3)
                entry.delete(0, END)
                entry.insert(END, table[i][j])
            count += 1
def setSpinbox(row):
    global refresherTableEntries
    rows = [11,12,13,14,15,16,17,18]
    rowNumber = str(rows[int((row-6)/4)])
    indexed = (int(rowNumber)-10)*4+3
    #print(rowNumber)
    sheet.range("L"+str(rows[int((row-6)/4)])).value = otherTableDict[row].get()
    #refresherTableEntries[row] = sheet.range("N"+str(rows[int((row-6)/4)])).value
    changeVal = sheet.range("N"+str(rows[int((row-6)/4)])).value
    #.set("hello")
    refresherTableEntries[indexed].delete(0, END)
    refresherTableEntries[indexed].insert(0, changeVal)
    #createRefreshersTable(secondTableRefresh())
    


createTable(refreshTableFunction())
createRefreshersTable(secondTableRefresh())

counter = 0
trueFalse = {
    True:"blue",
    False:"white"
}

correspondingDict = {
    "blue":"white",
    "white":"blue"
}

bookedDict = {
    "blue":"Booked",
    "white":""
}


for n in range(8):
    for i in range(16):
        print(trueFalse[matrix[n][i]=='Booked'])
        if alphabet[i] == 'Q':
            print('q')
        buttonDict[counter]=Button(root, width=80, text=f"{alphabet[i]}{str(n+1)}", command=lambda id=counter: toggleBooking(id), background=trueFalse[matrix[n][i]=='Booked'])
        buttonDict[counter].grid(row=n+1,column=i)
    
        #buttonDict[counter].pack()
        
        counter += 1
def on_closing():
    workbook.close()
    app.quit()
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()
