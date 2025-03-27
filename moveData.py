# Justin Maxwell - Carson Hendrix - Prof. Anderson

# DESCRIPTION: Loops through the data in the excel file and moves it to the corresponding sheet
# So it will split the last name, first name, and ID into a list, then appends the grade to the list as well
# Then puts it on the sheet made for that class


# 

from openpyxl import load_workbook

def organizeData(iCount):
    currSheet = myworkbook.active
    data = currSheet["B" + str(iCount)].value
    lName, fName, idNum = data.split('_')
    currSheet = myworkbook[currSheet["A" + str(iCount)].value]
    currSheet["A" + str(iCount)] = lName
    currSheet["B" + str(iCount)] = fName
    currSheet["C" + str(iCount)] = idNum

myworkbook = load_workbook("Poorly_Organized_Data_1.xlsx")
currSheet = myworkbook.active

value = currSheet["A2"].value
print(value)
iCount = 2

myworkbook.create_sheet("Algebra")
while currSheet["A" + str(iCount)].value == "Algebra":
    organizeData(iCount)
    currSheet = myworkbook["Algebra"]
    iCount += 1
myworkbook.create_sheet("Trigonometry")
while currSheet["A" + str(iCount)].value == "Trigonometry":
    organizeData(iCount)
    iCount += 1
myworkbook.create_sheet("Geometry")
while currSheet["A" + str(iCount)].value == "Geometry":
    organizeData(iCount)
    iCount += 1
myworkbook.create_sheet("Calculus")
while currSheet["A" + str(iCount)].value == "Calculus":
    organizeData(iCount)
    iCount += 1
myworkbook.create_sheet("Statistics")
while currSheet["A" + str(iCount)].value == "Statistics":
    organizeData(iCount)
    iCount += 1

myworkbook.save(filename= "Poorly_Organized_Data_1.xlsx")