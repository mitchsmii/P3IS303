# Justin Maxwell - Carson Hendrix - Prof. Anderson

# DESCRIPTION: Loops through the data in the excel file and moves it to the corresponding sheet
# So it will split the last name, first name, and ID into a list, then appends the grade to the list as well
# Then puts it on the sheet made for that class


#
from openpyxl import Workbook

def organizeData(iCount):
    data = currSheet["B" + str(iCount)]
    data.split('_')
    lName = data[0]
    fName = data[1]
    idNum = data[2]
    return lName, fName, idNum

myworkbook = Workbook()

currSheet = myworkbook.active

iCount = 2
while currSheet["A" + str(iCount)] == "Algebra":
    lName, fName, idNum = organizeData(iCount)
    iCount += 1
    currSheet["A" + str(iCount)] = lName
    currSheet["B" + str(iCount)] = fName
    currSheet["C" + str(iCount)] = idNum

myworkbook.save(filename= "Poorly_Organized_Data_1 (1).xlsx")
