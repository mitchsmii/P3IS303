# Justin Maxwell - Carson Hendrix - Prof. Anderson

# DESCRIPTION: Loops through the data in the excel file and moves it to the corresponding sheet
# So it will split the last name, first name, and ID into a list, then appends the grade to the list as well
# Then puts it on the sheet made for that class


#
from openpyxl import Workbook

def organizeData(iCount):
    data = currSheet["B" + iCount]
    data.split('_')
    lName = data[0]
    fName = data[1]
    idNum = data[2]
    currSheet["F" + str(iCount)] = lName
    currSheet["G" + str(iCount)] = fName
    currSheet["H" + str(iCount)] = idNum

myworkbook = Workbook()

currSheet = myworkbook.active

iCount = 2
while currSheet["A" + str(iCount)] == "Algebra":
    organizeData(iCount)
    iCount += 1

myworkbook.save(filename= "Poorly_Organized_Data_1 (1).xlsx")
