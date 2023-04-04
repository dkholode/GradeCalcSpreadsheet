import xlsxwriter as x

def getInt(prompt):
    while True:
        try:
            return int(input(prompt))
        except ValueError:
            print('Invalid input! Enter an Integer!')
        
def section(CatCount):
    print('You have', CatCount, 'categories. Let us customize each one!')
    Amt = []
    Names = []
    Points = []
    for i in range(CatCount):
        print('\n')
        Names.append(input("What is the naming scheme for these assignments? (Ex: HW #, Exam #, etc) "))
        Amt.append(getInt("How many of this assignment type are there? "))
        Points.append(input("How much is each assignment worth? (As a total of %) "))
    return Amt, Names, Points

def assemble(CatCount, Amt, Names, Points):
    worksheet.set_column(0, 0, 17.75)
    worksheet.write('A1', 'Current Percentage')
    worksheet.write('A4', 'Assignment Name')
    worksheet.write('B4', 'Score')
    worksheet.write('C4', 'Weight')
    format = workbook.add_format({'num_format': '0.00%'})
    worksheet.write_formula('A2', '{=SUM(B5:B500*C5:C500)/100}', format)

    row = 4
    col = 0
    for i in range(CatCount):
        AssignmentNum = list(range(int(Amt[i])))
        for x in range(int(Amt[i])):
            worksheet.write(row, col, Names[i] + " " + str(AssignmentNum[x]+1))
            row+=1
    row = 4
    col = 2
    for i in range(CatCount):
        for x in range(int(Amt[i])):
            worksheet.write(row, col, float(Points[i])/100,format)
            row+=1

if __name__ == "__main__":
    Course = input("What course is this for? ")
    workbook = x.Workbook(Course+'.xlsx')
    worksheet = workbook.add_worksheet()
    CatCount = getInt("How many categories does your class have? (HW, Exams, Quizzes, etc)? ")
    Amt, Names, Points = section(CatCount)
    assemble(CatCount, Amt, Names, Points)
    workbook.close()