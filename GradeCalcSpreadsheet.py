import xlsxwriter as x

workbook = x.Workbook('GradeCalc.xlsx')
worksheet = workbook.add_worksheet()

def terms():
    CatCount = int(input("How many categories does your class have? (HW, Exams, Quizzes, etc)? "))
    if CatCount < 0:
        print('\nPlease enter a whole positive integer\n')
        CatCount = int(input("How many categories does your class have? (HW, Exams, Quizzes, etc)? "))
    return CatCount

def section(CatCount):
    print('You have', CatCount, 'categories. We will now customize each one!')
    Amt = []
    Names = []
    Points = []
    for i in range(CatCount):
        print('\n')
        Names.append(input("What is the naming scheme for these assignments? (Ex: HW #, Exam #, etc) "))
        Amt.append(input("How many of this assignment type are there? "))
        Points.append(input("How many percent is each assignment worth? "))
    return Amt, Names, Points

def assemble(type, CatCount, Amt, Names, Points):
    worksheet.set_column(0, 0, 17.75)
    worksheet.write('A1', 'Current Percentage')
    worksheet.write('A4', 'Assignment Name')
    worksheet.write('B4', 'Weight')
    worksheet.write('C4', 'Score')

    format = workbook.add_format({'num_format': '0.00%'})

    row = 4
    col = 0
    for i in range(CatCount):
        AssignmentNum = list(range(int(Amt[i])))
        for x in range(int(Amt[i])):
            worksheet.write(row, col, Names[i] + " " + str(AssignmentNum[x]+1))
            row+=1

    row = 4
    col = 1
    for i in range(CatCount):
        for x in range(int(Amt[i])):
            worksheet.write(row, col, float(Points[i])/100,format)
            row+=1
    worksheet.write_formula('A2', '{=SUM(B5:B500*C5:C500)/100}', format)
    

CatCount = terms()
Amt, Names, Points = section(CatCount)
assemble(type, CatCount, Amt, Names, Points)

workbook.close()