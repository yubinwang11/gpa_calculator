import xlrd

def calcu(mark):  

    if 97 <= mark <= 100:
        return 4.0

    elif 93 <= mark <= 96:
        return 3.94

    elif 90 <= mark <= 92:
        return 3.85

    elif 87 <= mark <= 89:
        return 3.73

    elif 83 <= mark <= 86:
        return 3.55    

    elif 80 <= mark <= 82:
        return 3.32

    elif 77 <= mark <= 79:
        return 3.09

    elif 73 <= mark <= 76:
        return 2.78

    elif 70 <= mark <= 72:
        return 2.42

    elif 67 <= mark <= 69:
        return 2.08

    elif 63 <= mark <= 66:
        return 1.63

    elif 60 <= mark <= 62:
        return 1.15

    elif 0 <= mark < 60:
        return 0.0

sheet = xlrd.open_workbook(r'.\GPA_westlake.xlsx')
data = sheet.sheet_by_name('sheet')
credits, marks = data.col_values(0), data.col_values(1)
sum_credit = 0
sum_gpa = 0

for i in range(len(credits)):
    credit = credits[i]    
    gpa = calcu(marks[i]) 

    sum_credit += credit
    sum_gpa += credit * gpa

gpa = sum_gpa / sum_credit
print(gpa)


