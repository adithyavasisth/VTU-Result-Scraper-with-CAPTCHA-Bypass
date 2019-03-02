def grade(m):
    grade_point = 0
    grade_letter = ''
    if m >= 90:
        grade_letter += ',S+,'
        grade_point = 10
    elif m >= 80:
        grade_letter += ',S,'
        grade_point = 9
    elif m >= 70:
        grade_letter += ',A,'
        grade_point = 8
    elif m >= 60:
        grade_letter += ',B,'
        grade_point = 7
    elif m >= 50:
        grade_letter += ',C,'
        grade_point = 6
    elif m >= 45:
        grade_letter += ',D,'
        grade_point = 5
    elif m >= 40:
        grade_letter += ',E,'
        grade_point = 4
    elif m >= 0:
        grade_letter += ',F,'
        grade_point = 0
    elif m == -1:
        grade_letter += ',Ab,'
        grade_point = 0
    return grade_letter, grade_point


def calc(sub, subj, count1, count2, count3, record):
    c = count1 * 4 + count3 * 3 + count2 * 2
    cp = 0
    for i in range(0, count1):
        record += str(subj.pop(0))
        st, g = grade(sub.pop(0))
        record += st
        cp += g * 4
    if count3 != 0:
        for i in range(0, count3):
            record += str(subj.pop(0))
            st, g = grade(sub.pop(0))
            record += st
            cp += g * 3
    for i in range(0, count2):
        record += str(subj.pop(0))
        st, g = grade(sub.pop(0))
        record += st
        cp += g * 2
    gp = str(round((cp / c), 2))
    return record, gp


def gpa2(year, branch, low, high, sem, cycle):
    if int(sem) == 1 or int(sem) == 2:
        count1 = 5
        count2 = 2
        count3 = 0
    elif int(sem) == 5 or int(sem) == 6:
        count1 = 4
        count2 = 2
        count3 = 2
    elif int(sem) == 7 or int(sem) == 8:
        count1 = 3
        count2 = 3
        count3 = 2

    else:
        count1 = 6
        count2 = 2
        count3 = 0
    if cycle == 'P':
        marks_code = 36
        gpa_col = 9
    else:
        marks_code = 41
        gpa_col = 10
    pth = 'ExcelFiles2/'
    import pandas as pd
    df = pd.read_csv("test2.txt", sep=",")
    sub, subj = [], []
    with open('gpa2.txt', 'w+') as f:
        for i in range(0, df.shape[0]):
            record = ''
            string = ''
            record += df.iloc[i, 0] + ',' + df.iloc[i, 1] + ','
            for j in range(5, marks_code, 5):
                # record += df.iloc[i, j-3) + ','
                if int(sem) == 1 and cycle == 'C' and j == 35:
                    string = df.iloc[i, j - 3] + ',' + df.iloc[i, j + 1]
                elif df.iloc[i, j + 1] == 'P':
                    subj.append(df.iloc[i, j - 3])
                    sub.append(int(df.iloc[i, j]))
                elif df.iloc[i, j + 1] == 'A':
                    subj.append(df.iloc[i, j - 3])
                    sub.append(-1)
                else:
                    subj.append(df.iloc[i, j - 3])
                    sub.append(0)
            record, sgpa = calc(sub, subj, count1, count2, count3, record)
            percent = str(round((float(sgpa) - 0.750) * 10, 2))
            if cycle != 'C':
                record += sgpa + ',' + percent + ','
            else:
                record += string + ',' + sgpa + ',' + percent + ','
            print(record, end='\t')
            print('\n')
            f.write(record + '\n')
    f.close()
    cols = pd.read_csv("gpa2.txt").columns
    r = [0, 1, -3, -2]
    df1 = pd.read_csv("gpa2.txt", sep=",", usecols=cols[r])
    df1.to_csv("gpa2r.txt", sep=",", index=False)
    df2 = pd.read_csv("gpa2r.txt", sep=",", header=None)
    df2.columns = ['USN', 'Name', 'GPA', 'Percentage']
    try:
        df2 = df2.sort_values(by=[df2.columns[2]], ascending=False)
    except AttributeError:
        print(" ")
    writer = pd.ExcelWriter(pth + year + branch + str(low) + '-' + str(high - 1) + 'rank2.xls')
    df2.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()
