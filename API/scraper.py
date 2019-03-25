# coding=utf-8
import os
import re
import sys
import warnings

import cv2
import pytesseract
from PIL import Image
from bs4 import BeautifulSoup
from selenium import webdriver

# Tesseract Path
pytesseract.pytesseract.tesseract_cmd = r'.\Tesseract-OCR\tesseract'

if not sys.warnoptions:
    warnings.simplefilter("ignore")


def scrape(college, year, branch, low, high, semc):
    # Input for Branch and USNs
    # college = input("Enter the college code\n").upper()
    # year = input('Enter the year\n')
    # branch = input('Please enter the branch\n').upper()
    # low = int(input('Enter starting USN\n'))
    low = int(low)
    high = int(high) + 1
    if low >= 400:
        dip = 'Y'
    else:
        dip = 'N'
    # increment last USN to aid looping
    # high = int(input('Enter last USN\n')) + 1
    # semc = input('Enter the Semester\n')
    cycle = 'N'
    if low >= 400:
        dip = 'Y'
    else:
        dip = 'N'

    subcode = 52
    iloop = 8
    if semc == '1' or semc == '2':
        cycle = input('Enter the Cycle\n').upper()
        if cycle == 'P':
            iloop = 7
            subcode = 46
    if semc == '3' or semc == '4':
        iloop = 9
        subcode = 58
        if dip == 'Y':
            iloop = 10
            subcode = 64

    # Opens file for storing data
    with open('marks.txt', 'w+') as f:
        c = 0
        pf = ''
        driver = webdriver.Chrome('.\chromedriver_win32\chromedriver.exe')
        # For Loop to loop through all USNs
        for u in range(low, high):
            redo = True
            while redo:
                redo = False
                # IF condition to concatenate USN
                if u < 10:
                    usn = '1' + college + year + branch + '00' + str(u)
                elif u < 100:
                    usn = '1' + college + year + branch + '0' + str(u)
                else:
                    usn = '1' + college + year + branch + str(u)

                # Opens the vtu result login page, gets the usn and opens the result page
                driver.get('http://results.vtu.ac.in/resultsvitavicbcs_19/index.php')
                driver.save_screenshot('python_org.png')
                img = cv2.imread("python_org.png")
                crop_img = img[467:508, 667:885]
                cv2.imwrite('cap.png', crop_img)
                cv2.waitKey(0)
                tex = pytesseract.image_to_string(Image.open('cap.png'))
                tex = tex.strip(',')
                tex = tex.strip(' ')
                try:
                    captcha = int(tex)
                except:
                    print("Invalid CAPTCHA Detected.")
                    redo = True
                    continue
                us = driver.find_element_by_name("lns")
                cap = driver.find_element_by_name("captchacode")
                us.send_keys(usn)
                cap.send_keys(captcha)
                driver.find_element_by_id("submit").click()
                try:
                    soup = BeautifulSoup(driver.page_source)
                except:
                    # Catches all types of errors like Invalid Captcha and Invalid USN
                    alert = driver.switch_to.alert
                    if alert.text == "University Seat Number is not available or Invalid..!":
                        print("No results for : " + usn + "\n")
                        alert.accept()
                        continue
                    elif alert.text == "Invalid captcha code !!!":
                        print("Invalid CAPTCHA Detected for USN : " + usn + "\n")
                        alert.accept()
                        redo = True
                        continue

                # Finds all the table elements and stores in array tds
                tds = soup.find_all('td')
                ths = soup.find_all('th')
                divs = soup.find_all('div', attrs={'class': 'col-md-12'})
                divCell = soup.find_all('div', attrs={'class': 'divTableCell'})

                try:
                    sem = divs[5].div.text
                    sem = sem.strip('Semester : ')
                except AttributeError:
                    print("INVALID USN/ INCOMPATIBLE DATA : " + usn + "\n")

                # IF condition to check invalid page opener
                if tds[
                    0].text != 'University Seat Number ' or sem != semc:  # To check for Diploma and Number of students
                    print("INVALID USN/ INCOMPATIBLE DATA : " + usn + "\n")
                    continue

                record = ''
                # tds[1] holds USN number
                record += re.sub('[!@#$:]', '', tds[1].text)
                record += ','
                # tds[3] holds the name
                record += re.sub('[!@#$:]', '', tds[3].text)
                record += ','

                sortList1 = []
                for i in range(6, subcode, 6):
                    if (divCell[i].text[-3:]).isdigit():
                        sortList1.append(divCell[i].text[-3:])
                    else:
                        sortList1.append(divCell[i].text[-2:])
                sortList1.sort()

                ilist = []
                for i in range(0, iloop):
                    for j in range(6, subcode, 6):
                        if (divCell[j].text[-3:]).isdigit():
                            if divCell[j].text[-3:] == sortList1[i] and j not in ilist:
                                ilist.append(j)
                        else:
                            if divCell[j].text[-2:] == sortList1[i] and j not in ilist:
                                ilist.append(j)

                # Strips extra garbage from the retrieved USN text
                print(record, end='\t')
                # Loop that goes from 8 to 51 in steps of 6 because starting from 8, in steps of 6
                try:
                    for l in ilist:
                        # Checks if string has number
                        for j in range(l, l + 6):
                            if j == l + 1:
                                continue
                            else:
                                char = divCell[j].text
                                if char.isdigit():
                                    record = record + str(int(char)) + ','
                                else:
                                    record = record + char + ','
                                print(divCell[j].text, end='\t\t')
                                if j == l + 5:
                                    pf = pf + divCell[j].text + ','
                    f.write(record + '\n')
                    print('\n')
                except IndexError:
                    pass

    driver.quit()

    # Writing from text files to Excel Files.
    import xlwt

    book = xlwt.Workbook()
    ws = book.add_sheet('Sheet1')  # Add a sheet
    f = open('marks.txt', 'r+')

    alignment = xlwt.Alignment()  # Create Alignment
    font = xlwt.Font()  # Create the Font
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    font.bold = True
    style = xlwt.XFStyle()  # Create Style
    # style.alignment = alignment  # Add Alignment to Style

    data = f.readlines()  # read all lines at once
    for i in range(len(data)):
        row = data[i].split(',')
        # This will return a line of string data, you may need to convert to other formats depending on your use case
        for j in range(len(row)):
            if row[j].isdigit():
                ws.write(i, j, int(row[j]), style)  # Write to cell i, j
            else:
                ws.write(i, j, row[j], style)

    # Saving the Excel Files onto the designated Path
    pth = 'ExcelFiles/'
    if not os.path.exists(pth):
        os.makedirs(pth)
    if dip == 'N':
        book.save(pth + '1' + college + year + branch + str(low) + '-' + str(high - 1) + '.xls')
    else:
        book.save(pth + '1' + college + year + branch + str(low) + '-' + str(high - 1) + 'DIP.xls')
    f.close()

    files = ['cap.png', 'python_org.png', 'marks.txt']
    for file in files:
        try:
            os.remove(file)
        except:
            pass

    if dip != 'Y':
        from sgpa import gpa

        gpa(college, year, branch, low, high, semc, cycle)



