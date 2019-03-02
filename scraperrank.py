# coding=utf-8

import re
import sys
import warnings

import cv2
import pytesseract
from PIL import Image
from bs4 import BeautifulSoup
from selenium import webdriver

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files (x86)\Tesseract-OCR\tesseract'

if not sys.warnoptions:
    warnings.simplefilter("ignore")

# Input for Branch and USNs

college = ["4AI", "1BG", "1CR", "1AM", "1BI"]
year = input('Enter the year\n')
branch = input('Please enter the branch\n').upper()
low = int(input('Enter starting USN\n'))
if low >= 400:
    dip = 'Y'
else:
    dip = 'N'
# increment last USN to aid looping
high = int(input('Enter last USN\n')) + 1
semc = input('Enter the Semester\n')
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
    if dip == 'Y':
        iloop = 9
        subcode = 58

# Opens file for storing data
with open('test2.txt', 'w+') as f:
    c = 0
    pf = ''
    driver = webdriver.Chrome('C:\Program Files (x86)\chromedriver_win32\chromedriver.exe')

    # For Loop to loop through all USNs
    for x in college:
        for u in range(low, high):

            # IF condition to concatenate USN
            if u < 10:
                usn = x + year + branch + '00' + str(u)
            elif u < 100:
                usn = x + year + branch + '0' + str(u)
            else:
                usn = x + year + branch + str(u)

            # opens the vtu result login page, gets the usn and opens the result page
            driver.get('http://results.vtu.ac.in/resultsvitavicbcs_19/index.php')
            driver.save_screenshot('python_org.png')
            img = cv2.imread("python_org.png")
            crop_img = img[467:508, 667:885]
            cv2.imwrite('cap.png', crop_img)
            cv2.waitKey(0)
            tex = pytesseract.image_to_string(Image.open('cap.png'))
            tex = tex.strip(',')
            tex = tex.strip(' ')
            captcha = int(tex)
            us = driver.find_element_by_name("lns")
            cap = driver.find_element_by_name("captchacode")
            us.send_keys(usn)
            cap.send_keys(captcha)
            driver.find_element_by_id("submit").click()
            try:
                soup = BeautifulSoup(driver.page_source)
            except:
                alert = driver.switch_to.alert
                alert.dismiss()
                continue
            # Finds all the table elements and stores in array tds
            tds = soup.find_all('td')
            ths = soup.find_all('th')
            divs = soup.find_all('div', attrs={'class': 'col-md-12'})
            divCell = soup.find_all('div', attrs={'class': 'divTableCell'})

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

if dip != 'Y':
    from sgparank import gpa2

    gpa2(year, branch, low, high, semc, cycle)
