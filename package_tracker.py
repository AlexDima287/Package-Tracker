from urllib.request import urlopen
from bs4 import BeautifulSoup
import mechanicalsoup
import re
import pandas as pd
import xlrd

#inputs 
spreadsheet = 'spreadsheet.xlsx' #name of spreadhseet
column = 2 #column input, index starts at 0
# spd2 = r'C:\Users\alexd\OneDrive\Documents\MATLAB\spreadsheet.xlsx'


#setup
browser = mechanicalsoup.Browser()
url = "https://bnitracking.com/"
login_page = browser.get(url)
login_html = login_page.soup

#parse html for current status on delivery
def current_stat(array, i):
    form = login_html.select("form")[0]
    form.select("input")[0]["value"] = str(array[i,column])
    profiles_page = browser.submit(form, login_page.url)

    new_page = profiles_page.soup.get_text()

    res = re.search("Current Status", new_page)
    start_index = res.end() +3;

    for i in range (start_index,start_index +30):
        if new_page[i] == '\n':
            end_index = i
            break

    current_status = str(new_page[start_index:end_index])

    return current_status

#loop through spreadsheet
def excel_loop(spreadsheet, login_html):
    loc = (spreadsheet)
    sheet = pd.read_excel(loc)
    array = sheet.to_numpy()
    current_status = []

    for i in range(0,int(sheet.shape[0])):
        current_status.append(current_stat(array, i))
        print("In progress ...\r\n")

    new_col = pd.DataFrame(current_status)
    print(new_col)

    
    with pd.ExcelWriter(spreadsheet,mode='a') as writer:  
        new_col.to_excel(writer, sheet_name='Current Status')
    
        
excel_loop(spreadsheet,login_html)
    
 












