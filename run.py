# я старался очень-очень
# (ʘ‿ʘ✿)
# в процессе написания
# знаю, что лагает - исправлю
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

import requests
from bs4 import BeautifulSoup as soup

############### xlsx.opener ###############


def open_xls():
    wb = load_workbook(filename='ADD_py.xlsx')
    ws = wb.get_sheet_by_name('A1')

    my_list = []

    for row in ws.iter_rows(min_row=2, max_col=4, max_row=2):
        for cell in row:
            my_list.append(cell.value)
    # drug_name = my_list[1]
    # drug_dosage = my_list[2]
    # print(my_list)

    return my_list
###################### decoder ######################


def decoder(cod):
    return(cod.decode('cp1251')).encode('utf8')

########### selenium ####################
# вот здесь она лагает - не всегда догружается элемент(кнопка), в итоге кликать нечего и всё летит к чертям:)
def get_page(drug_name, drug_dosage):
    browser = webdriver.Chrome(executable_path='C:\Program Files (x86)\Google\Chrome\Application\chromedriver')
    browser.get("https://tabletki.ua/list/")
    wait = WebDriverWait(browser, 30)
    search_bar = browser.find_element_by_id("ctl00_ctl00_HeaderPanel_QueryPanel_QueryCtrl")
    search_bar.send_keys(drug_name)
    search_button = browser.find_element_by_id('ctl00_ctl00_HeaderPanel_QueryPanel_SubmitLink')
    search_button.click()
    inst_button = browser.find_element_by_xpath('//*[@id="smart-menu"]/ul/li[1]/a')
    inst_button.click()
    dosage_button = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_HeaderPanel_QueryPanel_ControlsPanel"]/div[1]/div[1]')))
    dosage_button.click()
    dropdown_list = wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="ctl00_HeaderPanel_QueryPanel_ControlsPanel"]/div[1]/div[1]/ul')))
    option = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[contains(text(), "%s")]' % drug_dosage))).click()
    my_page = browser.current_url
    # print(my_page)
    return my_page

######################## Page parsing ######################


def get_page_content(page):
    # page = browser.current_url
    # URL = 'https://tabletki.ua/%D0%A6%D0%B8%D0%B1%D0%BE%D1%80/25903/'
    r = requests.get('%s' % page)
    page = r.content
    page_soup = soup(page, "html.parser")
    # print(page_soup.prettify())
    drug_data = page_soup.find_all('div', {"id": "ctl00_MAIN_ContentPlaceHolder_NeedTranslatePanel"})
    # return drug_data
    # print(drug_data)

    for item in drug_data:
        # print(item.contents[1].find_all(['h2']))
        for header in item.contents[1].find_all(['h2']):
            print(header.get_text())
            for elem in header.next_elements:
                if elem.name and elem.name.startswith('h'):
                    break
                if elem.name == 'p':
                    print(elem.get_text())
    return


####################### docx ######################

#in progress....

# from docx import Document
# doc = Document()
# my_text = doc.add_paragraph(instruction_text())
#
# doc.save('Цибор.docx')

#########
def main():
    open_xls()
    my_list = open_xls()
    drug_name = my_list[1]
    drug_dosage = my_list[2]
    my_page = get_page(drug_name, drug_dosage)
    get_page_content(page=my_page)

    # print(drug_name)


if __name__ == "__main__":
     main()
