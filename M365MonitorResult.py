#coding=utf-8
from selenium import webdriver
import os

if os.path.exists('1.txt'):
    os.remove('1.txt')
file = open('1.txt', 'a')

def WriteDataToTXT():
    getSingularData = driver.find_elements_by_xpath(
        "//tr[@style='color:Black;background-color:MistyRose;']/td[@class='TAC']")
    getSingularName = driver.find_elements_by_xpath(
        "//tr[@style='color:Black;background-color:MistyRose;']/td[@class='TAL']")

    getpluralData = driver.find_elements_by_xpath(
        "//tr[@style='color:Black;background-color:MistyRose;white-space:nowrap;']/td[@class='TAC']")
    getpluralName = driver.find_elements_by_xpath(
        "//tr[@style='color:Black;background-color:MistyRose;white-space:nowrap;']/td[@class='TAL']")

    listSingularData = []
    listSingularName = []
    listpluralData = []
    listpluralName = []

    for i in getSingularData:
        listSingularData.append(i.text)
    for i in getSingularName:
        listSingularName.append(i.text)
    for i in getpluralData:
        listpluralData.append(i.text)
    for i in getpluralName:
        listpluralName.append(i.text)

    newSingularName = []
    newSingularData = []
    newpluralName = []
    checkCurrentPageSingularCount = int(len(listSingularData)/11)
    checkCurrentPagePluralCount = int(len(listpluralData)/11)

    for i in range(0, len(listpluralName), 3):
        splipList = listSingularName[i:i + 3]
        newSingularName.append(splipList[1])
        newSingularName.append(splipList[2])

        splipList = listpluralName[i:i + 3]
        newpluralName.append(splipList[1])
        newpluralName.append(splipList[2])

    if checkCurrentPageSingularCount - checkCurrentPagePluralCount == 1:
        newSingularName.append(listSingularName[-2])
        newSingularName.append(listSingularName[-1])
        newSingularData.append(listSingularData[-11])
        newSingularData.append(listSingularData[-10])
        newSingularData.append(listSingularData[-9])

    flag = 0
    for i in range(0, len(listpluralData), 11):
        splipListSingular = listSingularData[i:i + 11]
        splipListplural = listpluralData[i:i + 11]

        if flag <= i:
            file.writelines(
                '{:},{:},{:},{:},{:}\n'.format(newSingularName[flag], newSingularName[flag + 1], splipListSingular[0],
                                               splipListSingular[1], splipListSingular[2]))
            file.writelines(
                '{:},{:},{:},{:},{:}\n'.format(newpluralName[flag], newpluralName[flag + 1], splipListplural[0],
                                               splipListplural[1], splipListplural[2]))
            flag = flag + 2

    if checkCurrentPageSingularCount - checkCurrentPagePluralCount == 1 and i == len(listpluralData) - 11:
        file.writelines(
            '{:},{:},{:},{:},{:}\n'.format(newSingularName[flag], newSingularName[flag + 1],
                                           newSingularData[0],
                                           newSingularData[1], newSingularData[2]))


base_url = "https://exfocusb.extest.microsoft.com/Exchange15/Report.aspx?FocusId=2308122&GroupId=3437722&TestResultType=5"

# options = webdriver.ChromeOptions()
# options.binary_location = r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
# driver = webdriver.Chrome(chrome_options=options)

driver = webdriver.Chrome()
driver.get(base_url)
driver.implicitly_wait(50)
driver.find_elements_by_id("FilterOptionsExpander_TableContent")
driver.maximize_window()

splitgetTotal = ""
getTotal = driver.find_elements_by_xpath("//td[@style='text-align:center;width:100%;']/span/nobr")
for i in getTotal:
    splitgetTotal = i.text

splitgetTotal = splitgetTotal.replace("-", " ")
splitgetTotal = splitgetTotal.split(" ")
getTotalData = int(splitgetTotal[-1])
getDisplayData = int(splitgetTotal[-3])
jumps = int(getTotalData/getDisplayData)
if getTotalData % 50 != 0 and getTotalData > 50:
    jumps = jumps + 1

flagPageNumber = 1
currentNumber = ""
for n in range(0, jumps, 1):
    pageNumber = driver.find_elements_by_xpath("//input[contains(@class,'PagingTextBox')]")
    for i in pageNumber:
        currentNumber = i.get_attribute('value')
    if flagPageNumber == int(currentNumber):
        WriteDataToTXT()
        print(n + 1)

    clickButton = driver.find_elements_by_xpath("//input[@src='Images/Navigation/Page_Next.ico']")
    for a in clickButton:
        currentNumber = a.click()

    flagPageNumber = flagPageNumber + 1

file.close()
