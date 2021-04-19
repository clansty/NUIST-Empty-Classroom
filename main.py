import time
import xlwt
import json
from selenium import webdriver
from selenium.webdriver.support.ui import Select

USERNAME = ''  # 这里填写用户名和密码
PASSWORD = ''
BUILDING = '滨江'  # 滨江 / 揽江楼 / 明德 / 文德....
COLLECTOR = 'Clansty'

TIMES = ['上午1～2', '上午3～4', '下午5～6', '下午7～8', '晚上9～10', '晚上11～12']
WEEKDAYS = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']

data = []
# 列：i，星期
# 行：j，上课时段
for i in range(7):
    toInsert = []
    for j in range(6):
        toInsert.append('')
    data.append(toInsert)
# 现在 data 是一个包含七个（包含六个字符串的 List）的 List

driver = webdriver.Chrome()
driver.implicitly_wait(30)
driver.get('https://authserver.nuist.edu.cn/authserver/login')
driver.find_element_by_id('username').send_keys(USERNAME)
driver.find_element_by_id('password').send_keys(PASSWORD)
driver.find_element_by_id('login_submit').click()

time.sleep(2)
driver.get('http://bkxk.nuist.edu.cn/(S(sctgsmvdnwzoxhmpy44qvjmy))/public/newslist.aspx')
time.sleep(2)
driver.get('http://bkxk.nuist.edu.cn/(S(sctgsmvdnwzoxhmpy44qvjmy))/public/jxcdkebiaoall.aspx')

classroomType = Select(driver.find_element_by_id('DropDownList4'))
classroomType.select_by_value('多媒体教室')
driver.find_element_by_id('Button1').click()

classroomName = Select(driver.find_element_by_id('DropDownList3'))
allClassrooms = []
for i in classroomName.options:
    if (i.text.startswith(BUILDING)):
        allClassrooms.append(i.text)

for i in allClassrooms:
    classroomName = Select(driver.find_element_by_id('DropDownList3'))
    classroomName.select_by_value(i)
    print(i)
    time.sleep(1)
    # 遍历表格
    tbody = driver.find_element_by_id('TABLE1').find_element_by_tag_name('tbody')
    # 第 2～6 个 <tr>
    rows = tbody.find_elements_by_tag_name('tr')[1:7]
    # rowId: 上课时间段
    for rowId in range(6):
        # 第 2～8 个 <td>
        row = rows[rowId]
        cells = row.find_elements_by_tag_name('td')[1:9]
        # cellId: 星期几
        for cellId in range(7):
            cell = cells[cellId]
            if cell.text == ' ':
                # 这里没课
                data[cellId][rowId] += i[len(BUILDING):] + '，'  # 去掉前面楼的名称

# 保存到 json
jsonFile = open(BUILDING + '.json', 'w')
jsonFile.write(json.dumps(data))
jsonFile.close()

# 写表格
wb = xlwt.Workbook()
table = wb.add_sheet(BUILDING + ' 空教室列表')
# 第一个参数是行，第二个参数是列
table.write(0, 0, BUILDING + ' 空教室列表')
for i in range(6):
    # 写行标题
    table.write(i + 2, 0, TIMES[i])

for i in range(7):
    # 写列标题
    table.write(1, i + 1, WEEKDAYS[i])

table.write(8, 0, 'Data collected by ' + COLLECTOR)

for i in range(7):  # 列，星期几
    for j in range(6):
        # 从第三行第二列开始写
        table.write(j + 2, i + 1, data[i][j])

wb.save(BUILDING + '.xls')
