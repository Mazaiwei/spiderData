import time
import xlwt
import requests
from urllib3.exceptions import InsecureRequestWarning

requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

url = 'https://ggfw.hrss.gd.gov.cn/sydwbk/exam/details/spQuery.do'

allDatas = []
print('程序开始执行！')
for index in range(21):
    if index+1>9:
        num = f'{index+1}'
    else: num = f'0{index+1}'
    for page in range (25):
        data = {
            "bfa001": "2412121",
            "bab301": f"{num}",
            "page": f"{page+1}",
            "rows": "50"
        }

        headers = {
            '你的header'
        }
        resp = requests.post(url=url,headers=headers,data=data,verify=False)
        datas = resp.json()['rows']
        for index in datas:
            allDatas.append(index)
        print('添加一页数据进入列表，请等待数据录入Excel。程序运行结束显示才算结束，不然可能数据为空白')
        time.sleep(2)

wb = xlwt.Workbook()
sheet = wb.add_sheet('广东省事业编报名人数统计')
topList = ['招聘单位', '招聘岗位', '岗位代码', '聘用人数', '报名人数']
for index,list in enumerate(topList):
    sheet.write(0, index, list)
for index,data in enumerate(allDatas):
    sheet.write(index + 1, 0, data['aab004'])
    sheet.write(index + 1, 1, data['bfe3a4'])
    sheet.write(index + 1, 2, data['bfe301'])
    sheet.write(index + 1, 3, data['aab019'])
    sheet.write(index + 1, 4, data['aab119'])
    if index+1 % 50 == 0:
        print('准备录入新的一页数据')

wb.save('广东省事业编报名人数统计.xls')
print('程序运行结束！')


