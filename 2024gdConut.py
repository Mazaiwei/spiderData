import time
import xlwt
import requests
from urllib3.exceptions import InsecureRequestWarning

requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

url = 'https://ggfw.hrss.gd.gov.cn/sydwbk/exam/details/spQuery.do'

allDatas = []
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
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "Accept-Encoding": "gzip, deflate, br, zstd",
            "Accept-Language": "zh-CN,zh;q=0.9",
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Cookie": "JSESSIONID=DBKZc-IITBrZ-KgrAdWpYCLIbDrRveZML6M3s7jT6PR_tXFd9yVW!-988052771; _gscu_260182935=05460498k0095j21; _gscbrs_260182935=1; arialoadData=true; ../issoYH_MASS_TOKEN=7e83beb5-5ddd-4057-961f-7063891e9ec9; YH_MASS_TOKEN=aaadfe06-339f-4f62-9e87-ab7e453ad3b0",
            "Host": "ggfw.hrss.gd.gov.cn",
            "Origin": "https://ggfw.hrss.gd.gov.cn",
            "Pragma": "no-cache",
            "Referer": "https://ggfw.hrss.gd.gov.cn/sydwbk/center.do?nvt=1711973392060",
            "Sec-Ch-Ua": "\"Not A(Brand\";v=\"99\", \"Google Chrome\";v=\"121\", \"Chromium\";v=\"121\"",
            "Sec-Ch-Ua-Mobile": "?0",
            "Sec-Ch-Ua-Platform": "\"Windows\"",
            "Sec-Fetch-Dest": "empty",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Site": "same-origin",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
            "X-Requested-With": "XMLHttpRequest"
        }
        resp = requests.post(url=url,headers=headers,data=data,verify=False)
        datas = resp.json()['rows']
        for index in datas:
            allDatas.append(index)
        print('添加一页数据进入数组，请等待数据录入Excel。程序运行结束显示才算结束，不然可能数据为空白')
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
    if index+1 / 50 == 0:
        print('准备录入新的一页数据')

wb.save('广东省事业编报名人数统计.xls')
print('程序运行结束！')


