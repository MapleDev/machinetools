from bs4 import BeautifulSoup
from urllib.request import urlopen
import urllib
import re
import requests
import webbrowser
import xlwt

baseUrl = "https://www.machinetools.com"

# 列表页 page [1...50]
# listUrl = "https://www.machinetools.com/zh-CN/distributors/machines?_page_size=200&page=%s"
listUrl = baseUrl + "/zh-CN/distributors/machines?_page_size=200&page=%s"
# for test
# listUrl = baseUrl + "/zh-CN/distributors/machines?_page_size=1&page=%s"
# listUrl = baseUrl + "/zh-CN/distributors/machines?_page_size=10000&page=%s"

# 详细页
detailUrlTempl = baseUrl + "/zh-CN/companies/%s"
# 查看电话号码页
phoneRequestUrlTempl = baseUrl + "/zh-CN/companies/%s/phone_request"
# 查看网址页
websiteRequestUrlTempl = baseUrl + "/zh-CN/companies/%s/website_request"

txt = open(r'out.txt', 'w+', encoding='utf-8')

# 创建 xls 文件对象
wb = xlwt.Workbook()
# 新增一个表单
sh = wb.add_sheet('1st Sheet')
sh.write(0, 0, '名称')
sh.write(0, 1, '地址')
sh.write(0, 2, '电话')
sh.write(0, 3, '免费电话')
sh.write(0, 4, '传真')
sh.write(0, 5, '网址')
sh.write(0, 6, '简介')
sh.write(0, 7, '类型')
sh.write(0, 8, '代理品牌')
sh.write(0, 9, '详情页')

col = 1

# 修改这里
startPage = 12
for page in range(startPage, 51):
# for page in range(1, 2):
	print('------- start page: ' + str(page) + '-------')

	listUrlCompete = listUrl % (page)
	print(str(page) + ": 列表页 listUrlCompete = " + listUrlCompete)
	html = urlopen(listUrlCompete).read().decode('utf-8')

	# param = {"_page_size": '2'}
	# r = requests.get(listUrl, params = param)
	# listUrlCompete = r.url
	# html = r.text

	soup = BeautifulSoup(html, features='lxml')

	# all_td = soup.find_all('td')
	# month = soup.find_all('li', {"class": "month"})
	all_div = soup.find_all('div', {"class": "field-value-display trimmed show-for-small-only"})
	# print('all_div = ', all_div)
	for div in all_div:
		print('------- start col: ' + str(col) + '-------')

		# div = div.get_text()

		a = div.find('div').find('div').find('a')
		# print('a = ', a)
		
		distributors_href = a['href']
		# print('distributors_href = ', distributors_href)
		companyName = distributors_href.replace("/zh-CN/distributors/", "")
		# print("	公司名称companyName = " + companyName)
		detailUrl = detailUrlTempl % companyName
		print("	详情页 detailUrl = " + detailUrl)

		# fun1. with Cookie
		# header = { 'Cookie' : '__stripe_mid=d25220e0-3c00-46f4-8242-f01eece66b1b; _ga=GA1.2.1698389194.1553131771; _gid=GA1.2.567369068.1553131771; _mt_suid=eyJfcmFpbHMiOnsibWVzc2FnZSI6IklqUTNPVFU0TXlJPSIsImV4cCI6IjIwMTktMDQtMjBUMDE6MzA6MjYuNTkwWiIsInB1ciI6bnVsbH19--24bc79a77fde6a5d02d589977d07ee807f1fa522; _mt_sutoken=eyJfcmFpbHMiOnsibWVzc2FnZSI6IklqVTJPV1UzTVdFd1pHUTFOMlZqTkdVMFpEaG1PRFl5WVdRM1pERmxNeUk9IiwiZXhwIjoiMjAxOS0wNC0yMFQwMTozMDoyNi41OTBaIiwicHVyIjpudWxsfX0%3D--e27af86764c87295d898331a5e8b3c9a420371b4; _mt_uat-v2=eyJfcmFpbHMiOnsibWVzc2FnZSI6IklqWTFaVFk0T1dGaU0yVmtZVGM1T1dZNE9UWTNNMlZrTXpGak1HSTJNamcxWmpFek1XUm1aVFEwTURRME5EWmpObUprWTJRd016TmlaVGd5T0RBM1pUVWkiLCJleHAiOiIyMDE5LTA2LTIxVDAxOjMwOjI2LjU5MFoiLCJwdXIiOm51bGx9fQ%3D%3D--c805dd144dfda6bb66c7e8337b1742ef2f581cd2; __stripe_sid=8e2647f0-f3a3-427e-a1e1-56bd3a4137f0; _gat_UA-1392039-4=1; breakpoint=medium; _mt_session=19%2FY7EOhOHYd4bixnIT24KTIc%2FjV%2FlVO71LyMTbzwjcx7NDxZ%2F%2FZiKIvZ8TP9a0ig7ZoKVOsomq9VKWT4cq5jG2UiAiyEK%2BYvrfumBV9fl361pf0f9lQW0zf08DDhYk1iO9bjbw7gcYAIuOEzjc%3D--kn7tdCfnwx3R0FZE--K97CL2qU6e%2BA0cNh8Kujdw%3D%3D' }  
		# request = urllib.request.Request(detailUrl, headers = header)
		# response = urllib.request.urlopen(request)
		# detailUrlHtml = str(response.read().decode('utf-8'))

		# fun2. without Cookie
		detailUrlHtml = urlopen(detailUrl).read().decode('utf-8')

		detailUrlSoup = BeautifulSoup(detailUrlHtml, features='lxml')
		# detailPage_all_div = detailUrlSoup.find_all('div', {"class": "medium-8 columns"})
		# print('		detailPage_all_div = ', detailPage_all_div)
		detailPageAllDiv = detailUrlSoup.find('div', {"class": "medium-8 columns"})
		# =================================== 解析详情页		

		# for detailPageAllDiv in detailPage_all_div:
		# 公司 地址 电话
		fieldValueDisplayDiv = detailPageAllDiv.find('div').find('div').find('div')
		# print('		fieldValueDisplayDiv = ', fieldValueDisplayDiv)
		# allField = BeautifulSoup(fieldValueDisplayDiv, features='lxml').find_all('div', {"class": "field"})
		allField = BeautifulSoup(str(fieldValueDisplayDiv), features='lxml').find_all('div', {"class": "label"})
		# print("allField = ", allField)
		for field in allField:
			value = field.find_next_sibling()
			fieldText = field.get_text()
			valueText = value.get_text()

			if fieldText == '公司':
				sh.write(col, 0, valueText)
			if fieldText == '地址':
				valueText = valueText.replace("查看当地时间", "")
				sh.write(col, 1, valueText)
			elif fieldText == '电话':
				if valueText == '查看电话号码':
					phoneRequestUrl = phoneRequestUrlTempl % companyName
					# print('	需再次查询获取电话 phoneRequestUrl = ' + phoneRequestUrl)
					# request = urllib.request.Request(detailUrl, headers = header)
					# response = urllib.request.urlopen(request)
					# phoneRequestUrlHtml = str(response.read().decode('utf-8'))
					# print("	phoneRequestUrlHtml = " + phoneRequestUrlHtml)
					# txt.write(phoneRequestUrlHtml + "\n")
				else:
					sh.write(col, 2, valueText)
			elif fieldText == '免费电话':
				sh.write(col, 3, valueText)
			elif fieldText == '传真':
				sh.write(col, 4, valueText)
			elif fieldText == '网址':
				if valueText == '查看网址':
					websiteRequestUrl = websiteRequestUrlTempl % companyName
					# print('	需再次查询获取网址 websiteRequestUrl = ' + websiteRequestUrl)
				else:
					sh.write(col, 5, valueText)

			print("fieldText = " + fieldText + "; valueText = " + valueText)

		# 简介
		fieldValueDisplayDiv = detailUrlSoup.find('div', {"class": "text-block"})
		# print('		fieldValueDisplayDiv = ', fieldValueDisplayDiv)
		if not fieldValueDisplayDiv is None:
			valueText = fieldValueDisplayDiv.get_text()
			print("fieldText = " + '简介' + "; value = " + valueText)
			sh.write(col, 6, valueText)
		# 代理的品牌
		fieldValueDisplayDiv = detailUrlSoup.find_all('li', {"class": "accordion-item"})
		# print('		fieldValueDisplayDiv2 = ', fieldValueDisplayDiv)
		allField = BeautifulSoup(str(fieldValueDisplayDiv), features='lxml').find_all('a', {"class": "structured-link accordion-title"})
		for field in allField:
			value = field.find_next_sibling()
			fieldText = field.get_text()
			valueText = value.get_text()
			print("fieldText = " + fieldText + "; value = " + valueText)
			if fieldText == '公司类型':
				sh.write(col, 7, valueText)
			elif fieldText == '我们代理的品牌':
				sh.write(col, 8, valueText)

		sh.write(col, 9, detailUrl)

		# wb.save('machinetools.xls')
		print('------- end col: ' + str(col) + '-------')
		col = col + 1
		# txt.write(str(fieldValueDisplayDiv) + "\n")
			
		# ===================================


	print('------- end page: ' + str(page) + '-------')
	wb.save('machinetools_' + str(page) + '.xls')


txt.close()
# wb.save('machinetools.xls')

print('------- end -------')

