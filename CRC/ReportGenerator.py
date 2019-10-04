# pack to exe: powershell run: auto-py-to-exe
# one file is not recommended, use one directory instead.
# include the additional files on the run (icon.png)
from appJar import gui
import fetchGoogleSheet as fgs
import pdfkit
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import threading
from bs4 import BeautifulSoup
import datetime
from re import match
from os import getcwd



def getLineChart(template, x, y):
	if(len(x) != len(y)):
		return
	base = 4
	while len(x) > 4:
		x = x[1:]
	while len(y) > 4:
		y = y[1:]

	pseudoX = x.copy()
	pseudoY = y.copy()

	if len(x) < 4:
		pseudoX.append(' ')
		pseudoY.append(y[len(y)-1])

	# print(x)
	# print(pseudoX)
	# print(y)
	# print(pseudoY)

	plt.style.use('fast')
	fig = plt.figure(figsize = [5,5])
	matplotlib.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'sans-serif']
	ax = plt.axes()
	minimum = min(y)
	minimum = min(minimum, base)
	maximum = max(y)
	maximum = max(maximum, base)
	range_  = maximum - minimum
	ticks   = int(maximum / 8)+1 # a $ticks between every ticks

	ytick   = [4]
	for i in range(8):
		ytick.append(i*ticks)
	ytick.sort()

	plt.xticks(range(len(pseudoX)), pseudoX)
	plt.yticks(ytick, ytick)
	if template == "CRC_EN_Template.html":
		plt.xlabel("Testing Date")
		plt.ylabel("CTC Count")
	else:
		plt.xlabel("檢測日期")
		plt.ylabel("CTC 顆數")

	y0= [ytick[0] for i in range(len(pseudoY))]
	y1= [ytick[1] for i in range(len(pseudoY))]
	y2= [ytick[2] for i in range(len(pseudoY))]
	y3= [ytick[3] for i in range(len(pseudoY))]
	y4= [ytick[4] for i in range(len(pseudoY))]
	y5= [ytick[5] for i in range(len(pseudoY))]
	y6= [ytick[6] for i in range(len(pseudoY))]
	y7= [ytick[7] for i in range(len(pseudoY))]
	plt.plot(pseudoX,y0, '-', color='black', linewidth=1)
	plt.plot(pseudoX,y1, '-', color='black', linewidth=1)
	plt.plot(pseudoX,y2, '-', color='black', linewidth=1)
	plt.plot(pseudoX,y3, '-', color='black', linewidth=1)
	plt.plot(pseudoX,y4, '-', color='black', linewidth=1)
	plt.plot(pseudoX,y5, '-', color='black', linewidth=1)
	plt.plot(pseudoX,y6, '-', color='black', linewidth=1)
	plt.plot(pseudoX,y7, '-', color='black', linewidth=1)

	base = [base for i in range(len(pseudoY))]
	plt.plot(pseudoX, base, '-', color='red', linewidth=2)

	plt.plot(x, y, '-', color='black', linewidth=2, markersize=1)
	for i in range(len(x)):
		plt.plot(x[i], y[i], 'o', color='black')
	# print("Saving picture: " + str(id))
	fig.savefig("tmp.jpg",figsize=(7, 7), dpi=1200)

def printPdfFromHtml(clientInfo, count, date, filename, template, directory):
	path_wkthmltopdf = r'wkhtmltopdf\\bin\\wkhtmltopdf.exe'
	config = pdfkit.configuration(wkhtmltopdf=path_wkthmltopdf)
	option = {
		'margin-top': '0mm',
		'margin-bottom':'0mm',
		'margin-left': '0mm',
		'margin-right': '0mm',
		'page-size': 'A4'
		}
	soup = BeautifulSoup(open('CRC Monitor HTML\\' + template, "r", encoding = "utf-8").read(), "html5lib")
	soup = cannedResponse(soup, template, count, clientInfo)
	soup = fillInById(soup, clientInfo, count, date)
	
	with open("CRC Monitor HTML\\htmlWithInfo.html", "w", encoding='utf-8') as file:
		file.write(str(soup))
	# print(soup.prettify().encode('utf-8'))
	pdfkit.from_file('CRC Monitor HTML\\htmlWithInfo.html', directory + "/" + filename, configuration=config, options = option)

def cannedResponse(soup, template, count, clientInfo):
	if template == "CRC_CH_Template.html":
		if clientInfo[3] == 'M':
			soup.find(id='gender').string               = '男'
		if clientInfo[3] == 'F':
			soup.find(id='gender').string               = '女'
		if len(count) >=2:
			if count[-1] >= 4 and count[-2] >= 4:
				soup.find(id='currentCtcTrend').string  = '高'
				soup.find(id='recommendation').string   = '建議更深入地以醫學影像或其他檢測檢查'
				soup.find(id='comments').string         = '您目前的循環腫瘤細胞趨勢為高。建議與其他影像及實驗室檢測結果綜合判讀。'
			elif count[-1] < 4 and count[-2] < 4:
				soup.find(id='currentCtcTrend').string  = '低'
				soup.find(id='recommendation').string   = '每三個月持續追蹤'
				soup.find(id='comments').string         = '您目前的循環細胞腫瘤趨勢為低。建議您仍持續每三個月接受腸追蹤檢測。倘若期間您的健康出現任何變化，請立即尋求專業醫師的協助。'
			else:
				soup.find(id='currentCtcTrend').string  = '目前無法判定'
				soup.find(id='recommendation').string   = '每三個月持續追蹤'
				soup.find(id='comments').string         = '您目前的循環腫瘤細胞數目為(' + str(count[-1]) + ')，然而，您上一次的檢測結果為(' + str(count[-2]) + ')，後續變化仍具不確定性。建議可與其他影像或實驗室檢測結果綜合判讀並持續每三個月接受腸追蹤檢測。倘若期間您的健康出現任何變化，請立即尋求專業醫師的協助。'
		else:
			if count[-1] >= 4:
				soup.find(id='currentCtcTrend').string  = 'NA'
				soup.find(id='recommendation').string   = '每三個月持續追蹤'
				soup.find(id='comments').string         = '您目前的循環腫瘤細胞數目高於標準值。然而腸追蹤檢測至少需連續二次檢測的數據來觀察循環腫瘤細胞變化趨勢，以監測疾病的發展。建議您仍持續每三個月接受腸追蹤檢測。期間倘若您的健康發生變化，請立即尋求專業醫師的協助。'
			else:
				soup.find(id='currentCtcTrend').string  = 'NA'
				soup.find(id='recommendation').string   = '每三個月持續追蹤'
				soup.find(id='comments').string         = '您目前的循環腫瘤細胞數目低於標準值。然而腸追蹤檢測至少需連續二次檢測的數據來觀察循環腫瘤細胞變化，以監測疾病的發展。建議您仍持續每三個月接受腸追蹤檢測。期間倘若您的健康發生變化，請立即尋求專業醫師的協助。'
			
	elif template == "CRC_EN_Template.html":
		soup.find(id='gender').string                   = clientInfo[3]
		if len(count) >=2:
			if count[-1] >= 4 and count[-2] >= 4:
				soup.find(id='currentCtcTrend').string  = 'High'
				soup.find(id='recommendation').string   = 'Confirmation diagnosis with imaging and/or other testing'
				soup.find(id='comments').string         = 'Your current CTC trend is high. Co​relation with imaging and other laboratory testing is ​​​​recommended. '
			elif count[-1] < 4 and count[-2] < 4:
				soup.find(id='currentCtcTrend').string  = 'Low'
				soup.find(id='recommendation').string   = 'Test every 3 months'
				soup.find(id='comments').string         = 'Your current CTC trend is low. We recommend to continue assessing CTC counts every 3 months. Meanwhile, if your health condition deteriorates, it is advised you seek for  a professional medical care at your local hospital or clinic.'
			else:
				soup.find(id='currentCtcTrend').string  = 'Indeterminate'
				soup.find(id='recommendation').string   = 'Test every 3 months'
				soup.find(id='comments').string         = 'Your current CTC count is ' + str(count[-1]) + ' with an indeterminate trend given the prior reading ' + str(count[-2]) + '. Further correlation with imaging and laboratory testing is recommended. Testing with three month interval is recommended for recurrence monitoring. Meanwhile, if your health condition deteriorates, it is advised you seek for a professional medical care at your local hospital or clinic.'
		else:
			if count[-1] >= 4:
				soup.find(id='currentCtcTrend').string  = 'NA'
				soup.find(id='recommendation').string   = 'Test every 3 months'
				soup.find(id='comments').string         = 'Your current CTC count is above the cutoff (4). Two consecutive readings with 3 months interval are required for the CTC trend analysis.  We recommend to continue testing every 3 months to obtain your CTC trend before acting upon. Meanwhile, if your health condition deteriorates, it is advised you seek for  a professional medical care at your local hospital or clinic.'
			else:
				soup.find(id='currentCtcTrend').string  = 'NA'
				soup.find(id='recommendation').string   = '每三個月持續追蹤'
				soup.find(id='comments').string         = '您目前的循環腫瘤細胞數目低於標準值。然而腸追蹤檢測至少需連續二次檢測的數據來觀察循環腫瘤細胞變化，以監測疾病的發展。建議您仍持續每三個月接受腸追蹤檢測。期間倘若您的健康發生變化，請立即尋求專業醫師的協助。'
	elif template == "CRC_simplified_CH_Template.html":
		if clientInfo[3] == 'M':
			soup.find(id='gender').string               = '男'
		if clientInfo[3] == 'F':
			soup.find(id='gender').string               = '女'
		if len(count) >=2:
			if count[-1] >= 4 and count[-2] >= 4:
				soup.find(id='currentCtcTrend').string  = '高'
				soup.find(id='recommendation').string   = '建议更深入地以医学影像或其他检测检查'
				soup.find(id='comments').string         = '您目前的循环肿瘤细胞趋势为高。建议与其他影像及实验室检测结果综合判读。'
			elif count[-1] < 4 and count[-2] < 4:
				soup.find(id='currentCtcTrend').string  = '低'
				soup.find(id='recommendation').string   = '每三个月持续追踪'
				soup.find(id='comments').string         = '您目前的循环细胞肿瘤趋势为低。建议您仍持续每三个月接受肠追踪检测。倘若期间您的健康出现任何变化，请立即寻求专业医师的协助。'
			else:
				soup.find(id='currentCtcTrend').string  = '目前无法判定'
				soup.find(id='recommendation').string   = '每三个月持续追踪'
				soup.find(id='comments').string         = '您目前的循环肿瘤细胞数目为(' + str(count[-1]) + ')，然而，您上一次的检测结果为(' + str(count[-2]) + ')，后续变化仍具不确定性。建议可与其他影像或实验室检测结果综合判读并持续每三个月接受肠追踪检测。倘若期间您的健康出现任何变化，请立即寻求专业医师的协助。'
		else:
			if count[-1] >= 4:
				soup.find(id='currentCtcTrend').string  = 'NA'
				soup.find(id='recommendation').string   = '每三个月持续追踪'
				soup.find(id='comments').string         = '您目前的循环肿瘤细胞数目高于标准值。然而肠追踪检测至少需连续二次检测的数据来观察循环肿瘤细胞变化趋势，以监测疾病的发展。建议您仍持续每三个月接受肠追踪检测。期间倘若您的健康发生变化，请立即寻求专业医师的协助。'
			else:
				soup.find(id='currentCtcTrend').string  = 'NA'
				soup.find(id='recommendation').string   = '每三个月持续追踪'
				soup.find(id='comments').string         = '您目前的循环肿瘤细胞数目低于标准值。然而肠追踪检测至少需连续二次检测的数据来观察循环肿瘤细胞变化，以监测疾病的发展。建议您仍持续每三个月接受肠追踪检测。期间倘若您的健康发生变化，请立即寻求专业医师的协助。'
	return soup		

def fillInById(soup, clientInfo, count, date):
	#clientInfo = [name, bd, clientId, gender, inds[len(inds)-1], twId, sampleCollectingSite, vs, telephone, email]
	today = datetime.date.today().strftime("%Y/%m/%d")
	bdMonth, bdDay, bdYear = clientInfo[1].split("/")
	#First Page
	soup.find(id='requisitionNumber').string        = clientInfo[4]
	soup.find(id='patientName').string              = clientInfo[0]
	soup.find(id='idNumber').string                 = clientInfo[5]
	soup.find(id='dateOfBirth').string              = bdYear + "/" + bdMonth + "/" + bdDay
	soup.find(id='patientPhoneNumber').string       = clientInfo[8]
	soup.find(id='patientEmail').string             = clientInfo[9]
	soup.find(id='nameOfLab').string                = clientInfo[6]
	soup.find(id='labPhoneNumber').string           = ''
	soup.find(id='nameOfPhysician').string          = clientInfo[7]
	soup.find(id='dateOfCollection').string         = date[len(date)-1]
	soup.find(id='dateOfReport').string             = today
	#Second Page
	soup.find(id='currentCtcCount').string          = str(count[len(count)-1])
	


	soup.find(id='eSignDate1').string               = today
	soup.find(id='eSignDate2').string               = today
	# soup.find(id='tumorType').string                = ''
	# soup.find(id='dateOfDiagnosis').string          = ''
	# soup.find(id='pathologicalDiagnosis').string    = ''
	# soup.find(id='tumorStage').string               = clientInfo[9]
	# soup.find(id='metastasisSite').string           = ''
	# soup.find(id='dateOfMostRecentTreatment').string= ''
	# soup.find(id='table004-B1').string              = str(date[0])
	# soup.find(id='table004-B2').string              = str(count[0])
	while len(date) > 4:
		date = date[1:]
	while len(count) > 4:
		count = count[1:]
	for i in range(len(count)):
		columnA = "table005-A" + str(2+i)
		columnB = "table005-B" + str(2+i)
		soup.find(id = columnA).string              = str(date[i])
		soup.find(id = columnB).string              = str(count[i])

	return soup

def press(button):
	if button == "Cancel":
		app.stop()
	else:
		app.hide()
		name = str(app.entry("Name"))
		bd = str(app.entry("Birthday"))
		bd = bd.split("/")[1] + "/" + bd.split("/")[2] + "/" + bd.split("/")[0]
		print("Name:", name, "Birthday:", bd)

		filename = str(app.entry("Filename"))
		if filename == '':
			filename = 'report.pdf'
		if not match('^.*\.pdf$', filename):
			filename = filename + ".pdf"
		print("Filename: " + filename)

		template = ''
		if app.getOptionBox("Templates") == "Traditional Chinese":
			template = "CRC_CH_Template.html"

		elif app.getOptionBox("Templates") == "Simplified Chinese":
			template = "CRC_simplified_CH_Template.html"

		elif app.getOptionBox("Templates") == "English":
			template = "CRC_EN_Template.html"
			
		else:
			app.errorBox("Not found", "Error happened in initializing templates.")	
			return


		directory = app.entry("Save at")
		if directory == '':
			directory = getcwd()
		print("Output dir: " + directory)
		print("Press Ctrl+C (twice or more)to quit")

		credential = fgs.init()
		clientId, gender, twId = fgs.getId(credential, [name], [bd])
		if clientId != "Not found":
			print('client id: ' + clientId)
			inds, date= fgs.getRecord(credential, clientId)
			print("inds: " + str(inds))
			print("date: " + str(date))
			count, sampleCollectingSite, telephone, email, vs = fgs.fetchReportCount(credential, inds)
			print("count:" + str(count))
			print("sampleCollectingSite: " + str(sampleCollectingSite))
			print("vs: " + str(vs))
			clientInfo = [name, bd, clientId, gender, inds[len(inds)-1], twId, sampleCollectingSite, vs, telephone, email]
			
			# collection date 
			# x,y = fgs.getRecord(credential, clientId)
			# Create new threads
			thread1 = myThread(1, "Thread_getLineChart", [getLineChart, template, date.copy(), count.copy()])
			thread2 = myThread(2, "Thread_printPdfFromHtml", [printPdfFromHtml, clientInfo, count.copy(), date.copy(), filename, template, directory])


			# Start new Threads
			thread1.start()
			thread2.start()

			# Add threads to thread list
			threads.append(thread1)
			threads.append(thread2)

			# Wait for all threads to complete
			for t in threads:
				t.join()
			print("All thread joined")
			app.infoBox("Success", "The report is generated under the current folder with name " + filename + "\n" +
				"Rename it before generating a new one to avoid overwriting.\n" + 
				"\tRecords found: " + str(len(inds)) +
				"\n\tID Number: " + twId +
				"\n\tName of Lab: " + sampleCollectingSite + 
				"\n\tName of Physician: " + vs)
			app.clearAllEntries(callFunction=False)
		else:
			app.errorBox("Not found", "The information you entered is not found.")
			app.clearAllEntries(callFunction=False)
		app.show()

def executeThreadFunc(args):
	if args[0] == getLineChart:
		getLineChart(args[1], args[2], args[3])
	elif args[0] == printPdfFromHtml:
		printPdfFromHtml(args[1], args[2], args[3], args[4], args[5], args[6])
	elif args[0] == fgs.fetchReportCount:
		fgs.fetchReportCount(args[1], args[2])

class myThread (threading.Thread):
	def __init__(self, threadID, name, args):
		threading.Thread.__init__(self)
		self.threadID = threadID
		self.name = name
		self.args = args
	def run(self):
		print("Starting " + self.name)
		# Get lock to synchronize threads
		threadLock.acquire()
		print(self.name + " acquired lock.")
		try:
			executeThreadFunc(self.args)
		except Exception as e:
			print("Error raised in " + self.name + " by " + str(self.args[0]) + " with error message: " + str(e))
		# Free lock to release next thread
		threadLock.release()
		print(self.name + " released lock.")

###############
#MAIN FUNCTION#
###############
threadLock = threading.Lock()
threads = []
app = gui("CellmaxLife Report Generator", "450x200")
app.setBg("white")
app.setIcon("icon.gif")
app.setFont(12)
app.addLabelEntry("Name")
app.addLabelEntry("Birthday")
app.addLabelEntry("Filename")
app.setEntryDefault("Name", "Enter Name")
app.setEntryDefault("Birthday", "Birthday in YYYY/MM/DD")
app.setEntryDefault("Filename", "The filename you want to save as. (Default: report.pdf)")
app.addLabelOptionBox("Templates", ["Traditional Chinese", "Simplified Chinese", "English"])
app.addDirectoryEntry("Save at")
app.setEntryDefault("Save at", "(Default: current directory)")
app.addButtons(["Submit", "Cancel"], press)
app.enableEnter(press)
app.setFocus("Name")
app.go()
