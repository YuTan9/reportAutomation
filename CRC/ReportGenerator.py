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
import time
from bs4 import BeautifulSoup
import datetime
import re


def getLineChart(id, x, y):
	if(len(x) != len(y)):
		return
	base = y[0]
	while len(x) > 6:
		x = x[1:]
	while len(y) > 6:
		y = y[1:]

	pseudoX = x.copy()
	pseudoY = y.copy()

	if len(x) < 6:
		pseudoX.append(' ')
		pseudoY.append(y[len(y)-1])

	# print(x)
	# print(pseudoX)
	# print(y)
	# print(pseudoY)

	plt.style.use('fast')
	fig = plt.figure(figsize = [7,7])
	matplotlib.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'sans-serif']
	ax = plt.axes()
	minimum = min(y)
	minimum = min(minimum, base)
	maximum = max(y)
	maximum = max(maximum, base)
	range_  = maximum - minimum
	ticks   = int(maximum / 8)+1 # a $ticks between every ticks

	ytick   = []
	for i in range(8):
		ytick.append(i*ticks)

	plt.xticks(range(len(pseudoX)), pseudoX)
	plt.yticks(ytick, ytick)
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
	soup = BeautifulSoup(open('PanCA Monitor HTML\\' + template, "r", encoding = "utf-8").read(), "html5lib")
	soup = fillInById(soup, clientInfo, count, date)

	with open("PanCA Monitor HTML\\htmlWithInfo.html", "w", encoding='utf-8') as file:
		file.write(str(soup))
	# print(soup.prettify().encode('utf-8'))
	pdfkit.from_file('PanCA Monitor HTML\\htmlWithInfo.html', directory + "/" + filename, configuration=config, options = option)

def fillInById(soup, clientInfo, count, date):
	today = datetime.date.today().strftime("%Y/%m/%d")
	bdMonth, bdDay, bdYear = clientInfo[1].split("/")
	#First Page
	soup.find(id='requisitionNumber').string        = clientInfo[4]
	soup.find(id='patientName').string              = clientInfo[0]
	soup.find(id='idNumber').string                 = clientInfo[5]
	soup.find(id='dateOfBirth').string              = bdYear + "/" + bdMonth + "/" + bdDay
	if clientInfo[3] == 'M':
		soup.find(id='gender').string               = '男'
	if clientInfo[3] == 'F':
		soup.find(id='gender').string               = '女'	
	soup.find(id='patientPhoneNumber').string       = ''
	soup.find(id='patientEmail').string             = ''
	soup.find(id='nameOfLab').string                = clientInfo[6]
	soup.find(id='labPhoneNumber').string           = ''
	soup.find(id='nameOfPhysician').string          = clientInfo[7]
	soup.find(id='dateOfCollection').string         = date[len(date)-1]
	soup.find(id='dateOfReport').string             = today
	#Second Page
	soup.find(id='currentCtcCount').string          = str(count[len(count)-1])
	soup.find(id='baselineCount').string            = str(count[0])
	soup.find(id='currentCtcTrend').string          = ''

	soup.find(id='comments').string                 = ''
	soup.find(id='eSignDate1').string               = today
	soup.find(id='eSignDate2').string               = today
	soup.find(id='tumorType').string                = ''
	soup.find(id='dateOfDiagnosis').string          = ''
	soup.find(id='pathologicalDiagnosis').string    = ''
	soup.find(id='tumorStage').string               = clientInfo[9]
	soup.find(id='metastasisSite').string           = ''
	soup.find(id='dateOfMostRecentTreatment').string= ''
	soup.find(id='table004-B1').string              = str(date[0])
	soup.find(id='table004-B2').string              = str(count[0])
	while len(date) > 6:
		date = date[1:]
	while len(count) > 6:
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
		if not re.match('^.*\.pdf$', filename):
			filename = filename + ".pdf"
		print("Filename: " + filename)

		template = ''
		if app.getOptionBox("Templates") == "Traditional Chinese":
			template = "PanCA_CH_Template.html"

		elif app.getOptionBox("Templates") == "Simplified Chinese":
			template = "PanCA_simplified_CH_Template.html"

		elif app.getOptionBox("Templates") == "Captain":
			template = "PanCA_Captain_Template.html"

		elif app.getOptionBox("Templates") == "Elite":
			template = "PanCA_Elite_Template.html"

		elif app.getOptionBox("Templates") == "Shin Kong":
			template = "PanCA_SK_Template.html"

		elif app.getOptionBox("Templates") == "English":
			template = "PanCA_EN_Template.html"
			
		else:
			app.errorBox("Not found", "Error happened in initializing templates.")	
			return

		directory = app.entry("Save at")
		print("Press Ctrl+C (twice or more)to quit")

		credential = fgs.init()
		clientId, gender, twId = fgs.getId(credential, [name], [bd])
		if clientId != "Not found":
			print('client id: ' + clientId)
			inds, date, sampleCollectingSite, vs, tumorType, tnm = fgs.getRecord(credential, clientId)
			count = fgs.fetchReportCount(credential, inds)
			#clientInfo: [name, bd, clientId, gender, ind, twId, sampleCollectingSite, tnm]
			clientInfo = [name, bd, clientId, gender, inds[len(inds)-1], twId, sampleCollectingSite[0], vs[0], tumorType, tnm]
			print("inds: " + str(inds))
			print("date: " + str(date))
			print("count:" + str(count))
			# collection date 
			# x,y = fgs.getRecord(credential, clientId)
			# Create new threads
			thread1 = myThread(1, "Thread_getLineChart", [getLineChart, clientId, date.copy(), count.copy()])
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
				"\n\tName of Lab: " + sampleCollectingSite[0] + 
				"\n\tName of Physician: " + vs[0])
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
app.addLabelOptionBox("Templates", ["Traditional Chinese", "Simplified Chinese", "Captain", "Elite", "Shin Kong", "English"])
app.addDirectoryEntry("Save at")
app.addButtons(["Submit", "Cancel"], press)
app.enableEnter(press)
app.setFocus("Name")
app.go()
