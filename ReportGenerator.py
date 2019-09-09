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


def getLineChart(id, x, y):
	if(len(x) != len(y)):
		return
	while len(x) > 6:
		x = x[1:]
	while len(y) > 6:
		y = y[1:]

	pseudoX = x.copy()
	pseudoY = y.copy()

	if len(x) < 6:
		pseudoX.append(' ')
		pseudoY.append(y[len(y)-1])

	print(x)
	print(pseudoX)
	print(y)
	print(pseudoY)

	plt.style.use('fast')
	fig = plt.figure(figsize = [7,7])
	ax = plt.axes()
	minimum = min(y)
	maximum = max(y)
	range_  = maximum - minimum
	ticks   = int(maximum / 8)+1 # a $ticks between every ticks

	ytick   = []
	for i in range(8):
		ytick.append(i*ticks)

	plt.xticks(range(len(pseudoX)), pseudoX)
	plt.yticks(ytick, ytick)
	plt.xlabel("Test Date")
	plt.ylabel("CTC Count")

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

	base = [y[0] for i in range(len(pseudoY))]
	plt.plot(pseudoX, base, '-', color='red', linewidth=2)

	plt.plot(x, y, '-', color='black', linewidth=2, markersize=1)
	for i in range(len(x)):
		plt.plot(x[i], y[i], 'o', color='black')
	print("Saving picture: " + str(id))
	fig.savefig("tmp.jpg",figsize=(7, 7), dpi=100)

def printPdfFromHtml():
	path_wkthmltopdf = r'wkhtmltopdf\\bin\\wkhtmltopdf.exe'
	config = pdfkit.configuration(wkhtmltopdf=path_wkthmltopdf)
	option = {
		'margin-top': '0mm',
		'margin-bottom':'0mm',
		'margin-left': '0mm',
		'margin-right': '0mm',
		'page-size': 'A4'
 	}
 	soup = BeautifulSoup(open('PanCA Monitor HTML\\PanCA Monitor Report_low TW.html', "r", encoding = "utf-8").read(), "html")
	pdfkit.from_file('PanCA Monitor HTML\\PanCA Monitor Report_low TW.html', 'report.pdf', configuration=config, options = option)

def press(button):
	if button == "Cancel":
		app.stop()
	else:
		print("Name:", app.entry("Name"), "Birthday:", app.entry("Birthday"))
		app.hide()
		name = str(app.entry("Name"))
		bd = str(app.entry("Birthday"))
		credential = fgs.init()
		clientId = fgs.getId(credential, [name], [bd])
		print('client id: ' + clientId)
		if clientId != "Not found":
			inds, date = fgs.getRecord(credential, clientId)
			count = fgs.fetchReportCount(credential, inds)
			# print(inds)
			print(date)
			print(count)
			# collection date 
			# x,y = fgs.getRecord(credential, clientId)
			# Create new threads
			thread1 = myThread(1, "Thread_getLineChart", [getLineChart, clientId, date, count])
			thread2 = myThread(2, "Thread_printPdfFromHtml", [printPdfFromHtml])


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
		else:
			app.infoBox("Not found", "The information you entered is not found.")
			app.clearAllEntries(callFunction=False)
		app.show()




def executeThreadFunc(args):
	if args[0] == getLineChart:
		getLineChart(args[1], args[2], args[3])
	elif args[0] == printPdfFromHtml:
		printPdfFromHtml()
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
		except:
			print("Error raised in " + self.name + " by " + str(self.func) + ".")
		# Free lock to release next thread
		threadLock.release()
		print(self.name + " released lock.")



###############
#MAIN FUNCTION#
###############
threadLock = threading.Lock()
threads = []
app = gui("CellmaxLife Report Generator", "400x200")
app.setBg("white")
app.setIcon("icon.gif")
app.setFont(12)
app.addLabelEntry("Name")
app.addLabelEntry("Birthday")
app.setEntryDefault("Name", "Enter Name")
app.setEntryDefault("Birthday", "Birthday in MM/DD/YYYY")
app.addButtons(["Submit", "Cancel"], press)
app.setFocus("Name")
app.go()
