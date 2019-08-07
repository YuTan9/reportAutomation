# pack to exe: powershell run: auto-py-to-exe
# one file is not recommended, use one directory instead.
# include the additional files on the run (icon.png)
from appJar import gui
import fetchGoogleSheet
import pdfkit
import matplotlib.pyplot as plt

def getLineChart():
	x = ['2018-10-11','2018-10-12', '2018-10-13','2018-10-14','2018-10-15','2018-10-16']
	y = [1,2,3,5,2,7]
	plt.style.use('bmh')
	fig = plt.figure()
	ax = plt.axes()
	plt.xticks(range(6), ['2018-10-11','2018-10-12', '2018-10-13','2018-10-14','2018-10-15','2018-10-16'])
	plt.yticks([1,2,3,4,5,6,7,8,9,10], ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10'])
	plt.xlabel("Test Date")
	plt.ylabel("CTC Count")
	plt.plot(x,y, '.-')
	plt.show()
	fig.savefig("tmp.jpg")

def printPdfFromHtml():
	path_wkthmltopdf = r'wkhtmltopdf\\bin\\wkhtmltopdf.exe'
	config = pdfkit.configuration(wkhtmltopdf=path_wkthmltopdf)
	# config.default_options = {
	# 	margin_top: '0cm',
	#     margin_right: '0cm',
	#     margin_left: '0cm',
	#     margin_bottom: '0cm',

	#     # quiet: true, # No output during PDF generation
	#     # load_error_handling: 'abort', # Crash early
	#     # load_media_error_handling: 'abort', # Crash early
	#     # no_outline: true, # Disable the default outline
	#     # # disable_smart_shrinking: true, # Enable to keep the pixel/dpi ratio linear
 #    }
	option = {
		'margin-top': '0mm',
		'margin-bottom':'0mm',
		'margin-left': '0mm',
		'margin-right': '0mm',
		'page-size': 'A4'
 	}
	pdfkit.from_file('PanCA Monitor HTML\\PanCA Monitor Report_low TW.html', 'report.pdf', configuration=config, options = option)

def press(button):
	if button == "Cancel":
		app.stop()
	else:
		print("1:", app.entry("Attribute1"), "2:", app.entry("Attribute2"))
		# getLineChart()
		printPdfFromHtml()



app = gui("CellmaxLife Report Generator", "400x200")
app.setBg("white")
app.setIcon("icon.gif")
app.setFont(12)
app.addLabelEntry("Attribute1")
app.addLabelEntry("Attribute2")
app.addButtons(["Submit", "Cancel"], press)
app.setFocus("Attribute1")
app.go()
