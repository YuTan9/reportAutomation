# pack to exe: powershell run: auto-py-to-exe
# one file is not recommended, use one directory instead.
# include the additional files on the run (icon.png)
from appJar import gui
import fetchGoogleSheet
import pdfkit

def printPdfFromHtml():
	path_wkthmltopdf = r'wkhtmltopdf\\bin\\wkhtmltopdf.exe'
	config = pdfkit.configuration(wkhtmltopdf=path_wkthmltopdf)
	pdfkit.from_file('PanCA Monitor HTML\\template.html', 'report.pdf', configuration=config)

def press(button):
	if button == "Cancel":
		app.stop()
	else:
		print("1:", app.entry("Attribute1"), "2:", app.entry("Attribute2"))
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
