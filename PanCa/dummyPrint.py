import pdfkit
from bs4 import BeautifulSoup
import sys
template = str(sys.argv[1])

# printing html files to pdf with some text in blanks

path_wkthmltopdf = r'wkhtmltopdf\\bin\\wkhtmltopdf.exe'
config = pdfkit.configuration(wkhtmltopdf=path_wkthmltopdf)
option = {
	'margin-top': '0mm',
	'margin-bottom':'0mm',
	'margin-left': '0mm',
	'margin-right': '0mm',
	'page-size': 'A4'
	}
soup = BeautifulSoup(open(template, "r", encoding = "utf-8").read(), "html5lib")




soup.find(id='requisitionNumber').string        = "ABC123"
soup.find(id='patientName').string              = "abc甲乙丙"
soup.find(id='idNumber').string                 = "a123456789"
soup.find(id='dateOfBirth').string              = "2008/2/4"
soup.find(id='gender').string                   = '男'

soup.find(id='patientPhoneNumber').string       = '0945659209'
soup.find(id='patientEmail').string             = 'wqer@retfdg.wert'
soup.find(id='labPhoneNumber').string           = '8902347598'
soup.find(id='nameOfPhysician').string          = "John甲乙"
soup.find(id='dateOfCollection').string         = "2009/34/2"
soup.find(id='dateOfReport').string             = "2004/3/21"
#Second Page
soup.find(id='currentCtcCount').string          = "0"
soup.find(id='baselineCount').string            = "9"
soup.find(id='currentCtcTrend').string          = "Rising"

soup.find(id='comments').string                 = ''
soup.find(id='eSignDate1').string               = "2004/3/21"

if template != ".\\PanCA Monitor HTML\\USA_PanCA_EN_Template.html":
	soup.find(id='eSignDate2').string               = "2004/3/21"
	soup.find(id='nameOfLab').string                = "hospre"

soup.find(id='tumorType').string                = 'qwer'
soup.find(id='dateOfDiagnosis').string          = "2004/3/21"
soup.find(id='pathologicalDiagnosis').string    = ''
soup.find(id='tumorStage').string               = "甲乙丙"
soup.find(id='metastasisSite').string           = "2004/3/21"
soup.find(id='dateOfMostRecentTreatment').string= ''
soup.find(id='table004-B1').string              = "2004/3/21"
soup.find(id='table004-B2').string              = "5"

count = [1,2,3,4,5,6]
date = ["2004/3/21","2004/3/21","2004/3/21","2004/3/21","2004/3/21","2004/3/21"]
for i in range(len(count)):
	columnA = "table005-A" + str(2+i)
	columnB = "table005-B" + str(2+i)
	soup.find(id = columnA).string              = str(date[i])
	soup.find(id = columnB).string              = str(count[i])


with open("PanCA Monitor HTML\\htmlWithInfo.html", "w", encoding='utf-8') as file:
	file.write(str(soup))
# print(soup.prettify().encode('utf-8'))
pdfkit.from_file('PanCA Monitor HTML\\htmlWithInfo.html', "C:\\Users\\thoma\\Desktop\\dummy.pdf", configuration=config, options = option)