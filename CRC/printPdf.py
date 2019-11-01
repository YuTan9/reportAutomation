import pdfkit
import sys
# Print to pdf file with no blanks filled from a html file
path_wkthmltopdf = r'wkhtmltopdf\\bin\\wkhtmltopdf.exe'
filename = str(sys.argv[1])
config = pdfkit.configuration(wkhtmltopdf=path_wkthmltopdf)
option = {
	'margin-top': '0mm',
	'margin-bottom':'0mm',
	'margin-left': '0mm',
	'margin-right': '0mm',
	'page-size': 'A4'
	}

pdfkit.from_file(filename, "CRC_Template.pdf", configuration=config, options = option)