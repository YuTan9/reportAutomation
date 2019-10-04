import pdfkit
path_wkthmltopdf = r'wkhtmltopdf\\bin\\wkhtmltopdf.exe'
# filename = input('file name: ')
filename = "CRC_EN_Template.html"
config = pdfkit.configuration(wkhtmltopdf=path_wkthmltopdf)
option = {
	'margin-top': '0mm',
	'margin-bottom':'0mm',
	'margin-left': '0mm',
	'margin-right': '0mm',
	'page-size': 'A4'
	}

pdfkit.from_file('CRC Monitor HTML\\' + filename, "CRC_Template.pdf", configuration=config, options = option)