# Report Automation

A project at Cellmax Life for generating reports by fetching data from google sheets and filling into html templates.

## Execution

Run setup.py to install the packages needed
```
python setup.py
```
Run the main funciton (make sure all the sheets are accessible)
```
python ReportGenerator.py
```


## Built With

* [appJar](http://appjar.info/) - The user interface used
* [Google APIs](https://developers.google.com/sheets/api/quickstart/python) - Used for data retrieval
* [PDFKit](https://pdfkit.org/) - Used to print html to pdf
* [Beautiful Soup](https://www.crummy.com/software/BeautifulSoup/bs4/doc/) - Used for filling in data to specific id in html files
* [Matplotlib](https://matplotlib.org/) - Used for drawing the CTC count trend in the report

## Authors

* **Yu-Tang, Shen** - *Initial work*
