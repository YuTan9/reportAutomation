"""
BEFORE RUNNING:
---------------
Install the Python client library for Google APIs by running
`pip install --upgrade google-api-python-client`
"""
from pprint import pprint
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient import discovery
import os.path
import pickle

def init():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    credentials = None

    if os.path.exists('token.pickle'):
    	with open('token.pickle', 'rb') as token:
    		credentials = pickle.load(token)

    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret_52165147212-mhva9d566esd0t9jvrvnh9g9ca9oh00c.apps.googleusercontent.com.json', SCOPES)
            credentials = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(credentials, token)
	
def printRange():
    service = discovery.build('sheets', 'v4', credentials=credentials)
    spreadsheet_id = '1Q6CgVm7u4oT4G0Hsc1gAqxBtOd4z9Auv34KCOG2ZRpc' 
    ranges = 'Report Count!A1:B12'  
    request = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=ranges)
    response = request.execute()
    pprint(response.get('values', []))

    # PS C:\Users\thoma\Desktop\Code\reportGenerator> python .\googlescripttest.py
    # [['TRF ID', 'Collection Date'],
    #  ['IND905675', '2017/10/03'],
    #  ['IND905675', '2017/10/03'],
    #  ['IND906234', '2017/11/13'],
    #  ['IND906234', '2017/11/13'],
    #  ['IND906335', '2018/01/25'],
    #  ['IND906335', '2018/01/25'],
    #  ['IND905818', '2018/03/22'],
    #  ['IND908693', '2018/08/22'],
    #  ['IND910609', '2019/01/03'],
    #  ['IND911398', '2019/03/21'],
    #  ['IND911665', '2019/07/25']]