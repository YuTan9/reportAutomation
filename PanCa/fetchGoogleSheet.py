"""
BEFORE RUNNING:
---------------
Install the Python client library for Google APIs by running
`pip install --upgrade google-api-python-client`
"""
# from pprint import pprint
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient import discovery
import os.path
import pickle
from google.auth.transport.requests import Request

def init():
	SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
	credentials = None

	if os.path.exists('token.pickle'):
		with open('token.pickle', 'rb') as token:
			credentials = pickle.load(token)
			return credentials

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
		return credentials
	
def printRange(credentials, ss, range):
	service = discovery.build('sheets', 'v4', credentials=credentials)
	spreadsheet_id = '1Q6CgVm7u4oT4G0Hsc1gAqxBtOd4z9Auv34KCOG2ZRpc' 
	ranges = str(ss) + "!" + range  
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

def getId(credentials, name, bd):
	# credentials: credentials
	# name: [name]
	# bd: [bd]
	service = discovery.build('sheets', 'v4', credentials=credentials, cache_discovery=False)
	spreadsheet_id = '1Q6CgVm7u4oT4G0Hsc1gAqxBtOd4z9Auv34KCOG2ZRpc' 
	ranges = "Client Info!B:B" 
	request = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=ranges)
	response = request.execute()
	names = response.get('values', [])

	service = discovery.build('sheets', 'v4', credentials=credentials, cache_discovery=False)
	spreadsheet_id = '1Q6CgVm7u4oT4G0Hsc1gAqxBtOd4z9Auv34KCOG2ZRpc' 
	ranges = "Client Info!D:D" 
	request = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=ranges)
	response = request.execute()
	bds = response.get('values', [])
	
	i = 0
	for i in range(len(names)):
		if names[i] == name:
			if bds[i] == bd:
				service = discovery.build('sheets', 'v4', credentials=credentials, cache_discovery=False)
				spreadsheet_id = '1Q6CgVm7u4oT4G0Hsc1gAqxBtOd4z9Auv34KCOG2ZRpc' 
				ranges = "Client Info!A" + str(i+1)
				request = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=ranges)
				response = request.execute()
				clientId = response.get('values', [])[0][0]

				ranges = "Client Info!C" + str(i+1)
				request = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=ranges)
				response = request.execute()
				gender = response.get('values', [])[0][0]

				ranges = "Client Info!E" + str(i+1)
				request = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=ranges)
				response = request.execute()
				twId = ''
				if response.get('values', []) != []:
					twId = response.get('values', [])[0][0]

				
				return clientId, gender, twId
			else:
				continue
		else:
			continue

	return "Not found", '', ''

def getRecord(credentials, clientId):
	service = discovery.build('sheets', 'v4', credentials=credentials, cache_discovery=False)
	spreadsheet_id = '1Q6CgVm7u4oT4G0Hsc1gAqxBtOd4z9Auv34KCOG2ZRpc' 
	ranges = "RUO with ID!A:A"
	request = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=ranges)
	response = request.execute()
	clientIds = response.get('values', [])
	inds = [] #id
	date = []
	sampleCollectingSite= ['']
	vs   = ['']
	for i in range(len(clientIds)):
		if clientIds[i] == [clientId]:
			service = discovery.build('sheets', 'v4', credentials=credentials, cache_discovery=False)
			spreadsheet_id = '1Q6CgVm7u4oT4G0Hsc1gAqxBtOd4z9Auv34KCOG2ZRpc' 

			ranges = "RUO with ID!C" + str(i+1)
			request = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=ranges)
			response = request.execute()
			row = response.get('values', [])[0][0]
			if discovery.build('sheets', 'v4', credentials=credentials, cache_discovery=False).spreadsheets().values().get(spreadsheetId=spreadsheet_id, range="RUO with ID!H" + str(i+1)).execute().get('values', [])[0][0] == "PanCA Monitor":
				inds.append(row)

			ranges = "RUO with ID!B" + str(i+1)
			request = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=ranges)
			response = request.execute()
			row = response.get('values', [])[0][0]
			if discovery.build('sheets', 'v4', credentials=credentials, cache_discovery=False).spreadsheets().values().get(spreadsheetId=spreadsheet_id, range="RUO with ID!H" + str(i+1)).execute().get('values', [])[0][0] == "PanCA Monitor":
				date.append(row)

			ranges = "RUO with ID!Q" + str(i+1)
			request = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=ranges)
			response = request.execute()
			if response.get('values', []) != []:
				sampleCollectingSite[0] = response.get('values', [])[0][0]

			ranges = "RUO with ID!S" + str(i+1)
			request = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=ranges)
			response = request.execute()
			if response.get('values', []) != []:
				vs[0] = response.get('values', [])[0][0]

			ranges = "RUO with ID!K" + str(i+1)
			request = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=ranges)
			response = request.execute()
			tumorType = ''
			if response.get('values', []) != []:
				tumorType = response.get('values', [])[0][0]

			ranges = "RUO with ID!L" + str(i+1)
			request = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=ranges)
			response = request.execute()
			tnm = ''
			if response.get('values', []) != []:
				tnm = response.get('values', [])[0][0]

	return [inds, date, sampleCollectingSite, vs, tumorType, tnm]

def fetchReportCount(credentials, inds):
	service = discovery.build('sheets', 'v4', credentials=credentials, cache_discovery=False)
	spreadsheet_id = '1D43UNSNqtMXGQ91OlKEXYwdyBXrF_nwogw81aeX6-I8' 

	trfRange      = "RUO Patient Data!B:B"
	trfRequest    = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=trfRange)
	trfResponse   = trfRequest.execute()
	trfId         = trfResponse.get('values', [])

	sampleIdRange = "RUO Patient Data!C:C"
	sampleRequest = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=sampleIdRange)
	sampleResponse= sampleRequest.execute()
	sampleId      = sampleResponse.get('values', [])

	countRange    = "RUO Patient Data!R:R"
	countRequest  = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=countRange)
	countResponse = countRequest.execute()
	count         = countResponse.get('values', [])

	print(inds)
	arr = []
	founds = []
	tumorInfo = []
	tnm = ''
	for j in range(0,len(inds)):
		flag = 0
		for i in range(len(trfId)):
			if trfId[i] == []:
				raise ValueError("Some value on Google sheet is empty.")
			if str(trfId[i][0]) == str(inds[j]):
				print("trfId: " + str(trfId[i][0]) + " ind: " + str(inds[j]))
				for found in founds:
					if found == sampleId[i][0]:
						flag = 1
						break
					else:
						continue
				if flag == 0:
					print(str(inds[j]) + "found with count: " + str(count[i][0]) + " sampleId: " + str(sampleId[i][0]))
					arr.append(int(count[i][0]))
					founds.append(sampleId[i][0])
					break
				elif flag ==1:
					flag = 0
					continue

				
	# print(arr)
	return arr

