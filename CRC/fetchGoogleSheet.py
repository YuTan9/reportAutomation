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
	ranges = "Sample Reception with ID!A:A"
	request = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=ranges)
	response = request.execute()
	clientIds = response.get('values', [])
	inds = [] #id
	date = []
	for i in range(len(clientIds)):
		if clientIds[i] == [clientId]:
			service = discovery.build('sheets', 'v4', credentials=credentials, cache_discovery=False)
			spreadsheet_id = '1Q6CgVm7u4oT4G0Hsc1gAqxBtOd4z9Auv34KCOG2ZRpc' 

			ranges = "Sample Reception with ID!B" + str(i+1)
			request = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=ranges)
			response = request.execute()
			row = response.get('values', [])[0][0]
			inds.append(row)

			ranges = "Sample Reception with ID!H" + str(i+1)
			request = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=ranges)
			response = request.execute()
			row = response.get('values', [])[0][0]
			arr = row.split("/")
			row = str(arr[2]) + str("/") + str(arr[0]) + str("/") + str(arr[1]) 
			date.append(row)


	return [inds, date]

def fetchReportCount(credentials, inds):
	service = discovery.build('sheets', 'v4', credentials=credentials, cache_discovery=False)
	spreadsheet_id = '1vjTJ5e8ElREm6NHx2qQuhGuqUwwiREaFjMG3L12L8og' 

	trfRange      = "Report Generation 2!A:A"
	trfRequest    = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=trfRange)
	trfResponse   = trfRequest.execute()
	trfId         = trfResponse.get('values', [])

	# this is a modified version from ruo report generation, which sample id was used to distinguish different tests with
	# same ind sequence. whereas test date is used to distinguish different tests on CRC report generation. 
	sampleIdRange = "Report Generation 2!B:B"
	sampleRequest = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=sampleIdRange)
	sampleResponse= sampleRequest.execute()
	sampleId      = sampleResponse.get('values', [])

	countRange    = "Report Generation 2!S:S"
	countRequest  = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=countRange)
	countResponse = countRequest.execute()
	count         = countResponse.get('values', [])

	scsRange      = "Report Generation 2!Q:Q"
	scsRequest    = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=scsRange)
	scsResponse   = scsRequest.execute()
	scs           = scsResponse.get('values', [])

	telRange      = "Report Generation 2!F:F"
	telRequest    = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=telRange)
	telResponse   = telRequest.execute()
	tel           = telResponse.get('values', [])

	emailRange    = "Report Generation 2!J:J"
	emailRequest  = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=emailRange)
	emailResponse = emailRequest.execute()
	emails        = emailResponse.get('values', [])

	vsRange       = "Report Generation 2!M:M"
	vsRequest     = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=vsRange)
	vsResponse    = vsRequest.execute()
	vss           = vsResponse.get('values', [])

	arr = []
	founds = []
	sampleCollectingSite = ''
	telephone = ''
	email = ''
	vs = ''
	for j in range(0,len(inds)):
		flag = 0
		for i in range(len(trfId)):
			if trfId[i] == []:
				raise ValueError("Some value on Google sheet is empty.")
			if str(trfId[i][0]) == str(inds[j]):
				for found in founds:
					if found == sampleId[i][0]:
						flag = 1
						break
				if flag == 0:
					print(str(inds[j]) + "found with count: " + str(count[i][0]) + " sampleId: " + str(sampleId[i][0]))
					arr.append(int(count[i][0]))
					founds.append(sampleId[i][0])
					if j == len(inds)-1:
						if scs[i] != []:
							sampleCollectingSite = scs[i][0]
						if tel[i] != []:
							telephone = tel[i][0]
						if emails[i] != []:
							email = emails[i][0]
						if vss[i] != []:
							vs = vss[i][0]
					break
				elif flag ==1:
					flag = 0
					continue

	return arr, sampleCollectingSite, telephone, email, vs

def printRange(credentials, ss, range):
	service = discovery.build('sheets', 'v4', credentials=credentials)
	spreadsheet_id = '1Q6CgVm7u4oT4G0Hsc1gAqxBtOd4z9Auv34KCOG2ZRpc' 
	ranges = str(ss) + "!" + range  
	request = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=ranges)
	response = request.execute()
	pprint(response.get('values', []))
