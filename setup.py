import os.system
os.system('curl https://bootstrap.pypa.io/get-pip.py -o get-pip.py')
os.system('python get-pip.py')
os.system('pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib')
