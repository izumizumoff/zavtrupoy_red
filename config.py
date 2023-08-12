import json

token = None
calendar = None
host = None
username = None
password = None
known_hosts = None
rsa_key = None

with open('.settings/path.json', 'r') as file:
    PATH = json.load(file)

# for google token
with open(PATH, 'r') as secret:
    exec(secret.read())
TOKEN = token
# id google calendar
CALENDAR = calendar
SFTP_HOST = host
SFTP_USERNAME = username
SFTP_PASSWORD = password
KNOWN_HOSTS = known_hosts
RSA_KEY = rsa_key