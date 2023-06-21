import requests
from requests_ntlm import HttpNtlmAuth

# SharePoint site URL
site_url = "https://bakerhughes.sharepoint.com/sites/DTNightsWatch/"

# SharePoint API endpoint URL
endpoint_url = site_url + "_api/web/lists/getbytitle('Documents')/items"

# Active Directory username and password
username = "TIBERIU.ROCIU@bakerhughes.com"
password = "qweasdzxcQAZ724!?"

# Authenticate with the SharePoint site using NTLM authentication
session = requests.Session()
session.auth = HttpNtlmAuth(username, password, session)

# Make a request to the SharePoint API endpoint
response = session.get(endpoint_url)

# Check if the request was successful
if response.status_code == requests.codes.ok:
    # Process the response data
    data = response.json()
    for item in data['d']['results']:
        print(item['Title'])
else:
    # Print the error message
    response.raise_for_status()
