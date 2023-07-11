from shareplum import Site
from shareplum import Office365
from requests_ntlm import HttpNtlmAuth
import os

# SharePoint site credentials
username = "TIBERIU.ROCIU@bakerhughes.com"
password = "qweasdzxcQAZ724!?"
site_url = "https://bakerhughes.sharepoint.com/sites/DTNightsWatch/"

# Authentication
auth = HttpNtlmAuth(username, password)
authcookie = Office365(site_url, username=username, password=password).GetCookies()

# SharePoint site and document library
site = Site(site_url, auth=authcookie, auth_ntlm=auth)
doc_lib = site.List('Shared Documents')

# File to download
file_name = "Project Status - CCD to DocuSign.xlsx"
file_path = os.path.join(os.getcwd(), file_name)

# Iterate over documents and download the file
for doc in doc_lib.GetListItems():
    if doc['Name'] == file_name:
        with open(file_path, 'wb') as f:
            f.write(doc_lib.GetAttachment(doc['ID'], file_name))
        print(f"File downloaded to {file_path}")
        break
else:
    print(f"{file_name} not found in SharePoint document library")