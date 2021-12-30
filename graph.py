import requests
import msal 
import atexit
import os
import json 
import sys  
import pandas as pd
from shareplum import Site,Office365
from shareplum.site import Version


username = 'sushantshinde@orientindia.net'
password = 'Dilip@123'
sharepoint_url = 'https://orienttechnologies.sharepoint.com'
sharepoint_site = 'https://orienttechnologies.sharepoint.com/sites/sushant_ETL'
sharepoint_doc = 'https://orienttechnologies.sharepoint.com/:x:/s/QuikHr/EamixkIFFFBPi9m_QrWJrZEB7CqltE1eYw1Cj6sEe99Mcw?e=H6ujJW'
authcookie = Office365(sharepoint_url,username=username, password=password).GetCookies()
site = Site(sharepoint_site,version=Version.v365,authcookie=authcookie)
folder = site.Folder('Shared Documents/usage')

TENANT_ID = '5418abcf-d755-44c2-9ed7-7aac942abee7'
CLIENT_ID = '1f6f3d25-faf1-4870-b319-5797b2552ca9'
CLIENT_SECRET = 'qpH7Q~v.UUdB.PDxR8TFI0PznYUHSkKk2r8fO'

# TENANT_ID = 'fd0c8920-9100-4b4d-9013-4eabd6baa482'
# CLIENT_ID = 'ae343709-5a36-4635-8bb6-036888bad492'
# CLIENT_SECRET = 'srN7Q~hHqbPXp0_R4kYenZlwiqfO-FIhD3VV'

AUTHORITY = "https://login.microsoftonline.com/"+TENANT_ID

# AUTHORITY = "https://login.microsoftonline.com/consumers"
# AUTHORITY = "https://login.microsoftonline.com/consumers/",
ENDPOINT = "https://graph.microsoft.com/v1.0"
# 'qpH7Q~v.UUdB.PDxR8TFI0PznYUHSkKk2r8fO'
# v2Q7Q~DRNvT~AbPn3ZhHTtBI4tWd-z-IH-kfz

#ridhima srN7Q~hHqbPXp0_R4kYenZlwiqfO-FIhD3VV~
SCOPES = [
    'User.Read',
    'User.Read.All',
    'User.ReadWrite.All',
    'User.Invite.All',
    'User.Export.All',
    'Directory.ReadWrite.All',
    'Directory.Read.All',
    'Reports.Read.All'
]
scope = ['https://graph.microsoft.com/.default']
cache = msal.SerializableTokenCache()
print(cache)
if os.path.exists('token_cache.bin'):
    cache.deserialize(open('token_cache.bin','r').read())

atexit.register(lambda : open('token_cache.bin','w').write(cache.serialize()) if cache.has_state_changed else None)

app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)

# print(app)

accounts = app.get_accounts()
# print(accounts)
# app1 =  msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential = CLIENT_SECRET)
# result = app1.acquire_token_for_client(scope)
# print(result)

result = None
if len(accounts) > 0:
    result = app.acquire_token_silent(SCOPES, account=accounts[0])

if result is None:
    flow = app.initiate_device_flow(scopes=SCOPES)
   
    if 'user_code' not in flow:
        raise Exception('Failed to create device flow')
    # print(flow)
    print(flow['message'])

    sys.stdout.flush()  

    result = app.acquire_token_by_device_flow(flow)
    # print(result)

if 'access_token' in result:
    result1 = requests.get('https://graph.microsoft.com/v1.0/me', headers={'Authorization': 'Bearer ' + result['access_token']})
    result1.raise_for_status()
    # print(result1.json())

    # print(result['access_token'])
    headers ={'Authorization': 'Bearer ' + result['access_token'],'Content-Type':'application/json'}
#     data = {
#     "businessPhones": [
#     "+1 425 555 0109"
#   ],
#   "officeLocation": "18/2111"
# }
#     res = json.dumps(data)

#     update = requests.patch('https://graph.microsoft.com/v1.0/me',headers=headers,json=res)
#     update.raise_for_status()
#     print(update)
#     re = update.json()
#     print(res)

    users = requests.get('https://graph.microsoft.com/v1.0/users', headers={'Authorization': 'Bearer ' + result['access_token']})
    #print(users.json())
    #print(re.status_code)
    params = {
  "accountEnabled": True,
  "displayName": "zameer",
  "mailNickname": "zameer",
  "userPrincipalName": "zameer@s5fr.onmicrosoft.com",
  "passwordProfile" : {
    "forceChangePasswordNextSignIn": True,
    "password": "xWwvJ]6NMw+bWH-d"
  }
}
    # par = json.dumps(params)
    # create_users = requests.post('https://graph.microsoft.com/v1.0/users', headers=headers,data=par)
    # print(create_users.json())

    

    
  # period_value = "D30"
    update = requests.get("https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserDetail(period='D30')", headers=headers)
    boday = update.status_code
    data = update.text
    if update.status_code == 200:
        filePath = 'sharepoint.csv'
        with open(filePath, "wb") as f: 
            f.write(update.content)

        with open('/home/user/workspace/graph_api/sharepoint.csv','rb') as output:
            folder.upload_file(output, 'sharepoint.csv')
        # json_data = json.loads(update.text)
        # print(json_data)
    



# print(update.json())
# flow = app.initiate_device_flow(scopes=SCOPES)
# accounts = app.get_accounts()
# print(accounts)
# print(flow)
# result = app.acquire_token_by_device_flow(flow)
# print(result)
# admin@s5fr.onmicrosoft.com
# zeya Dojo5526\

# scre      c0iBrRtkti


# s5fr.onmicrosoft.com


# 5418abcf-d755-44c2-9ed7-7aac942abee7