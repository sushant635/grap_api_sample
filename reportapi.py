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

AUTHORITY = "https://login.microsoftonline.com/"+TENANT_ID

ENDPOINT = "https://graph.microsoft.com/v1.0"


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


if os.path.exists('token_cache_data.bin'):
    cache.deserialize(open('token_cache_data.bin','r').read())

atexit.register(lambda : open('token_cache_data.bin','w').write(cache.serialize()) if cache.has_state_changed else None)

app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)

accounts = app.get_accounts()


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


if 'access_token' in result:

    headers ={'Authorization': 'Bearer ' + result['access_token'],'Content-Type':'application/json'}

    update = requests.get("https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserDetail(period='D30')", headers=headers)
    
    if update.status_code == 200:
        filePath = 'sharepoint.csv'
        with open(filePath, "wb") as f: 
            f.write(update.content)

        with open('/home/user/workspace/graph_api/sharepoint.csv','rb') as output:
            folder.upload_file(output, 'sharepoint.csv')
# for activation user 
    activations = requests.get("https://graph.microsoft.com/v1.0/reports/getOffice365ActivationsUserDetail",headers=headers)
    
    if activations.status_code == 200:
        with open('activations.csv',"wb") as f:
            f.write(activations.content)

#for teams data
    teams_details = requests.get("https://graph.microsoft.com/v1.0/reports/getTeamsDeviceUsageUserDetail(period='D30')",headers=headers)
    if teams_details.status_code == 200:
        with open('teams_details.csv',"wb") as f:
            f.write(teams_details.content)


    outlook_usage = requests.get("https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='D30')",headers=headers)
    if outlook_usage.status_code == 200:
        with open('outlook_usage.csv',"wb") as f:
            f.write(outlook_usage.content)

    onedrive_usage = requests.get("https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='D30')",headers=headers)
    if onedrive_usage.status_code == 200:
        with open('onedrive_usage.csv',"wb") as f:
            f.write(onedrive_usage.content)

    sharepoint_usage = requests.get("https://graph.microsoft.com/v1.0/reports/getSharePointActivityUserDetail(period='D30')",headers=headers)
    if sharepoint_usage.status_code == 200:
        with open('sharepoint_usage.csv',"wb") as f:
            f.write(sharepoint_usage.content)


    skype_usage = requests.get("https://graph.microsoft.com/v1.0/reports/getSkypeForBusinessDeviceUsageUserDetail(period='D30')",headers=headers)
    if skype_usage.status_code == 200:
        with open('skype_usage.csv','wb') as f:
            f.write(skype_usage.content) 
    
        
    



    


    

