import requests
import msal 
import atexit
import os
import json 
import sys  
import pandas as pd
import numpy as np
from shareplum import Site,Office365
from shareplum.site import Version
from datetime import *


username = 'sushantshinde@orientindia.net'
password = 'Dilip@123'
sharepoint_url = 'https://orienttechnologies.sharepoint.com'
sharepoint_site = 'https://orienttechnologies.sharepoint.com/sites/sushant_ETL'
sharepoint_doc = 'https://orienttechnologies.sharepoint.com/:x:/s/QuikHr/EamixkIFFFBPi9m_QrWJrZEB7CqltE1eYw1Cj6sEe99Mcw?e=H6ujJW'
authcookie = Office365(sharepoint_url,username=username, password=password).GetCookies()
site = Site(sharepoint_site,version=Version.v365,authcookie=authcookie)
folder = site.Folder('Shared Documents/usage')


# TENANT_ID = '5418abcf-d755-44c2-9ed7-7aac942abee7'
# CLIENT_ID = '1f6f3d25-faf1-4870-b319-5797b2552ca9'
# CLIENT_SECRET = 'qpH7Q~v.UUdB.PDxR8TFI0PznYUHSkKk2r8fO'

TENANT_ID = 'fd0c8920-9100-4b4d-9013-4eabd6baa482'
CLIENT_ID = '40bcf090-8cb0-4998-bf24-50bdd79f7946'
CLIENT_SECRET = '06a88177-220d-4e40-8ce5-442986aedc55'

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



def is_month(data_str):
    new_date = datetime.strptime(data_str,"%Y-%m-%d").date()
    today_date = datetime.today()
    return (new_date.month == today_date.month) and (new_date.year == today_date.year)

def isNaN(num):
    return num!= num

if os.path.exists('token_cache_data.bin'):
    cache.deserialize(open('token_cache_data.bin','r').read())

atexit.register(lambda : open('token_cache_data.bin','w').write(cache.serialize()) if cache.has_state_changed else None)

app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)


accounts = app.get_accounts()
print(accounts)


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
    print('working the data')

    headers ={'Authorization': 'Bearer ' + result['access_token'],'Content-Type':'application/json'}

    update = requests.get("https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserDetail(period='D30')", headers=headers)
    
    if update.status_code == 200:
        filePath = 'sharepoint.csv'
        with open(filePath, "wb") as f: 
            f.write(update.content)

        
        df = pd.read_csv('sharepoint.csv')
        # print(df['Exchange Last Activity Date'])
        exchange_list = []
        for i in df['Exchange Last Activity Date']:
            if isNaN(i) != True:
                temp = is_month(i)
                exchange_list.append(temp)
            else:
                temp = False
                exchange_list.append(temp)


        onedrive_list = []
        for i in df['OneDrive Last Activity Date']:
            if isNaN(i) != True:
                temp = is_month(i)
                onedrive_list.append(temp)
            else:
                temp = False
                onedrive_list.append(temp)
        
        sharepoint_list = []
        for i in df['SharePoint Last Activity Date']:
            if isNaN(i) != True:
                temp = is_month(i)
                sharepoint_list.append(temp)
            else:
                temp = False
                sharepoint_list.append(temp)

        skype_list = []
        for i in df['Skype For Business Last Activity Date']:
            if isNaN(i) != True:
                temp = is_month(i)
                skype_list.append(temp)
            else:
                temp = False
                skype_list.append(temp)

        teams_list = []
        for i in df['Teams Last Activity Date']:
            if isNaN(i) != True:
                temp = is_month(i)
                teams_list.append(temp)
            else:
                temp = False
                teams_list.append(temp)

        df['Exchange Last Activity Date Flag'] = exchange_list
        df['OneDrive Last Activity Date Flag'] = onedrive_list
        df['SharePoint Last Activity Date Flag'] = sharepoint_list
        df['Skype For Business Last Activity Date Flag'] = skype_list
        df['Teams Last Activity Date Flag'] = teams_list 


        df.to_csv('all_data.csv',index=False)

        with open('/home/user/workspace/graph_api/all_data.csv','rb') as output:
            folder.upload_file(output, 'all_data.csv')
        
# # for activation user 
#     activations = requests.get("https://graph.microsoft.com/v1.0/reports/getOffice365ActivationsUserDetail",headers=headers)
    
#     if activations.status_code == 200:
#         with open('activations.csv',"wb") as f:
#             f.write(activations.content)

# #for teams data
#     teams_details = requests.get("https://graph.microsoft.com/v1.0/reports/getTeamsDeviceUsageUserDetail(period='D30')",headers=headers)
#     if teams_details.status_code == 200:
#         with open('teams_details.csv',"wb") as f:
#             f.write(teams_details.content)


#     outlook_usage = requests.get("https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='D30')",headers=headers)
#     if outlook_usage.status_code == 200:
#         with open('outlook_usage.csv',"wb") as f:
#             f.write(outlook_usage.content)

#     onedrive_usage = requests.get("https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='D30')",headers=headers)
#     if onedrive_usage.status_code == 200:
#         with open('onedrive_usage.csv',"wb") as f:
#             f.write(onedrive_usage.content)

#     sharepoint_usage = requests.get("https://graph.microsoft.com/v1.0/reports/getSharePointActivityUserDetail(period='D30')",headers=headers)
#     if sharepoint_usage.status_code == 200:
#         with open('sharepoint_usage.csv',"wb") as f:
#             f.write(sharepoint_usage.content)


#     skype_usage = requests.get("https://graph.microsoft.com/v1.0/reports/getSkypeForBusinessDeviceUsageUserDetail(period='D30')",headers=headers)
#     if skype_usage.status_code == 200:
#         with open('skype_usage.csv','wb') as f:
#             f.write(skype_usage.content) 
    
        
    



    


    

