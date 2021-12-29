import requests
import msal 
import jwt
import json 
from datetime import datetime

accessToken = None 
requestHeaders = None 
tokenExpiry = None 
queryResults = None 
graphURI = 'https://graph.microsoft.com'


def msgraph_auth():
    global accessToken
    global requestHeaders
    global tokenExpiry

    TENANT_ID = 'f8cdef31-a31e-4b4a-93e4-5f571e91255a'
    # authority = 'https://login.microsoftonline.com/' + TENANT_ID
    authority = "https://login.microsoftonline.com/consumers"
    CLIENT_ID = '587e3c4f-cef5-4fc2-9287-6210928a91aa'
    CLIENT_SECRET = 'qpH7Q~v.UUdB.PDxR8TFI0PznYUHSkKk2r8fO'
    scope = ['https://graph.microsoft.com/.default']
    # scope = [
    #     'User.Read',
    #     'User.ReadWrite',
    #     'User.ReadBasic.All'
    # ]

    app = msal.ConfidentialClientApplication(CLIENT_ID, authority=authority, client_credential = CLIENT_SECRET)

    print(app)

    try:
        accessToken = app.acquire_token_silent(scope, account=None)
        print(accessToken)
        if not accessToken:
            try:
                accessToken = app.acquire_token_for_client(scopes=scope)
                print(accessToken)
                if accessToken['access_token']:
                    print('New access token retreived....')
                    requestHeaders = {'Authorization': 'Bearer ' + accessToken['access_token']}
                else:
                    print('Error aquiring authorization token. Check your tenantID, clientID and clientSecret.')

            except:
                pass

        else:
            print('Token retreived from MSAL Cache....')
        
        decodedAccessToken = jwt.decode(accessToken['access_token'], verify=False)
        accessTokenFormatted = json.dumps(decodedAccessToken, indent=2)
        print('Decoded Access Token')
        print(accessTokenFormatted)

        # Token Expiry
        tokenExpiry = datetime.fromtimestamp(int(decodedAccessToken['exp']))
        print('Token Expires at: ' + str(tokenExpiry))
        return
    except Exception as err:
        print(err)
    

def msgraph_request(resource,requestHeaders):
    # Request
    results = requests.get(resource, headers=requestHeaders).json()
    return results

msgraph_auth()

print(graphURI +'/v1.0/me',requestHeaders)
queryResults = msgraph_request(graphURI +'/v1.0/me',requestHeaders)

try:
    df = pd.read_json(json.dumps(queryResults['value']))
    # set ID column as index
    df = df.set_index('id')
    print(str(df['displayName'] + " " + df['mail']))

except:
    print(json.dumps(queryResults, indent=2))
