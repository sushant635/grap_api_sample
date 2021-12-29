from O365 import Account
from configparser import ConfigParser
from ms_graph.client import MicrosoftGraphClient
from pprint import pprint


# credentials = ('1f6f3d25-faf1-4870-b319-5797b2552ca9','qpH7Q~v.UUdB.PDxR8TFI0PznYUHSkKk2r8fO')

# account = Account(credentials)
# print('account details',account)
# m = account.new_message()
# print(m)
# # m.to.add('sushantshinde@orintindia.net')
# # m.subject = 'Testing!'
# # m.body = "George Best quote: I've stopped drinking, but only while I'm asleep."
# # m.send()




scopes = [
    'User.Read',
    'User.Read.All',
    'User.ReadWrite.All',
    'User.Invite.All',
    'User.Export.All',
    'Directory.ReadWrite.All',
    'Directory.Read.All',
    'Reports.Read.All'
]
# Initialize the Parser.
config = ConfigParser()

# Add the Section.
config.add_section('graph_api')

# Set the Values.
config.set('graph_api', 'client_id', '1f6f3d25-faf1-4870-b319-5797b2552ca9')
config.set('graph_api', 'client_secret', 'qpH7Q~v.UUdB.PDxR8TFI0PznYUHSkKk2r8fO')
config.set('graph_api', 'redirect_uri', 'http://localhost:8000/callback')

# Write the file.
with open(file='token_cache.bin', mode='w+') as f:
    config.write(f)
print(config)

# Initialize the Client.
graph_client = MicrosoftGraphClient(
    client_id=client_id,
    client_secret=client_secret,
    redirect_uri=redirect_uri,
    scope=scopes,
    credentials='config/ms_graph_state.jsonc'
)

# Login to the Client.
graph_client.login()


# Grab the User Services.
user_services = graph_client.users()