import requests, sys, json, re, urllib, os

### Update these values ###
CACHE_LOCATION = 'C:\ProgramData\M365_Notify'
TENANT_ID = '' # Tenant UUID
CLIENT_ID = '' # Client ID (from Azure AD App Registration)
CHANNELS = [
    # List of links to Teams channels
    'https://teams.microsoft.com/l/channel/.../General?groupId=...&tenantId=...'
]
###########################

GRAPH_API = 'https://graph.microsoft.com/v1.0'

if not os.path.exists(CACHE_LOCATION): os.mkdir(CACHE_LOCATION)

try:
    fp = open(CACHE_LOCATION+'\cache.json', 'r')
    cache = json.load(fp)
    fp.close()
except: cache = {}

def write_cache():
    with open(CACHE_LOCATION+'\cache.json', 'w') as fp: json.dump(cache, fp, indent = 4)

if 'refresh_token' not in cache:
    print('Please retrieve an authorization code from')
    print("\033[36m"+'https://login.microsoft.com/%s/oauth2/v2.0/authorize?'%(TENANT_ID)+'&'.join(
            ['%s=%s'%(key, urllib.parse.quote(value, safe='')) for key, value in {
                'client_id': CLIENT_ID,
                'response_type': 'code',
                'redirect_uri': 'https://login.microsoftonline.com/common/oauth2/nativeclient',
                'scope': 'offline_access ChannelMessage.Send ServiceHealth.Read.All',
                'response_mode': 'query'
            }.items()]
        )+"\033[0m"
    )
    print('and paste the redirect URL below.')
    redirect_url = input('URL: ')
    code = re.search('code=([^&]+)', redirect_url).group(1)
    token_request = requests.post(
        'https://login.microsoft.com/%s/oauth2/v2.0/token'%(TENANT_ID),
        data = {
            'client_id': CLIENT_ID,
            'grant_type': 'authorization_code',
            'scope': 'offline_access ChannelMessage.Send ServiceHealth.Read.All',
            'code': code,
            'redirect_uri': 'https://login.microsoftonline.com/common/oauth2/nativeclient'
        }
    )
    if 'refresh_token' not in token_request.json(): sys.exit('Authorization failed')
    cache['refresh_token'] = token_request.json()['refresh_token']
    write_cache()

access_request = requests.post(
    'https://login.microsoft.com/%s/oauth2/v2.0/token'%(TENANT_ID),
    data = {
        'client_id': CLIENT_ID,
        'grant_type': 'refresh_token',
        'scope': 'offline_access ChannelMessage.Send ServiceHealth.Read.All',
        'refresh_token': cache['refresh_token'],
        'redirect_uri': 'https://login.microsoftonline.com/common/oauth2/nativeclient'
    }
)
access_token = access_request.json()['access_token']

issue_request = requests.get(
    '%s/admin/serviceAnnouncement/issues?$filter=IsResolved%%20eq%%20false'%(GRAPH_API),
    headers = {'Authorization': 'Bearer %s'%(access_token)}
)
if 'value' not in issue_request.json(): sys.exit('Graph API unavailable')

for channel in CHANNELS:

    channel_id, team_id = re.search('channel/([^/]+)/.+groupId=([0-9a-z\-]+)', channel).groups()
    if channel_id not in cache: cache[channel_id] = {}

    for issue_id in cache[channel_id].keys():
        if issue_id not in [issue['id'] for issue in issue_request.json()['value']]: 
            requests.post(
                '%s/teams/%s/channels/%s/messages/%s/replies'%(GRAPH_API, team_id, channel_id, cache[channel_id][issue_id]['message_id']),
                headers = {'Authorization': 'Bearer %s'%(access_token)},
                json = {
                    'body': {
                        'content': '<i>Issue was closed.</i>',
                        'contentType': 'html'
                    }
                }
            )
            cache[channel_id].pop(issue_id)

    for issue in issue_request.json()['value']:

        if issue['id'] not in cache[channel_id]:
            new_message_request = requests.post(
                '%s/teams/%s/channels/%s/messages'%(GRAPH_API, team_id, channel_id),
                headers = {'Authorization': 'Bearer %s'%(access_token)},
                json = {
                    'subject': '%s - %s'%(issue['id'], issue['title']),
                    'body': {
                        'content': issue['impactDescription'],
                        'contentType': 'html'
                    },
                    'importance': 'high' if issue['classification'] == 'incident' else 'normal'
                }
            )
            cache[channel_id][issue['id']] = {'message_id': new_message_request.json()['id'], 'updates': []}

        for post in issue['posts']:
            if post['createdDateTime'] not in cache[channel_id][issue['id']]['updates']:
                requests.post(
                    '%s/teams/%s/channels/%s/messages/%s/replies'%(GRAPH_API, team_id, channel_id, cache[channel_id][issue['id']]['message_id']),
                    headers = {'Authorization': 'Bearer %s'%(access_token)},
                    json = {
                        'body': {
                            'content': re.search('Current\sstatus:\s(.*)', post['description']['content']).group(1),
                            'contentType': 'html'
                        }
                    }
                )
                cache[channel_id][issue['id']]['updates'].append(post['createdDateTime'])

write_cache()
