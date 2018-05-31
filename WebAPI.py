import requests
import json

crmorg = 'https://tfkcdev.crm.dynamics.com'
clientid = '482acc8b-71bb-41e8-8066-7f7263fea4fb'
username = 'sysadmin@tfkc.onmicrosoft.com'
password = 'P1ssw1rd'
tokenendpoint = 'https://login.microsoftonline.com/482acc8b-71bb-41e8-8066-7f7263fea4fb/oauth2/token'

crmwebapi = 'https://tfkcdev.api.crm.dynamics.com/api/data/v9.0/'

tokenpost = {
    'client_id': clientid,
    'resource': crmorg,
    'username': username,
    'password': password,
    'grant_type' : 'password'
}

tokenres = requests.post(tokenendpoint, data=tokenpost)

accesstoken = ''
print(tokenres.json)

accesstoken = tokenres.json()['access_token']
