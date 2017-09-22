import requests
from lxml import html
import httplib2
import re
import os
import json
from sys import flags

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import time

finder = re.compile("toastr\.warning\(\"(.*?)\",")

s = None
csrftoken = None
username = ''
password = ''

def create_session():
    global s, csrftoken, username, password
    s = requests.Session()
    r = s.get("https://think-infinity.thon.org/int/solicitation/canning/location-registration/")
    r = s.post("https://webaccess.psu.edu/", {"login": username, 
                                              "password": password,
                                              "service": "cosign-think-infinity.thon.org",
                                              "required": "",
                                              "ref": "https://think-infinity.thon.org/int/solicitation/canning/location-registration/"})
                                              
    content = html.fromstring(r.content)
    csrftoken = content.find(".//input[@name='csrfmiddlewaretoken']").value

def register_location(trip, name, addr_line_1, addr_line_2, city, state, zip):
    global s, csrftoken

    cookie = s.cookies

    s.headers.update({'referer': "https://think-infinity.thon.org/int/solicitation/canning/location-registration/"})

    r = s.post("https://think-infinity.thon.org/int/solicitation/canning/location-registration/", {
               "csrfmiddlewaretoken": csrftoken,
               "canning_trip": trip,
               "location_name": name,
               "loc_addr_one": addr_line_1,
               "loc_addr_two": addr_line_2,
               "loc_addr_city": city,
               "loc_addr_state": state,
               "loc_addr_zip": zip,
               "add_location": "Add Location"
                })
                
    results = finder.findall(str(r.content))
    if results:
        print(trip)
        print(name)
        print(addr_line_1)
        print(addr_line_2)
        print(city)
        print(state)
        print(zip)
        print(results)
        print("====================================")
        raise Exception(str(results))
    
def get_credentials():
    credential_dir = os.path.join('./', '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, 'sheets.googleapis.com-canning.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, tools.argparser.parse_args([]))
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials
                

SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = './client_secret.json'
APPLICATION_NAME = 'Canning Location Registration'

credentials = get_credentials()
http = credentials.authorize(httplib2.Http())
discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                'version=v4')
service = discovery.build('sheets', 'v4', http=http,
                                discoveryServiceUrl=discoveryUrl)
                                
spreadsheetId = '1V_NZelZmLIghOsOXY0UXIADZz3YDXw9CH8vqhBM1opk'
rangeName = 'Form Responses 1!A2:J'


trip_id_regex = re.compile("[a-zA-Z\-\s,]+\(([0-9]+)\)")


with open('credentials.txt', 'r') as f:
    username = f.readline().strip()
    password = f.readline().strip()

while True:

    try:
        s = None

        result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
        values = result.get('values', [])

        for i, v in enumerate(values):
            if len(v) == 9:
                if s == None:
                    create_session()
            
                trip_id = int(trip_id_regex.match(v[2]).group(1))
                
                try:
                    register_location(trip=trip_id, name=v[3], addr_line_1=v[4], addr_line_2=v[5], city=v[6], state=v[7], zip=v[8])
                    values = ['Registered']
                except Exception as e:
                    values = ['ERROR: {}'.format(e)]
                    
                request = service.spreadsheets().values().update(spreadsheetId=spreadsheetId, range="Form Responses 1!J{}".format(i+2), body={
                                                                                                                                      "range": "Form Responses 1!J{}".format(i+2),
                                                                                                                                      "values": [
                                                                                                                                        values
                                                                                                                                      ],
                                                                                                                                    },
                                                        valueInputOption="USER_ENTERED")
                response = request.execute()
            
    except Exception as e:
        print(e)
    time.sleep(10)