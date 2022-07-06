from __future__ import print_function

import os.path
import base64
import csv
from datetime import datetime

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def addToCsv(date, data):
    """Adds data to a csv file"""
    del data['Tell us about yourself']
    del data['County']
    del data['About your organization']
    del data['Tell us about your Reserve planning history']

    if not os.path.exists('data.csv'): 
        with open('data.csv', 'w') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['Date', *[d.strip() for d in data.keys()]])
            writer.writerow([date, *[d.strip() for d in data.values()]])
    else:
        with open('data.csv', 'a') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow([date, *[d.strip() for d in data.values()]])
    return
            

def parseBody(body):
    """Parses the body of an email and returns a dictionary of the data"""
    body = body.split(b'\n')
    data = {}
    for line in body:
        if b':' in line:
            key, value = line.split(b':')
            data[key.decode('utf-8')] = value.decode('utf-8')

    return data

def parseEmail(email):
    """Parses an email and returns a dictionary of the data"""
    bodyBase64 = ''

    if (email['payload']['mimeType'] == 'multipart/mixed'):
        bodyBase64 = email['payload']['parts'][0]['body']['data']
    else:
        bodyBase64 = email['payload']['body']['data']
    body = base64.urlsafe_b64decode(bodyBase64)
    date = datetime.fromtimestamp(int(email['internalDate'])/1000.0)

    print(date)
    parsedBody = parseBody(body)
    addToCsv(date, parsedBody)
    return 
    


def main():
    """Reads emails from a gmail account and parses them"""
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        # Call the Gmail API
        service = build('gmail', 'v1', credentials=creds)
        # folders = service.users().labels().list(userId='me').execute()
        # print(folders)

        nextPageToken = ''
        while (True):
            results = service.users().messages().list(userId='me', labelIds=['Label_3'], maxResults=20, pageToken=nextPageToken).execute()
            for result in results['messages']:
                print('============================================================')
                message = service.users().messages().get(userId='me', id=result['id']).execute()
                print(message)
                parseEmail(message)
                print('\n')
            if('nextPageToken' in results or nextPageToken==''):
                nextPageToken = results['nextPageToken']
                print('nextPageToken={}'.format(nextPageToken))
            else:
                break

        

    except HttpError as error:
        # TODO(developer) - Handle errors from gmail API.
        print(f'An error occurred: {error}')


if __name__ == '__main__':
    main()