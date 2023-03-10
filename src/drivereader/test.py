from __future__ import print_function

import os.path
import json

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build, Resource
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = [
    'https://www.googleapis.com/auth/drive.metadata.readonly',
    # 'https://www.googleapis.com/auth/drive'
]


def main():
    """Shows basic usage of the Drive v3 API.
    Prints the names and ids of the first 10 files the user has access to.
    """
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
        service:Resource = build('drive', 'v3', credentials=creds)
        print(type(service.files().list()))

        # Call the Drive v3 API
        results = service.files().list(
            pageSize=10, fields="nextPageToken, files(id, name)").execute()
        items = results.get('files', [])

        if not items:
            print('No files found.')
            return
        print('Files:')
        for item in items:
            print(u'{0} ({1})'.format(item['name'], item['id']))
    except HttpError as error:
        # TODO(developer) - Handle errors from drive API.
        print(f'An error occurred: {error}')

def search_folder():
    """Search file in drive location

    Load pre-authorized user credentials from the environment.
    TODO(developer) - See https://developers.google.com/identity
    for guides on implementing OAuth2 for the application.
    """
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    try:
        # create drive api client
        service = build('drive', 'v3', credentials=creds)
        files = []
        page_token = None
        while True:
            # pylint: disable=maybe-no-member
            response = service.files().list(q="mimeType = 'application/vnd.google-apps.folder'",
                                            spaces='drive',
                                            fields='nextPageToken, '
                                                   'files(parents, name)',
                                            pageToken=page_token).execute()
            for file in response.get('files', []):
                # Process change
                print(F'Found file: {file.get("name")}, {file.get("parents")}')
            files.extend(response.get('files', []))
            page_token = response.get('nextPageToken', None)
            if page_token is None:
                break

    except HttpError as error:
        print(F'An error occurred: {error}')
        files = None

    return files

def search_teacher():
    """Search file in drive location

    Load pre-authorized user credentials from the environment.
    TODO(developer) - See https://developers.google.com/identity
    for guides on implementing OAuth2 for the application.
    """
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    try:
        # create drive api client
        service = build('drive', 'v3', credentials=creds)
        files = []
        page_token = None
        while True:
            # pylint: disable=maybe-no-member
            response = service.files().list(q="mimeType = 'application/vnd.google-apps.folder' and '1vV5M8ZPeeFBKReEa1n7eFhMJev8Z7luu' in parents",
                                            spaces='drive',
                                            fields='nextPageToken, '
                                                   'files(parents, name)',
                                            pageToken=page_token).execute()
            for file in response.get('files', []):
                # Process change
                print(F'Found file: {file.get("name")}, {file.get("parents")}')
            files.extend(response.get('files', []))
            page_token = response.get('nextPageToken', None)
            if page_token is None:
                break

    except HttpError as error:
        print(F'An error occurred: {error}')
        files = None

    return files

def search_test_1():
    """Search file in drive location

    Load pre-authorized user credentials from the environment.
    TODO(developer) - See https://developers.google.com/identity
    for guides on implementing OAuth2 for the application.
    """
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    try:
        # create drive api client
        service = build('drive', 'v3', credentials=creds)
        files = []
        page_token = None
        while True:
            # pylint: disable=maybe-no-member
            # response = service.files().list(q="name contains 'Teacher'",
            #                                 spaces='drive',
            #                                 fields='nextPageToken,'
            #                                         'files(id, name, parents)',
            #                                 pageToken=page_token).execute()
            response = service.files().list(q="mimeType = 'application/vnd.google-apps.folder' and '1vV5M8ZPeeFBKReEa1n7eFhMJev8Z7luu' in parents",
                                            spaces='drive',
                                            fields='nextPageToken, '
                                                   'files(parents, name, id)',
                                            pageToken=page_token).execute()
            print(response)
            for file in response.get('files', []):
                # Process change
                print(F'Found file: {file.get("name")}, {file.get("id")}, {file.get("parents")}')
            files.extend(response.get('files', []))
            page_token = response.get('nextPageToken', None)
            if page_token is None:
                break

    except HttpError as error:
        print(F'An error occurred: {error}')
        files = None

    return files

def search_test():
    """Search file in drive location

    Load pre-authorized user credentials from the environment.
    TODO(developer) - See https://developers.google.com/identity
    for guides on implementing OAuth2 for the application.
    """
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    data_to_dump = {}

    try:
        # create drive api client
        service = build('drive', 'v3', credentials=creds)
        files = []
        page_token = None

        response:dict = service.files().list(q="name contains 'Project DriveReader'",
                                      spaces='drive',
                                      fields='nextPageToken, files(id)',
                                      pageToken=page_token).execute()
        # for file in response.get("files"):
        #     print(file)
        main_folder_id = response.get("files")[0].get("id")
        # print(main_folder_id)

        def search_folders(parent_id:str, parent_name:str, start_str:str):
            new_response = service.files().list(q=f"'{parent_id}' in parents",
                                                spaces='drive',
                                                fields='nextPageToken, files(id, name, mimeType)',
                                                pageToken=page_token).execute()

            parent_key = []

            # print(f"Parent name: {parent_name}")
            for file in new_response.get("files"):
                # print(f"{start_str}{file.get('name')}, {file.get('mimeType')}")
                if file.get('mimeType')[28:] == "folder":
                    returned_folder = search_folders(file.get("id"), file.get("name"), start_str+"\t")
                    # print(returned_folder)
                    parent_key.append(returned_folder)
                else:
                    parent_key.append((file.get("name")))#, file.get('mimeType')[28:]))
            return {parent_name:parent_key}

        data_to_dump = search_folders(main_folder_id, "Project DriveReader","")
        print(data_to_dump)
        with open("data.json", "w") as f:
            json_obj = json.dumps(data_to_dump, indent=4)
            f.write(json_obj)

    except HttpError as error:
        print(F'An error occurred: {error}')
        files = None

    return files


if __name__ == '__main__':
    search_test()