from __future__ import print_function

import os.path
import json
import sys
import time

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build, Resource
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = [
    'https://www.googleapis.com/auth/drive.metadata.readonly',
]

class DriveReader():
    """This project aims to read files in a drive and categorize them."""

    def __init__(self) -> None:
        """Initialize the class."""
        self.creds = None
        self.validate_user()

    def validate_user(self):
        """Validate the program if the user who runs it is registered."""
        # Take user's name.
        user = input("Enter your credentials:")

        # Check if the user is in the database of authorized users.
        with open("authorized.json") as file:
            json_obj = json.load(file)
        if user not in json_obj["authors"]:
            print("Invalid user")
            return

        # The file stores user's access and refresh tokens, and is created 
        # automatically when first authorization flow is completed.
        if os.path.exists(f"{user}.json"):
            self.creds = Credentials.from_authorized_user_file(f"{user}.json", 
                                                               SCOPES)

        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    "credentials.json", SCOPES)
                self.creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(f"{user}.json", "w") as token:
                token.write(self.creds.to_json())

        # Create a connection with drive.
        self.service = build("drive", "v3", credentials=self.creds)

    def search_for_all_folders(self):
        """A function to identify all folders in a drive."""
        try:
            files = []
            page_token = None
            while True:
                # Search for all folders in drive.
                response:dict = self.service.files().list(
                    corpora="user",
                    q="mimeType='application/vnd.google-apps.folder'",
                    spaces='drive',
                    fields="nextPageToken, files(name, parents)",
                    supportsAllDrives=False,
                    pageToken=page_token
                ).execute()
                file: dict
                # Loop through all files in the response
                for file in response.get("files", []):
                    # Process change
                    print(F"Found file: {file.get('name')}, {file.get('parents')}")
                files.extend(response.get("files", []))
                page_token = response.get("nextPageToken", None)
                if page_token is None:
                    break

        except HttpError as error:
            print(f"An error occurred: {error}")
            files = None

        return files

    def categorize_folders_from_drive(self):
        """Search for folders and files in the Project folder."""
        while True:
            try:
                files = []

                response:dict = self.service.files().list(
                    q="name contains 'Project DriveReader'",
                    spaces='drive',
                    fields='nextPageToken, files(id)'
                ).execute()
                try:
                    main_folder_id = response.get("files")[0].get("id")
                except IndexError:
                    print("Could not find the folder.")
                    return

                else:
                    if main_folder_id is not None:
                        with open("database.json", "w") as file:
                            json_obj = json.dumps({})
                            file.write(json_obj)
                        files = self.search_folders(main_folder_id, 
                                                    "Project DriveReader")
                        with open("data.json", "w") as f:
                            json_obj = json.dumps(files, indent=4)
                            f.write(json_obj)
                        time.sleep(10)
                    else:
                        print("Could not find the folder.")

            except HttpError as error:
                print(F'An error occurred: {error}')
                files = None

    def search_folders(self, parent_id:str, parent_name:str):
        """
        A recursive function to keep listing all files and folders located
        inside of a folder.
        """
        try:
            page_token = None
            files = []
            while True:
                response = self.service.files().list(
                    q=f"'{parent_id}' in parents and trashed = false",
                    spaces='drive',
                    fields='nextPageToken, files(id, name, mimeType)',
                    pageToken=page_token
                ).execute()

                for file in response.get("files"):
                    # print(f"{start_str}{file.get('name')}, {file.get('mimeType')}")
                    if file.get('mimeType')[28:] == "folder":
                        returned_folder = self.search_folders(file.get("id"), 
                                                              file.get("name"))
                        files.append(returned_folder)
                    else:
                        self.categorize_data(parent_name, file.get("name"))
                        files.append((file.get("name")))
                page_token = response.get("nextPageToken", None)
                if page_token is None:
                    break

        except HttpError as error:
            print(f"An error occurred: {error}")

        return {parent_name:files}

    def categorize_data(self, teacher_name:str, file_name:str):
        """Categorize the data from files to a database."""
        name_data = file_name.split("_", 2)
        if len(name_data) < 3:
            return
        year, type, name = name_data
        with open("database.json", "r") as file:
            try:
                data = json.load(file)
            except json.decoder.JSONDecodeError:
                data = {}
        teacher_data = data.get(teacher_name)
        if teacher_data is None:
            data.update({teacher_name:{year:{type:1}}})
        else:
            if teacher_data.get(year) is not None:
                year_data = teacher_data[year]
                if year_data.get(type) is not None:
                    count = year_data[type]
                    year_data.update({type:count+1})
                else:
                    year_data.update({type:1})
                teacher_data.update({year:year_data})
            else:
                teacher_data.update({year:{type:1}})
            data.update({teacher_name:teacher_data})
        print(data)
        with open("database.json", "w+") as file:
            json_send_obj = json.dumps(data, indent=4)
            file.write(json_send_obj)

    def main(self):
        """The main function of the class."""
        try:
            while True:
                command = input("Enter command:")
                if command == "Search all folders":
                    self.search_for_all_folders()
                elif command == "quit" or command == "exit":
                    sys.exit()
                elif command == "search" or command == "run":
                    self.categorize_folders_from_drive()
                # elif command == "run":
                #     time.sleep(100)
                elif command == "data":
                    self.categorize_data("")
                else:
                    print("Invalid Command")
        except KeyboardInterrupt:
            print("\n\nExiting the program by interrupt.")


if __name__ == '__main__':
    try:
        DR = DriveReader()
    except KeyboardInterrupt:
        print("\n\nExiting program by interrput.")
    if DR.creds and DR.creds.valid:
        try:
            DR.main()
        except KeyboardInterrupt:
            print("\n\nExiting program by interrput.")
    else:
        sys.exit()