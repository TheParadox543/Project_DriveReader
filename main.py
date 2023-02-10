from __future__ import print_function

import os
import os.path
import json
import re
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

#TODO multiple versions of the same file

class DriveReader():
    """This project aims to read files in a drive and categorize them."""

    def __init__(self) -> None:
        """Initialize the class."""
        self.creds = None
        self.validate_user()
        try:
            with open("database.json", "r") as file:
                self.database = json.load(file)
        except json.decoder.JSONDecodeError:
            self.database = {}
        except FileNotFoundError:
            self.database = {}
        try:
            with open("base_data.json", "r") as file:
                self.data = json.load(file)
        except json.decoder.JSONDecodeError:
            self.data = {}
        except FileNotFoundError:
            self.data = {}
        self.exempt = []

    def validate_user(self):
        """Validate the program if the user who runs it is registered."""
        # Take user's name.
        # user = input("Enter your credentials:")
        user = "sam_christ"

        # # * Check if the user is in the database of authorized users.
        # with open("authorized.json") as file:
        #     json_obj = json.load(file)
        # if user not in json_obj["authors"]:
        #     print("Invalid user")
        #     return

        # * If using in colab, write data to a file to read it.
        # token_data = {
        #     "token": "ya29.a0AVvZVsq8eUGpg0u9HQyI-7QbepWZaII6cNfKL3M4o5pqKpmJyAkdhRAFZtyVlnQgrYkoflZ-QlpBfDEXl9uzNH5iT2DfAITeV90Pbjqtam5RpAw5tg0qHaw2vBoFnaArEjIx3HcXp9hb-W3PEFyPkYGHvQfqT1oaCgYKAZgSAQASFQGbdwaIpxNZuRCQhdNrvsJeal5xTw0166", 
        #     "refresh_token": "1//0g8z1TL7yIR9nCgYIARAAGBASNwF-L9IrAH69OICd-hAn12CsP-Q9CRFUZRKmm3QxyKKmrTwDKBeRjPMBet6OUMgwpUE4sE1Ood4", 
        #     "token_uri": "https://oauth2.googleapis.com/token", 
        #     "client_id": "387082150823-sclbdmg71jaqpsi1clv8hcqc3dvb7beg.apps.googleusercontent.com", 
        #     "client_secret": "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", 
        #     "scopes": ["https://www.googleapis.com/auth/drive.metadata.readonly"], 
        #     "expiry": "2023-02-09T11:11:14.527394Z"
        # }

        # The file stores user's access and refresh tokens, and is created 
        # automatically when first authorization flow is completed.
        if os.path.exists(f"{user}.json"):
            self.creds = Credentials.from_authorized_user_file(f"{user}.json", 
                                                               SCOPES)

        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                # * If using in colab, write data to a file to read it.
                # credential_data ={
                #     "installed": {
                #         "client_id": "387082150823-sclbdmg71jaqpsi1clv8hcqc3dvb7beg.apps.googleusercontent.com",
                #         "project_id":"drivereader-376706",
                #         "auth_uri":"https://accounts.google.com/o/oauth2/auth",
                #         "token_uri":"https://oauth2.googleapis.com/token",
                #         "auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs",
                #         "client_secret":"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",
                #         "redirect_uris":["http://localhost"]
                #     }
                # }
                flow = InstalledAppFlow.from_client_secrets_file(
                    "credentials.json", SCOPES)
                self.creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(f"{user}.json", "w") as token:
                token.write(self.creds.to_json())

        # Create a connection with drive.
        self.service = build("drive", "v3", credentials=self.creds)

    def search_id_file(self, name:str):
        """Search for a specific file or folder by its given name.
        
        - name:str: The name of the file to search for."""
        try:
            response = self.service.files().list(
                q=f"name contains '{name}'",
                spaces="drive",
                fields="files(name, id, parents)"
            ).execute()
            for file in response.get("files", []):
                print("File found", file)

        except HttpError:
            print("Error occured.")

    def search_for_all_folders(self):
        """A function to identify all folders in a drive."""
        try:
            files = []
            page_token = None
            while True:
                # Search for all folders in drive.
                response:dict = self.service.files().list(
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
        # while True:
        try:
            # * Search for folder "CUCS - Faculty". If files are in a 
            # * different folder, change the query.
            main_folder_name = "CUCS - Faculty"
            response:dict = self.service.files().list(
                q=f"name contains '{main_folder_name}' and \
                    mimeType='application/vnd.google-apps.folder'",
                spaces='drive',
                fields='nextPageToken, files(id, name, parents)'
            ).execute()
            try:
                main_folder_id = response.get("files")[0].get("id")
                # main_folder_id = "1YB805MsCxVaCmWuwB3jIewp5RNOKX65n"
            except IndexError:
                print("Could not find the folder.")
                return

            else:
                if main_folder_id is not None:
                    self.database_new = {}
                    self.data_new = {}
                    # Run the program through the recursive function.
                    self.data_new = self.search_folders(main_folder_id, 
                                                main_folder_name)
                    self.update_json_files()
                    time.sleep(3)
                else:
                    print("Could not find the folder.")

        except HttpError as error:
            print(F'An error occurred: {error}')

    def search_folders(self, parent_id:str, parent_name:str):
        """
        A recursive function to keep listing all files and folders located
        inside of a folder.

        Parameters
        ---------
        - parent_id`str`: The id of the parent folder.
        - parent_name`str`: The name of the parent folder. 

        Returns
        -------
        A dictionary with the file in the parent folder as key.
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
                    # // print(f"{start_str}{file.get('name')}, {file.get('mimeType')}")
                    # Check if the file is of folder type
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
        """Categorize the data from files to a database.

        Parameters
        ----------
        - teacher_name`str`: The name of the folder to associate in the database.
        - file_name`str`: The name of the file extracted from the file name if
                        it is given in the proper format.
        """

        # Take data from file name and verify if it is in proper format.
        break_point = file_name[4]
        name_data = file_name.split(break_point, 2)
        if len(name_data) < 3:
            self.exempt.append((file_name, teacher_name))
            return

        year, type, name = name_data
        type = type.upper()
        if re.search("[a-zA-Z]", year) or re.search("\d", type):
            self.exempt.append((file_name, teacher_name))
            return

        # Extract teacher data, supply if not already there.
        teacher_data = self.database_new.get(teacher_name, {year:{type:0}})

        # Extract year data of a teacher, supply if not already there.
        year_data = teacher_data.get(year, {type:0})

        # Take a count of how many papers of a type are there in a year.
        count = year_data.get(type, 0)

        # Update all the data as required.
        year_data.update({type:count+1})
        teacher_data.update({year:year_data})
        self.database_new.update({teacher_name:teacher_data})

    def update_json_files(self):
        """Update the json files if there is a change in data."""
        # if self.database_new != self.database:
        #     # print("Database updated.")
        self.database = self.database_new
        with open("database.json", "w") as file:
            json_obj = json.dumps(self.database_new, indent=4)
            file.write(json_obj)

        try:
            with open("dept_name_list.json", "r") as file:
                text = json.load(file)
        except FileNotFoundError:
            text = {"None": "None"}
        self.dept_json = {}
        for name in self.database_new:
            if name in text:
                dept_name = text.get(name, "Unkown")
                dept_data = self.dept_json.get(dept_name, {"0":{"0":0}})
                for year in self.database_new[name]:
                    year_data = dept_data.get(year, {"0":"0"})
                    for type in self.database_new[name][year]:
                        count = year_data.get(type, 0)
                        count += self.database[name][year][type]
                        year_data.update({type: count})
                        if "0" in year_data:
                            year_data.pop('0')
                    dept_data.update({year:year_data})
                    if "0" in dept_data:
                        dept_data.pop("0")
                self.dept_json.update({dept_name:dept_data})
                if "0" in self.dept_json:
                    self.dept.pop("0")
        with open("dept_data.json", "w") as file_write:
            json_obj = json.dumps(self.dept_json, indent=4)
            file_write.write(json_obj)

        # if self.data_new != self.data:
        #     # print("Data updated.")
        self.data = self.data_new
        with open("base_data.json", "w") as f:
            json_obj = json.dumps(self.data_new, indent=4)
            f.write(json_obj)

        with open("exempt.json", "w") as file:
            json_obj = json.dumps(self.exempt, indent=4)
            file.write(json_obj)

        os.system("start NOTEPAD.EXE database.json")

    def main(self):
        """The main function of the class."""
        # while True:
            # command = input("Enter command:")
            # if command == "Search all folders":
            #     self.search_for_all_folders()
            # elif command.startswith("find"):
            #     print(command.split()[1])
            #     self.search_id_file(command.split()[1])
            # elif command == "search" or command == "run":
        self.categorize_folders_from_drive()
            # elif command == "data":
            #     self.categorize_data("")
            # elif command == "quit" or command == "exit":
            #     sys.exit()
            # else:
            #     print("Invalid Command")


if __name__ == "__main__":
    try:
        DR = DriveReader()
    except KeyboardInterrupt:
        print("\n\nExiting program by interrupt.")
    if DR.creds and DR.creds.valid:
        try:
            DR.main()
        except KeyboardInterrupt:
            print("\n\nExiting program by interrupt.")
    else:
        sys.exit()