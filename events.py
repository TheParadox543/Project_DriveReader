from __future__ import print_function

import os
import os.path
import json
import re
import sys

# Install necessary libraries with pip install -r requirements.txt
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

# If modifying these scopes, delete the file token.json.
SCOPES = [
    "https://www.googleapis.com/auth/drive.metadata"
]

class DriveReader():
    """This project aims to read files in a drive and categorize them."""

    def __init__(self) -> None:
        """Initialize the class."""
        self.creds = None
        self.research_id = "1-kD9F_NrWLANsFNWGih9r2Il9BoAvCsu"
        self.validate_user()

    def validate_user(self):
        """Validate the program if the user who runs it is registered."""

        # * If using in colab, write data to a file to read it.
        # token_data = {
        #     "token": "ya29.a0AVvZVsq8eUGpg0u9HQyI-7QbepWZaII6cNfKL3M4o5pqKpmJ
        # yAkdhRAFZtyVlnQgrYkoflZ-QlpBfDEXl9uzNH5iT2DfAITeV90Pbjqtam5RpAw5tg0q
        # Haw2vBoFnaArEjIx3HcXp9hb-W3PEFyPkYGHvQfqT1oaCgYKAZgSAQASFQGbdwaIpxNZ
        # uRCQhdNrvsJeal5xTw0166", 
        #     "refresh_token": "1//0g8z1TL7yIR9nCgYIARAAGBASNwF-L9IrAH69OICd-
        # hAn12CsP-Q9CRFUZRKmm3QxyKKmrTwDKBeRjPMBet6OUMgwpUE4sE1Ood4", 
        #     "token_uri": "https://oauth2.googleapis.com/token", 
        #     "client_id": "387082150823-sclbdmg71jaqpsi1clv8hcqc3dvb7beg.apps.
        # googleusercontent.com", 
        #     "client_secret": "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", 
        #     "scopes": ["https://www.googleapis.com/auth/drive.metadata.readonly"], 
        #     "expiry": "2023-02-09T11:11:14.527394Z"
        # }

        # The file stores user's access and refresh tokens, and is created 
        # automatically when first authorization flow is completed.
        if os.path.exists("token.json"):
            self.creds = Credentials.from_authorized_user_file("token.json", 
                                                               SCOPES)

        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                # * If using in colab, write data to a file to read it.
                # credential_data ={
                #     "installed": {
                #         "client_id": "387082150823-sclbdmg71jaqpsi1clv8hcqc3
                # dvb7beg.apps.googleusercontent.com",
                #         "project_id":"drivereader-376706",
                #         "auth_uri":"https://accounts.google.com/o/oauth2/auth",
                #         "token_uri":"https://oauth2.googleapis.com/token",
                #         "auth_provider_x509_cert_url":"https://www.googleapis
                # .com/oauth2/v1/certs",
                #         "client_secret":"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",
                #         "redirect_uris":["http://localhost"]
                #     }
                # }
                flow = InstalledAppFlow.from_client_secrets_file(
                    "credentials.json", SCOPES)
                self.creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open("token.json", "w") as token:
                token.write(self.creds.to_json())

        # Create a connection with drive.
        self.service = build("drive", "v3", credentials=self.creds)

    def search_folder(self):
        """Search for a specific folder."""
        try:
            response = self.service.files().get(fileId=self.research_id).execute()
            print(response)
        except HttpError as error:
            print(f"An error occurred: {error}")

    def sort_files_in_folder(self):
        """Sort the files in the folder."""
        # Set data empty and then start execution.
        self.data = {}
        self.exempt = []

        try:
            page_token = None
            while True:
                # Search for all files with the folder as parent.
                response = self.service.files().list(
                    q=f"'{self.research_id}' in parents and trashed = false",
                    spaces='drive',
                    fields='nextPageToken, files(name)'
                ).execute()

                for file in response.get("files"):
                    # print(file)
                    if file.get("name") is not None:
                        self.classify_file(file.get("name"))
                page_token = response.get("nextPageToken", None)

                if page_token is None:
                    break

        except HttpError as error:
            print(f"An error occurred: {error}")

        else:
            with open("data.json", "w") as file:
                data_obj = json.dumps(self.data, indent=4)
                file.write(data_obj)
            with open("exempt.json", "w") as file:
                exempt_obj = json.dumps(self.exempt, indent=4)
                file.write(exempt_obj)

    def classify_file(self, name:str):
        """Classify the file in categories based on naming structure."""
        # try:
        #     with open("classification.json", "r") as file:
        #         self.classification = json.load(file)
        # except FileNotFoundError:
        #     self.classification = {}
        # except json.decoder.JSONDecodeError:
        #     self.classification = {}

        try:
            date, category, extra = name.split("_", 2)
        except ValueError:
            self.exempt.append(name)
        else:
            try:
                year, month = int(date[:4]), int(date[4:6])
                # print(year, month, day)
                if month > 0 and month < 5:
                    year = f"{year-1}-{year}"
                else:
                    year = f"{year}-{year+1}"
            except ValueError:
                self.exempt.append(name)
            else:
                year_data = self.data.get(year, {"0": {"0": 0}})
                category_val = year_data.get(category, 0)
                year_data.update({category: category_val + 1})
                if "0" in year_data:
                    year_data.pop("0")
                self.data.update({year: year_data})
                if "0" in self.data:
                    self.data.pop("0")

    def main(self):
        """The main function of the class."""
        self.sort_files_in_folder()
        # self.classify_file("20220730_cprs_rv_1.pdf")

class ExcelWorker():
    """The class that will handle interaction with the excel workbooks."""

    def __init__(self) -> None:
        """Initialize the class"""
        self.classification = list(dict())
        self.read_classification_exl()

    def read_classification_exl(self):
        """Read the classification categories."""
        # Load the workbook, if not exit the program.
        try:
            self.workbook:Workbook = load_workbook("doc_classification.xlsx")
        except FileNotFoundError: 
            print("Classification file not found.")
            sys.exit()
        worksheet: Worksheet
        # Loop through all sheets and categorize each code.
        for worksheet in self.workbook:
            ws_dict = {}
            for row in worksheet.iter_rows(min_col=2, max_col=3):
                code, name = row[0].value, row[1].value
                if code is not None and name is not None:
                    ws_dict[code] = name
            self.classification.append({worksheet.title: ws_dict})
        with open("classification.json", "w") as file:
            class_obj = json.dumps(self.classification, indent=4)
            file.write(class_obj)
        return self.classification

if __name__ == "__main__":
    exl = ExcelWorker()
    # Driver Code
    # DR = DriveReader()
    # if DR.creds and DR.creds.valid:
    #     try:
    #         DR.main()
    #     except KeyboardInterrupt:
    #         print("\n\nExiting program by interrupt.")
    # else:
    #     sys.exit()