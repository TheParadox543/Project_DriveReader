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

class ExcelWorker():
    """The class that will handle interaction with the excel workbooks."""

    def __init__(self) -> None:
        """Initialize the class"""
        self.classification: dict[dict[str, str]] = {}
        self.code_groups = {}
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
                    self.code_groups[code] = worksheet.title
            self.classification.update({worksheet.title: ws_dict})
        with open("classification.json", "w") as file:
            class_obj = json.dumps(self.classification, indent=4)
            file.write(class_obj)
        return self.classification

    def write_to_excel(self, drive_data: dict[str, dict[str, int]]):
        """Write data from the drive to the excel sheet."""

class DriveReader():
    """This project aims to read files in a drive and categorize them."""

    def __init__(self) -> None:
        """Initialize the class."""
        self.excelWorker = ExcelWorker()
        self.categories = self.excelWorker.classification
        self.code_keys = self.excelWorker.code_groups
        self.creds = None
        self.research_id = "1-kD9F_NrWLANsFNWGih9r2Il9BoAvCsu"
        self.validate_user()

    def validate_user(self):
        """Validate the program if the user who runs it is registered."""
        # The file stores user's access and refresh tokens, and is created 
        # automatically when first authorization flow is completed.
        if os.path.exists("token.json"):
            self.creds = Credentials.from_authorized_user_file("token.json", 
                                                               SCOPES)

        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    "credentials.json", SCOPES)
                self.creds = flow.run_local_server(port=0)
            # Save the credentials for the next run.
            with open("token.json", "w") as token:
                token.write(self.creds.to_json())

        # Create a connection with drive.
        self.service = build("drive", "v3", credentials=self.creds)

        # Set data empty and then start execution.
        self.data = {}
        self.exempt = []

    def sort_files_in_folder(self, category_name: str):
        """Sort the files in the folder."""
        folders = self.search_folder(category_name)
        if folders is None or folders == []:
            return
        for folder in folders:
            folder_id, folder_name = folder.get("id"), folder.get("name")
            try:
                page_token = None
                while True:
                    # Search for all files with the folder as parent.
                    response = self.service.files().list(
                        q=f"'{folder_id}' in parents and trashed = false",
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

    def search_folder(self, category_name: str):
        """Search for a specific folder."""
        try:
            response = self.service.files().list(
                q=f"name contains '{category_name}' and mimeType = \
                    'application/vnd.google-apps.folder'"
            ).execute()
            return response.get("files", None)
        except HttpError as error:
            print(f"An error occurred: {error}")
            return None

    def classify_file(self, name:str):
        """Classify the file in categories based on naming structure."""
        try:
            date, code, extra = name.split("_", 2)
            code = code.upper()
            if code not in self.code_keys:
                raise KeyError
        except ValueError:
            self.exempt.append(name)
        except KeyError:
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
                category = self.code_keys[code]
                category_data = self.data.get(category, {"0": {"0": 0}})
                year_data = category_data.get(year, {"0": 0})
                code_val = year_data.get(code, 0)
                year_data.update({code: code_val + 1})
                if "0" in year_data:
                    year_data.pop("0")
                category_data.update({year: year_data})
                if "0" in category_data: 
                    category_data.pop("0")
                self.data.update({category: category_data})
                if "0" in self.data:
                    self.data.pop("0")

    def main(self):
        """The main function of the class."""
        for category in self.categories:
            self.sort_files_in_folder(category)
        # self.classify_file("20220730_cprs_rv_1.pdf")


if __name__ == "__main__":
    # Driver Code
    try:
        DR = DriveReader()
        if DR.creds and DR.creds.valid:
            DR.main()
        else:
            sys.exit()
    except KeyboardInterrupt:
        print("\n\nExiting program by interrupt.")