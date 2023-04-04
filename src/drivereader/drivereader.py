"""
This project intends to read the required files from Google Drive, and
then based on the naming convention applied to those files, analyse and
categorize the available data into the necessary format.

Project Creation: Feb 3, 2023
Lead Developer: Sam Alex Koshy
Project Contributors: Ashok Immanuel, Rohini V.
"""


# Install necessary libraries with pip install -r requirements.txt
from __future__ import print_function

# Import in-built modules.
import logging
from io import BytesIO
from json import dumps, load
from os import path, remove
from os import system as ossystem
from pprint import PrettyPrinter
from sys import exit as sysexit
from typing import TypeVar

# Import project specific modules.
from google.auth.exceptions import RefreshError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

from .excel import ExcelWorker

# If modifying these scopes, delete the file token.json.
SCOPES = [
    "https://www.googleapis.com/auth/drive"
]

pp = PrettyPrinter(indent=4)

# Using the logs.
logger_monitor = logging.getLogger(__name__)
logger_monitor.setLevel(logging.ERROR)
handler = logging.FileHandler("drive_reader_logs.log")
handler.setFormatter(logging.Formatter("%(asctime)s:%(levelname)s:%(name)s: %(message)s"))
logger_monitor.addHandler(handler)


# * Declare a few types to help with understanding.
Category = TypeVar("Category", bound=str)
Classification = TypeVar("Classification", bound=str)
Code = TypeVar("Code", bound=str)
Name = TypeVar("Name", bound=str)
Year = TypeVar("Year", bound=str)

def sort_dictionary(unsorted_dict: dict[str, ], reverse=False):
    """A util function to sort the keys of a dictionary.

    Parameters
    ---------
    - unsorted_dict`dict[str, Any]`: The dictionary that needs to be sorted.
    - reverse`bool`: Whether the keys need to be sorted in reverse order.

    Returns
    ------
    - sorted_dict`dict[str, Any]`: The dictionary with keys sorted.
    """
    order_list: list[str] = sorted(unsorted_dict.keys(), reverse=reverse)
    sorted_dictionary = dict({i: unsorted_dict[i] for i in order_list})
    return sorted_dictionary


class DriveReader():
    """This project aims to read files in a drive and categorize them."""

    def __init__(self) -> None:
        """Initialize the class."""
        self.creds = None
        self.initialize_connection()

    def initialize_connection(self):
        """Make the initial connection with drive."""
        # The file stores user's access and refresh tokens, and is created
        # automatically when first authorization flow is completed.
        if path.exists("token.json"):
            self.creds = Credentials.from_authorized_user_file(
                "token.json",
                SCOPES
            )

        if not self.creds or not self.creds.valid:
            refresh = False
            if self.creds and self.creds.expired and self.creds.refresh_token:
                try:
                    self.creds.refresh(Request())
                except RefreshError:
                    remove("token.json")
                else:
                    refresh = True
            if refresh is False:
                flow = InstalledAppFlow.from_client_secrets_file(
                    "credentials.json", SCOPES)
                self.creds = flow.run_local_server(port=0)
            # Save the credentials for the next run.
            with open("token.json", "w") as token:
                token.write(self.creds.to_json())

        # Create a connection with drive.
        self.service = build("drive", "v3", credentials=self.creds)

        if self.creds and self.creds.valid:
            return "Connection made."
        else:
            return "Connection failed."

    def search_file(self, file_name: str):
        """Search for a specific file."""
        try:
            response = self.service.files().list(
                q=f"name contains '{file_name}'"
            ).execute()
            return response.get("files", None)

        except HttpError as error:
            logger_monitor(f"An error occurred: {error}")
            return None

    def search_folder(self, category_name: str):
        """Search for a specific folder."""
        try:
            response = self.service.files().list(
                q=f"name contains '{category_name}' and mimeType = \
                    'application/vnd.google-apps.folder'"
            ).execute()
            folders = response.get("files", None)
            if len(folders) > 0:
                return folders[0]
            else: return None

        except HttpError as error:
            logger_monitor(f"An error occurred: {error}")
            return None

    def download_sheet(self):
        """Download the required excel sheet."""
        self.excel_sheet_id = "1b5yJfOIWCHXdr7VbFxoNLs_SI5zPR7CL0MsCI1zqWaM"
        try:
            mime_type = "application/vnd.openxmlformats-officedocument"
            mime_type += ".spreadsheetml.sheet"
            request = self.service.files().export_media(fileId=self.excel_sheet_id,
                                                        mimeType=mime_type)
            file = BytesIO()
            downloader = MediaIoBaseDownload(file, request)
            done = False
            while done is False:
                status, done = downloader.next_chunk()

            while True:
                try:
                    with open("data/doc_classification.xlsx", "wb") as write_file:
                        write_file.write(file.getbuffer())
                except PermissionError:
                    ossystem("taskkill /im EXCEL.EXE naac.xlsx")
                else:
                    break

            logger_monitor(F'Download {int(status.progress() * 100)}%.')
            return True

        except HttpError as error:
            logger_monitor(f"{error} has occurred.")
            return False

    def categorize_files(self):
        """Categorize the files in the various folders according to code."""
        try:
            with open("data/folders.json", "r") as file:
                folder_names: list[str] = load(file)
        except FileNotFoundError:
            logger_monitor("Please specify the folders to search in `folders.json`.")
            self.data = None
            return

        self.data: dict[Category, dict[Year, dict[Code, int]]] = {
            i: {
                "0": {
                    "0": 0
                }
            } for i in self.categories}
        self.exempt: list[tuple[Name, str]] = []

        for folder_search in folder_names:
            folder = self.search_folder(folder_search)
            folder_id, folder_name = folder.get("id"), folder.get("name")
            try:
                page_token = None
                while True:
                    # Search for all files with the folder as parent.
                    response = self.service.files().list(
                        q=f"'{folder_id}' in parents and trashed = false",
                        spaces='drive',
                        fields='nextPageToken, files(name)',
                        pageToken=page_token
                    ).execute()

                    for file in response.get("files"):
                        # logger_monitor(file)
                        if file.get("name") is not None:
                            if_failed = self.classify_file(file.get("name"))
                            if if_failed:
                                self.exempt.append((if_failed, folder_name))
                    page_token = response.get("nextPageToken", None)

                    if page_token is None:
                        break

            except HttpError as error:
                logger_monitor(f"An error occurred: {error}")

        else:
            for category in list(self.data.keys()):
                while "0" in self.data[category]:
                    del self.data[category]["0"]
                for year, year_data in self.data[category].items():
                    while "0" in year_data:
                        del year_data["0"]
                    self.data[category][year] = sort_dictionary(year_data)
                self.data[category] = sort_dictionary(self.data[category], True)
                if "0" in self.data[category]:
                    self.data.pop(category)

    def classify_file(self, name:str):
        """Classify the file in categories based on naming structure."""
        try:
            date, code, extra = name.split("_", 2)
            code = code.upper()
            if code not in self.code_list:
                raise KeyError
        except ValueError:
            return name
        except KeyError:
            return name
        else:
            try:
                year, month = int(date[:4]), int(date[4:6])
                # print(year, month, day)
                if month > 0 and month < 5:
                    year = f"{year-1}-{year}"
                else:
                    year = f"{year}-{year+1}"
            except ValueError:
                return name
            else:
                category = self.code_list[code][1]
                category_data = self.data.get(category, {"0": {"0": 0}})
                year_data = category_data.get(year, {"0": 0})
                code_val = year_data.get(code, 0)
                year_data.update({code: code_val + 1})
                category_data.update({year: year_data})
                self.data.update({category: category_data})

    def main(self):
        """The main function of DriveReader class."""
        self.download_sheet()
        self.excelWorker = ExcelWorker()
        self.code_list = self.excelWorker.code_list
        self.categories = self.excelWorker.classification_list.values()
        self.categorize_files()
        if self.data is not None:
            with open("data/data.json", "w") as file:
                data_obj = dumps(self.data, indent=4)
                file.write(data_obj)
            with open("data/exempt.json", "w") as file:
                exempt_obj = dumps(self.exempt, indent=4)
                file.write(exempt_obj)
            self.excelWorker.write_data_to_excel(self.data, self.exempt)
            self.excelWorker.write_naac_data_to_excel(self.data)


if __name__ == "__main__":
    # Driver Code
    try:
        DR = DriveReader()
        if DR.creds and DR.creds.valid:
            DR.main()
        else:
            print("Could not run the program due to invalid credentials.")
            print("Fix credentials and try again.")
            sysexit()
    except KeyboardInterrupt:
        print("\n\nExiting program by interrupt.")