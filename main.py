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
]

class DriveReader():
    """This project aims to read files in a drive and categorize them."""

    def __init__(self) -> None:
        """Initialize the class."""
        self.creds = None

    def validate_user(self):
        """Validate the program if the user who runs it is registered."""
        # Take user's name.
        user = input("Enter your credentials:")

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

    def search_for_all_folders(self):
        """A function to identify all folders in a drive."""

    def main(self):
        """The main function of the class."""
        self.validate_user()


if __name__ == '__main__':
    DR = DriveReader()
    DR.main()