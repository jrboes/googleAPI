import os
import time
import json
import googleapiclient.errors
import google.oauth2.credentials
import google.auth.transport.requests
from googleapiclient.discovery import build
import googleapi.spreadsheet


class Client():
    """Create a connection to a Google drive API."""

    def __init__(self, cred_file='../.cred/token.json'):
        self.current_uid = None
        self.sheets = {}
        self.files = {'sheets': [], 'folders': []}

        # TODO: Need a process to request a token if not already stored.
        gapi = os.path.join(
            os.path.dirname(googleapi.__file__), cred_file)
        with open(gapi, 'r') as f:
            credentials = google.oauth2.credentials.Credentials(**json.load(f))

        # Refresh the token as necessary
        if credentials.expired and credentials.refresh_token:
            credentials.refresh(google.auth.transport.requests.Request())

            with open(gapi, 'w') as f:
                json.dump(credentials, f)

        # Connect to drive and sheets API
        api = {}
        builds = [['sheets', 'v4'], ['drive', 'v3']]
        for entry in builds:
            api[entry[0]] = build(
                *entry,
                credentials=credentials,
                cache_discovery=False)

        self.api = api

    def get_files(self):
        """Return the files and folders of a Google Drive."""
        page_token = None
        while True:
            request = self.api['drive'].files().list(
                q='trashed = false',
                fields='nextPageToken, files(id, name, parents, mimeType)',
                pageToken=page_token)
            response = self._execute_requests(request)

            for f in response.get('files'):
                ID, name, parent = f.get('id'), f.get('name'), f.get('parents')
                if f.get('mimeType').endswith('folder'):
                    self.files['folders'] += [{
                        'title': name,
                        'parent': parent,
                        'id': ID}]
                elif f.get('mimeType').endswith('spreadsheet'):
                    self.files['sheets'] += [{
                        'title': name,
                        'parent': parent,
                        'id': ID}]

            page_token = response.get('nextPageToken', None)
            if page_token is None:
                return self.files

    def get_spreadsheet(self, title):
        """Return a Google Spreadsheet from Google Drive via title."""
        sheets = self.files.get('sheets')
        if not sheets:
            sheets = self.get_files().get('sheets')

        if isinstance(title, str):
            id = [_.get('id') for _ in sheets if _.get('title') == title][0]

        request = self.api['sheets'].spreadsheets().get(spreadsheetId=id)
        response = self._execute_requests(request)
        sheet = googleapi.spreadsheet.SpreadSheet(
            client=self, response=response)

        return sheet

    def move(self, file_id, folder_id):
        """Move the location of a spreadsheet from one file to another."""

        # Catch this from self.sheets if it already exists
        file = self.api['drive'].files().get(
            fileId=file_id, fields='parents').execute()
        previous_parents = ",".join(file.get('parents'))

        # Move the file to the new folder
        file = self.api['drive'].files().update(
            fileId=file_id, addParents=folder_id,
            removeParents=previous_parents,
            fields='id, parents').execute()

    def _execute_requests(self, request):
        """Execute a request to the Google Sheets API v4."""
        try:
            response = request.execute(num_retries=3)
        except googleapiclient.errors.HttpError as error:
            if error.resp['status'] == '429' and self.check:
                time.sleep(10)
                response = request.execute(num_retries=3)
            else:
                raise

        return response
