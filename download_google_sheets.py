import json
from os import path
import os
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import requests
from pathlib import Path
from logger import logger
from dotenv import load_dotenv
load_dotenv()


class SheetDownloader:
    SAVE_SHEETS_DIR = "downloaded-spreadsheets"

    def __init__(self, sheets_file) -> None:
        Path(self.SAVE_SHEETS_DIR).mkdir(parents=True, exist_ok=True)
        try:
            self.sheets_file = open(sheets_file, 'r')
            self.sheets_json = json.loads(self.sheets_file.read())
        except Exception as e:
            print(e)
            logger.critical(e)
            logger.debug("Exiting")
            exit()
        try:
            self.gauth = GoogleAuth()
            # Try to load saved client credentials
            client_secret = os.environ.get(
                "GOOGLE_CREDENTIAL_FILE_LOCATION", 'credentials.txt')
            client_secret_file = Path(client_secret)
            if not client_secret_file.is_file():
                msg = f"Invalid Google Auth File `{client_secret}` provided"
                logger.critical(msg)
                print(msg)
                client_secret = 'credentials.txt'
                f = open(client_secret, 'w')
                # f.save()
                msg = f"Using Default Google Auth File `{client_secret}` "
                logger.warning(msg)
                print(msg)

            self.gauth.LoadCredentialsFile(client_secret)
            if self.gauth.credentials is None:
                # Authenticate if they're not there
                msg = 'new authentication'
                print(msg)
                logger.debug(msg)

                self.gauth.GetFlow()
                self.gauth.flow.params.update({'access_type': 'offline'})
                self.gauth.flow.params.update({'approval_prompt': 'force'})
                self.gauth.LocalWebserverAuth()

            elif self.gauth.access_token_expired:
                # Refresh them if expired
                msg = 'refresh authentication'
                print(msg)
                logger.debug(msg)
                self.gauth.Refresh()
            else:
                # Initialize the saved creds
                msg = 'initializing the saved creds'
                print(msg)
                logger.debug(msg)
                self.gauth.Authorize()
            # Save the current credentials to a file
            self.gauth.SaveCredentialsFile(client_secret)
        except Exception as e:
            print(e)
            logger.exception(e)

    def download_spreadsheet(self, spreadsheet_id, filename):
        filename = filename+".xlsx"
        filename = path.join(self.SAVE_SHEETS_DIR, filename)
        url = 'https://docs.google.com/spreadsheets/d/' + \
            spreadsheet_id + '/export?format=xlsx'
        headers = {'Authorization': 'Bearer ' +
                   self.gauth.credentials.access_token}
        try:
            res = requests.get(url, headers=headers)
            with open(filename, 'wb') as f:
                f.write(res.content)
            msg = "File downloaded"
            return (True, msg, filename)
        except Exception as e:
            msg = "Error occured here"
            logger.error(e)
            return (False, msg, str(e))

    def download_all_spreadsheets(self):
        downloaded_sheets = []
        for sheet in self.sheets_json:
            if sheet['should_download']:
                spreadsheet_id = sheet['spreadsheet_id']
                filename = sheet['filename']
                success, message, extra = self.download_spreadsheet(
                    spreadsheet_id, filename)
                if success:
                    sheet['filename'] = extra
                    downloaded_sheets.append(sheet)
                    print(message+" at "+extra)
                    logger.debug(message+" at "+extra)
                else:
                    print(message+":"+extra)
                    logger.error((message+":"+extra))
            else:
                msg = f"`{sheet['spreadsheet_name']}` is marked as not to be downloaded"
                print(msg)
                logger.debug(msg)
        return downloaded_sheets
