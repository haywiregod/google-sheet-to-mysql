import os
from types import prepare_class
from excel_to_sql import ExcelToSQL
from download_google_sheets import SheetDownloader
from logger import logger
from dotenv import load_dotenv

logger.info('Start Sync')
load_dotenv()


def main():
    logger.info('Start Main')
    sheet_location = os.environ.get("SHEET_LOCATION", "sheets.json")
    print(sheet_location)
    sheet_location = os.path.abspath(sheet_location)
    print(sheet_location)
    logger.debug(f"SHEET_LOCATION: {sheet_location}")
    sheet_downloader = SheetDownloader(sheet_location)
    downloaded_sheets = sheet_downloader.download_all_spreadsheets()
    # print(downloaded_sheets)
    for sheet in downloaded_sheets:
        wb = ExcelToSQL(sheet['filename'])
        for moveable_sheet in sheet['move_to_sql']:
            sheet_name = moveable_sheet['sheet_name']
            consider_columns = moveable_sheet['consider_columns']
            table_name = moveable_sheet['table_name']
            host = moveable_sheet['host']
            database = moveable_sheet['database']
            user = moveable_sheet['user']
            password = moveable_sheet['password']

            prepared_sheet = wb.get_prepared_sheet(
                sheet_name, consider_columns)
            wb.insert_into_table(host, database, user,
                                 password, table_name, prepared_sheet, True)
    logger.info('Exit Main')


if __name__ == '__main__':
    main()
    logger.info('Exit Sync')
