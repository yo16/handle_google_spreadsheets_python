from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from datetime import datetime

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
]

SERVICE_ACCOUNT_FILE = "credentials/service_account.json"

SPREADSHEET_ID = "1TBpDvJOayrXaqNz4BOi2cB2zVMMMkgUoIKYqJCLduq0"
SHEET_NAME1 = "シート1"  # あえて片仮名、アクセスできるかどうか
SHEET_NAME2 = "シート2"


def main():
    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=SCOPES
    )

    service = build("sheets", "v4", credentials=creds)

    sheet = service.spreadsheets()

    # 値の取得(SHEET_NAME1のB2の値を取得)
    result = sheet.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=SHEET_NAME1 + "!B2"
    ).execute()
    values = result.get("values", [])
    print(values)

    # 値を設定(SHEET_NAME2のA1に"Hello, World!"を設定)
    sheet.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=SHEET_NAME2 + "!A1",
        valueInputOption="USER_ENTERED",
        body={
            "values": [["Hello, World!"]]
        }
    ).execute()

    # 最終行と最終列
    result = sheet.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=SHEET_NAME1
    ).execute()
    values = result.get("values", [])
    last_row = len(values)
    last_column = max((len(row) for row in values), default=0)

    # 行列の最後に挿入(SHEET_NAME1の最終列に、今の日時を設定)
    sheet.values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=SHEET_NAME1 + f"!{chr(ord('A') + last_column)}{last_row + 1}",
        valueInputOption="USER_ENTERED",
        body={
            "values": [[datetime.now().strftime("%Y-%m-%d %H:%M:%S")]]
        }
    ).execute()


if __name__ == "__main__":
    main()
