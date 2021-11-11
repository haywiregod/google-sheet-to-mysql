# Google Sheets to MySQL
A python module to download google sheets from your account and insert their data into tables

## Settings
1) View `sheets.json.example`
1) Copy `env.example` to `.env`
1) Download your `client_secrets.json` from https://console.cloud.google.com/apis/credentials
2) `client_secrets.json` should be at the same level as all the scripts


## How to run?
1) Create virtual environment. (eg: Use `virtualenv venv`)
2) Install requirements using `pip install -r requirements.txt`
3) run `python sync.py` 
