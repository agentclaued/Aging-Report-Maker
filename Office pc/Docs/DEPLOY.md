# WalMartSummary deployment — Windows Server 2019

Use this guide to install the current WalMartSummary application on a new Windows Server 2019 computer. The application runs with Python, FastAPI/Uvicorn, and (for WalMartSummary) Microsoft SQL Server.

## 1. Copy the application

Copy the complete application folder to a permanent server location, for example:

```text
D:\WalMartSummary
```

Keep all folders and files together, including `app`, `companies`, `templates`, `output`, `requirements.txt`, `app_runtime.env`, and both `.bat` launcher files.

Do not run the application directly from a zipped folder, Downloads, or a network location that the service account cannot write to.

## 2. Install Python (64-bit)

1. Download and install a current **64-bit Python 3.11, 3.12, or 3.13** release from [python.org](https://www.python.org/downloads/windows/).
2. During setup, select **Add python.exe to PATH**.
3. Open a new Command Prompt and confirm the installation:

```bat
py --version
py -m pip --version
```

The app launcher does not depend on PATH once `PYTHON_EXE` is set, but selecting the PATH option makes installation and troubleshooting easier.

## 3. Set the project Python path

Open `app_runtime.env` in Notepad. Change this line:

```env
PYTHON_EXE=
```

to the full path of the installed `python.exe`. Typical examples are:

```env
PYTHON_EXE=C:\Users\Administrator\AppData\Local\Programs\Python\Python313\python.exe
```

or, for an all-users installation:

```env
PYTHON_EXE=C:\Program Files\Python313\python.exe
```

To find the exact path, run this in Command Prompt:

```bat
py -c "import sys; print(sys.executable)"
```

Paste the displayed path after `PYTHON_EXE=`. Do not add quotation marks. Save the file.

`PYTHON_EXE` in `app_runtime.env` is the only Python setting required by this application. You do not need to create a Windows environment variable.

## 4. Install Python packages

Open Command Prompt **as the Windows user that will run the app**, then run:

```bat
cd /d D:\WalMartSummary
"C:\Program Files\Python313\python.exe" -m pip install --upgrade pip
"C:\Program Files\Python313\python.exe" -m pip install -r requirements.txt
```

Replace both example Python paths with the same exact path used in `app_runtime.env`.

## 5. Install the SQL Server ODBC driver

WalMartSummary uses `pyodbc` and requires **Microsoft ODBC Driver 17 for SQL Server**, because that is the configured default in `app_runtime.env`.

Install the x64 driver from Microsoft, then confirm it appears in Windows **Apps & features**. If your server uses a different installed driver, update this config value to its exact name:

```env
SQL_DRIVER=ODBC Driver 17 for SQL Server
```

## 6. Configure SQL Server and the app

Edit `app_runtime.env` to match the server environment:

```env
SQL_SERVER=localhost\SQLEXPRESS
SQL_DATABASE=WallMartSummary
SQL_DRIVER=ODBC Driver 17 for SQL Server
TRUSTED_CONNECTION=yes
HOST=192.168.2.123
PORT=8040
PYTHON_EXE=C:\Program Files\Python313\python.exe
DEFAULT_COMPANY_CODE=wallmartsummary
```

Important:

- `SQL_SERVER`: use `localhost\SQLEXPRESS` only if SQL Server Express is on this same Windows server. Otherwise use the SQL Server host and instance provided by your database administrator, for example `SQLSERVER01\SQLEXPRESS`.
- `SQL_DATABASE`: the `WallMartSummary` database and its stored procedures must already exist before using the WallMartSummary SQL workflow.
- `TRUSTED_CONNECTION=yes`: the Windows account starting the app must have permission to connect to this database. Grant it the required database permissions through SQL Server Management Studio.
- `HOST`: set this to the server's LAN IP address, or `0.0.0.0` to listen on all network adapters. Do not use an IP address that belongs to an old computer.
- `PORT`: leave `8040` unless it is already in use.

## 7. Allow network access

If users will open the app from other computers, create an inbound Windows Firewall rule for the configured TCP port (default `8040`). From an elevated PowerShell window:

```powershell
New-NetFirewallRule -DisplayName "WalMartSummary TCP 8040" -Direction Inbound -Protocol TCP -LocalPort 8040 -Action Allow
```

If you changed `PORT`, use that port number in the command. If the app is used only on the server itself, this rule is not needed.

## 8. Start and test

Double-click `start_wallmartsummary_app.bat`, or run:

```bat
cd /d D:\WalMartSummary
start_wallmartsummary_app.bat
```

Leave that Command Prompt open while the app is running. Open the app in a browser:

```text
http://SERVER-IP:8040/
```

For example: `http://192.168.2.123:8040/`.

To stop it, close the start window or run `stop_wallmartsummary_app.bat`.

## Troubleshooting

### “Python was not found”

The value in `app_runtime.env` is empty, misspelled, or points to a non-existent file. Run:

```bat
py -c "import sys; print(sys.executable)"
```

Copy that exact output into `PYTHON_EXE=` and run the launcher again.

### “No module named uvicorn” or another module error

Packages were installed under a different Python installation. Run `-m pip install -r requirements.txt` using the exact `PYTHON_EXE` path configured in `app_runtime.env`.

### SQL Server or ODBC connection error

Check that SQL Server is running, the database name and instance in `app_runtime.env` are correct, ODBC Driver 17 is installed, and the Windows account running the launcher has SQL permissions.

### Browser cannot open the app from another PC

Confirm the URL uses the new server's current LAN IP, the `HOST` setting permits that IP, TCP port 8040 is allowed through Windows Firewall, and no other program is already using the port.

