# Setup in Office PC

How WalMartSummary (Wal Mart / Amazon / Regular Accounts file reports) runs on the office PC, shows in the Easy Solutionz tools hub, and how to debug or change it.

- Office PC address: `192.168.2.14`  |  Tools hub (portal): `http://192.168.2.14:8500`
- This app: port **8507** -> `http://192.168.2.14:8507`
- App folder on the office PC: `C:\Repos\walmart_Ageing_Office pc_Ans`
- Source (development) folder: `D:\WalMartDashBoard v2 Claude\Ready to deploy`, copy for the office PC: `...\Ready to deploy\Office pc`
- Ports in use by the hub: 8500 portal, 8502 Ageing, 8503 QVC/HSN, 8504 Settlements, 8505 Dropship, 8506 Walmart Fines, **8507 Wal Mart Summary**.

---

## 1. What was built / changed (summary of the work so far)

| Area | What |
|---|---|
| App Control Center path bug | The bat/ps1/pyw files assumed the folder was two levels below the project root. It is one level. All root lookups now use `..` (`%~dp0..`, `'..'`, `parents[1]`). |
| BOM | Four bat files had a UTF-8 BOM that broke `@echo off`. Removed. Bat files must be saved **without BOM**. |
| Tray icon | Visible only while the app URL answers. Menu: Open app, Restart server, Open log folder, **Stop the app**. |
| Boot/restart auto-start | `22` (boot, no recovery), `33` (login, tray, recovery), **`55`** (PC start/restart, hidden, with recovery), **`66`** (PC start, visible console window like the other tools). |
| Renames | `22 1 ON - Install Auto-start at PC boot (no recovery).bat`, `22 2 OFF - UnInstall Auto-start at PC boot (no recovery).bat`. |
| Output files | Generated reports (Wal Mart status, Amazon workbook, Regular Accounts report) are deleted from `output/` after the browser finishes downloading. |
| Full Ageing input | Upload the Regular Accounts style workbook (sheet `Full Ageing`, header row with `Acct` + `Ref Type`, title rows above are skipped). Wal Mart keeps Acct `WC9085, WA9085, SA9200, WM9263`. Amazon keeps `AM9285, AM9286`, optional `ME9145`. |
| Amazon checkboxes | Amazon upload box has three checkboxes: AM9285, AM9286 (checked by default), ME9145 (optional). |
| Health address | `/_stcore/health` added (same as `/health`) so the hub portal can show Online/Offline. |
| Office pc folder | Runtime copy for the office PC with `HOST=0.0.0.0`, `PORT=8507`. |

Rules for creating or changing the control center are also in the project `README.md` ("App Control Center - rules").

---

## 2. Install on the office PC (step by step)

1. **Copy** the contents of the `Office pc` folder into `C:\Repos\walmart_Ageing_Office pc_Ans`. These must sit directly in that folder: `app`, `companies`, `templates`, `App Control Center`, `app_runtime.env`, `requirements.txt`.
2. **Python** 3.12/3.13 installed (tick "Add python.exe to PATH"). Check: `python --version`.
3. **Packages** (the 66 installer does this into a `venv`; manual way):
   ```
   cd /d "C:\Repos\walmart_Ageing_Office pc_Ans"
   python -m pip install -r requirements.txt
   ```
4. **Firewall** (Admin Command Prompt, once):
   ```
   netsh advfirewall firewall add rule name="WalMartSummary 8507" dir=in action=allow protocol=TCP localport=8507 profile=domain,private
   ```
5. **Test by hand** (new Command Prompt, not a hub window):
   ```
   cd /d "C:\Repos\walmart_Ageing_Office pc_Ans"
   python -m uvicorn app.main:app --host 0.0.0.0 --port 8507
   ```
   Expect `Uvicorn running on http://0.0.0.0:8507`. Open `http://192.168.2.14:8507`. Ctrl+C stops it.
6. **Choose ONE way to start it automatically** (only one may own port 8507):

   | File in `App Control Center` | Behaviour |
   |---|---|
   | `11 1 ON - Start in background.bat` | Hidden, tray icon, until log off/restart. |
   | `33 1 ON - Auto-start at login (tray, recovery).bat` | Starts at your login, restarts on crash. |
   | `55 1 ON - Install Auto-start on PC restart (with recovery).bat` | Starts at PC start/restart, hidden, restarts on crash. Admin. |
   | **`66 1 ON - Install Auto-start with console window (like other tools).bat`** | Same style as the other hub tools: visible console window ~45 s after boot, own `venv`. closing the window or a crash is repaired within 10 minutes. Admin, asks for the Windows password. |

   Before installing one, stop the others: `11 2 OFF`, `55 2 OFF`, `66 2 OFF`, `33 2 OFF`, `22 2 OFF`.
7. **Task for 66:** `Tools-walmart_summary`. Check: `schtasks /Query /TN Tools-walmart_summary`. Run now: `schtasks /Run /TN Tools-walmart_summary`.

---

## 3. Show it in the tools hub portal

The portal (`C:\Repos\tools_hub\portal\app.py`) reads `C:\Repos\tools_hub\tools.json` on every page load (no restart needed) and checks Online by calling `http://localhost:<port>/_stcore/health`.

Add this entry to `tools.json` (comma after the previous entry, **no** comma after the last one):

```json
  {
    "id": "walmart_summary",
    "name": "Wal Mart Summary",
    "description": "Upload a Wal Mart, Amazon or Regular Accounts file, or a Full Ageing workbook, and download the status report.",
    "port": 8507,
    "status": "live",
    "departments": ["ar"]
  }
```

- Leave out `repo`, `folder`, `app`: the hub's `auto_update.ps1` / `deploy_to_server.ps1` use them to clone and update Git repos, and this app is updated by copying files.
- JSON has **no comments**. If you need a note, add an unused key such as `"_comment": "..."`.
- Validate: `python -c "import json;json.load(open(r'C:\Repos\tools_hub\tools.json',encoding='utf-8'));print('valid')"`
- Confirm the entry is saved: `findstr walmart_summary C:\Repos\tools_hub\tools.json`
- Hard-refresh the portal: Ctrl+F5. Check "All tools" and the AR department.
- `auto_update.ps1` pulls the hub `live` branch every 5 minutes; a local-only edit can be overwritten. For a permanent change commit it to the `tools_hub` repo and `git push origin main:live` (your decision).

---

## 4. Using the app

- **Wal Mart Summary:** upload the Wal Mart inquiry file **or** a Full Ageing workbook. Full Ageing rows are filtered by `Acct` to `WC9085, WA9085, SA9200, WM9263`. Needed columns: Acct, Div, Name, Ref#, Date, Due Date, Order#, Current, 1-30 Days, 31-60 Days, 61-90 Days, Over 90 Days, Claim#, Comment (Ref Type optional).
- **Amazon Accounts:** upload the Amazon status file **or** a Full Ageing workbook. Tick the accounts to include: AM9285, AM9286 (default), ME9145 (optional). Needs Acct, Ref Type, Date, Order#, Total, Comment, and CR rows with valid dates for the CR status.
- **Full Ageing sheet rules:** sheet name `Full Ageing`; the header row (with `Acct` and `Ref Type`) must be within the first 15 rows.
- **Downloads:** the generated report downloads once and is then deleted from `output/`. Run the report again if you need another copy.

---

## 5. Updating the app later (editing)

1. Edit in the source folder `D:\WalMartDashBoard v2 Claude\Ready to deploy` (backup first, e.g. `file.2026-11-05.bak`; add a note to `edit log\ai edit history.txt`).
2. Copy the same changed files into `Office pc` (keep both in sync) and then to `C:\Repos\walmart_Ageing_Office pc_Ans` on the office PC.
3. **Restart the app** so Python loads the change:
   - 11/55: tray icon -> Restart server, or `11 2 OFF` then `11 1 ON`
   - 66: close the console window, then `schtasks /Run /TN Tools-walmart_summary`
   - by hand: Ctrl+C and run the uvicorn command again
4. Do not copy `output`, `logs`, `venv` or `.bak` files between PCs.
5. Verify: `curl http://localhost:8507/_stcore/health` returns `{"status":"ok",...}`.

Where things are:

| Change | File |
|---|---|
| Routes, upload forms, downloads | `app\main.py` |
| Wal Mart report logic | `app\adapters\wallmart_python_processor.py` |
| Amazon report logic | `app\adapters\amazon_file_processor.py` |
| Full Ageing reader | `app\adapters\full_ageing_source.py` |
| Regular Accounts logic | `app\adapters\regular_accounts_processor.py` |
| Screen layout/wording | `templates\index.html` |
| Host/port | `app_runtime.env` (office PC: `HOST=0.0.0.0`, `PORT=8507`) |
| Tray icon app | `App Control Center\background_tray.pyw` |

---

## 6. Debugging

Always start with the log: `C:\Repos\walmart_Ageing_Office pc_Ans\logs\server.log` (hidden modes) or the console window (mode 66).

| Symptom | Check / fix |
|---|---|
| Nothing on port 8507 | `netstat -ano \| findstr :8507` blank = not running. Run the uvicorn command by hand and read the error. |
| `No module named ...` | `python -m pip install -r requirements.txt` (for mode 66 use `venv\Scripts\python.exe -m pip install -r requirements.txt`). |
| `Could not import module "app.main"` | Files are one folder too deep. `app` and `app_runtime.env` must be directly in the app folder. |
| Works on the office PC, not from other PCs | Firewall rule missing (step 2.4). Make sure `HOST=0.0.0.0`. |
| Another Streamlit window ("streamlit.exe") is open | That is a different tool (8502, 8503...). The Wal Mart app is a separate window/process. |
| Portal card missing | Entry not in `C:\Repos\tools_hub\tools.json` (run `findstr`), invalid JSON, or edit overwritten by `auto_update`. Check "All tools" count. |
| Portal card says Offline but the app opens | The app must answer `/_stcore/health`. `curl http://localhost:8507/_stcore/health` must be 200. Copy the new `app\main.py` and restart the app. |
| 11 says "started" but nothing runs | An old tray process holds the one-instance lock. Run `11 2 OFF` (or `stop_background.ps1`) first. |
| Bat shows `n╗┐@echo` | BOM in the bat file. Re-save as ANSI/UTF-8 **without** BOM. |
| `The variable cannot be validated ... MSFT_TaskExecAction` | Old `*.ps1`: `$action` clashes with the `-Action` parameter. Use the fixed scripts (`$taskAction`). Check: `findstr /C:"taskAction" window_autostart.ps1`. |
| `schtasks /Query` says "cannot find the file" | The task was never created; run the installer script directly and read the error: `powershell -NoProfile -ExecutionPolicy Bypass -File "...\App Control Center\window_autostart.ps1" -Action Install`. A wrong or empty Windows password also prevents task creation. |
| Port already in use | `netstat -ano \| findstr :8507` then `taskkill /PID <PID> /F`; stop other start modes. |
| Report fails: missing column | The message lists the column. Rename in the source file or tell the developer. |
| Amazon: "no valid Date values in Ref Type CR rows" | Existing rule: the file needs CR rows with valid dates. |
| Full Ageing: "no header row with both Acct and Ref Type" | Header not in the first 15 rows, or the columns are named differently. |
| Full Ageing: "no rows for Acct ..." | Those account codes are not in the `Acct` column (or the wrong accounts are ticked). |
| Console window (66) not visible | The task user must be logged in at the PC. The app still runs without the window. |

Useful commands:

```
schtasks /Query /TN Tools-walmart_summary
schtasks /Run /TN Tools-walmart_summary
schtasks /End /TN Tools-walmart_summary
netstat -ano | findstr :8507
curl http://localhost:8507/_stcore/health
type "C:\Repos\walmart_Ageing_Office pc_Ans\logs\server.log"
```

---

## 7. Remove / roll back

- Remove the console-window task: `66 2 OFF - UnInstall Auto-start with console window.bat` (then close the window).
- Remove the restart task: `55 2 OFF ...`. Remove login auto-start: `33 2 OFF ...`. Remove boot task: `22 2 OFF ...`.
- Remove the portal card: delete the `walmart_summary` entry from `tools.json` (fix commas, re-validate JSON).
- Restore a file: copy its `*.bak` back over it (backups exist only in the source folder, not in `Office pc`).
