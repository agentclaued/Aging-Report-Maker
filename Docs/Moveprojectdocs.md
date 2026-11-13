# Move WalMartSummary to Another Windows PC

## Purpose

This guide moves the WalMartSummary project to another Windows PC where Codex and Python are installed. It is written for both an operator and a Codex AI maintainer. Follow the **production runtime** path when the new PC only needs to run the app. Follow the **full development workspace** path when the new PC must also maintain source code, history, documentation, and prior work.

The current primary service is **Wal Mart Summary**. Its normal upload and report route is file-only Python processing: it does not need SQL Server. The project also retains optional/manual SQL workflow tools and additional marketplace flows (Amazon Accounts, Regular Accounts, and New Agening); those have extra requirements listed below.

## Choose what to copy

| Need on the new PC | Copy this | What it includes |
| --- | --- | --- |
| Run the app only | `WallMartSummary Ready to Deploy/WallMartSummary App/` | Clean application runtime, active profiles, templates, launchers, sanitized settings, and empty writable folders. |
| Continue all development with Codex | The entire current project folder | Runtime plus documentation, edit history, source backups, previous outputs, working files, and deployment bundle. |

For a true full project move, copy the entire root folder. Do not copy only `app/`: the app also depends on `templates/`, `companies/`, `app_runtime.env`, the launcher files, and `requirements.txt`.

## Before copying on the old PC

1. Stop the running app with `stop_wallmartsummary_app.bat`, or close the window that started it.
2. Wait until no report download or upload is running.
3. Make one complete backup copy of the root folder before transfer. Keep that original copy until the target PC has passed the validation checklist.
4. Decide whether historical customer uploads, generated workbooks, and logs are allowed to move. They can contain sensitive business data. Copy them only under the organization’s retention and access rules.
5. Do not run the project from a ZIP file. Extract it first on the target PC.

### What must travel with a full workspace

- `app/` — FastAPI app, file processing, SQL helper code, workbook generation, logs.
- `templates/` — active browser UI, especially `templates/index.html`.
- `companies/` — active Wal Mart profile and workbook style template; also preserved Amazon, Regular Accounts, and archived/reference profiles.
- `app_runtime.env` — environment-specific runtime settings; review before use on the new PC.
- `requirements.txt` — Python package list.
- `start_wallmartsummary_app.bat` and `stop_wallmartsummary_app.bat` — approved launch and stop commands.
- `output/` and `duplicate_staging/` — writable runtime folders. Preserve them when historical output matters; otherwise create/keep them empty.
- `edit log/` — append-only AI edit history and run logs.
- `Documention Both/`, `README.md`, `Docs/HUMAN_GUIDE.md`, `AGENTS.md`, `AI EDIT CONFIG BY EACH COMPANY TO UPDATE.md`, and this guide — operating and maintainer context.
- `WallMartSummary Ready to Deploy/` — clean deployment bundle and deployment documentation.

Keep `extras/` with the full workspace as reference material, but do not move files from `extras/` back into the active runtime unless a specific change requires them.

## Target-PC requirements

### Required for the primary Wal Mart Summary report

- Windows with a supported 64-bit Python installation.
- Codex installed if the PC will be used for AI-assisted maintenance. Codex is not needed merely to run the app.
- A writable permanent folder, for example `D:\WalMartSummary`. Do not use Downloads, a read-only network share, or a synchronized folder that may lock Excel files.
- Python packages from `requirements.txt`: FastAPI, Uvicorn, python-multipart, pandas, openpyxl, xlsxwriter, Jinja2, and pyodbc.

### Required only for optional/manual features

- **SQL workflow tools:** SQL Server access, the Microsoft ODBC Driver named in `app_runtime.env` (currently `ODBC Driver 17 for SQL Server`), `pyodbc`, and Windows/database permissions. The standard Wal Mart upload/report path does not connect to SQL Server.
- **Regular Accounts native PivotTables:** Microsoft Excel installed for the Windows account that runs the app. This is not needed for Wal Mart Summary or Amazon processing.
- **Remote browser access:** a chosen LAN host/IP and an inbound Windows Firewall rule limited to trusted users. Never expose this development server directly to the public internet.

## Copy and install steps

1. Copy the chosen folder using an external drive, approved network share, or approved transfer service.
2. On the target PC, extract/place it in a permanent writable location, for example `D:\WalMartSummary`.
3. Open Command Prompt in that folder.
4. Find the Python interpreter you intend to use:

```bat
py -c "import sys; print(sys.executable)"
```

5. Install packages using that exact interpreter:

```bat
"C:\Path\To\python.exe" -m pip install --upgrade pip
"C:\Path\To\python.exe" -m pip install -r requirements.txt
```

6. Open `app_runtime.env` in a text editor and set the values for the target PC. Do not put passwords or secrets in this file.
7. Start the app by double-clicking `start_wallmartsummary_app.bat`, or run it from Command Prompt.
8. Keep the launcher window open while the app is running.

## Configure `app_runtime.env`

`app_runtime.env` is the authoritative runtime configuration. `app/runtime_config.py` reads it, and both launcher files take `HOST`, `PORT`, and `PYTHON_EXE` from it.

Use these settings carefully:

| Setting | Target-PC action |
| --- | --- |
| `PYTHON_EXE` | Set the complete path to the selected `python.exe`. The launcher otherwise checks only its built-in Python 3.14 locations. |
| `HOST` | Use `127.0.0.1` for this PC only, or a specific trusted LAN IP when coworkers must open the app. Use `0.0.0.0` only when intentionally listening on every adapter. |
| `PORT` | Keep `8040` if available, or choose an unused approved port. The start and stop launchers read this value. |
| `OUTPUT_DIR`, `TEMPLATES_DIR` | Keep the relative defaults (`output`, `templates`) unless code is deliberately changed and tested. |
| `DEFAULT_COMPANY_CODE` | Keep `wallmartsummary`. Do not add multi-company logic unless specifically requested. |
| `SQL_SERVER`, `SQL_DATABASE`, `SQL_DRIVER`, `TRUSTED_CONNECTION` | Leave unchanged for file-only Wal Mart reporting. Change only when optional SQL tools are intended and the target SQL environment is approved. |

The current source defaults are: app title `WalMartSummary SQL FastAPI`, database `WallMartSummary`, port `8040`, and company code `wallmartsummary`.

## First-run validation

After startup, open the following on the target PC:

```text
http://127.0.0.1:8040/health
```

It should return JSON with `"status":"ok"`. If the configured port is not 8040, use that port instead.

Then run one approved non-sensitive Wal Mart sample through the UI:

1. Select **Wal Mart Summary**.
2. Upload an XLSX, XLS, or CSV with the required columns.
3. Click **Process Wal Mart File**.
4. Confirm that the browser downloads `Wal Mart Status Summary YYYY-MM-DD HH-MM-SS.xlsx`.
5. Open the workbook and confirm that `WallMartSummary_Sheet`, the status sheets, dates, totals, and currency formatting look correct.

The expected primary input columns are `Acct`, `Div`, `Name`, `Ref#`, `Date`, `Due Date`, `Order#`, `Current`, `1-30 Days`, `31-60 Days`, `61-90 Days`, `Over 90 Days`, `Claim#`, and `Comment`. `Ref Type` is optional and will be refreshed from `Ref#`; extra columns are retained.

## Operating notes after the move

- Start: `start_wallmartsummary_app.bat`.
- Stop: close the launch window cleanly when possible, or run `stop_wallmartsummary_app.bat`. It terminates the process listening on the configured port, so use a port dedicated to this app.
- Outputs: generated workbooks go to `output/` and may be downloaded by the browser.
- Duplicate uploads: `duplicate_staging/` is part of normal runtime handling and must remain writable.
- The primary Wal Mart report is Python/file-only. Do not diagnose an SQL connection before checking the file, columns, dependencies, and app log.
- Keep a tested previous complete folder as rollback. For rollback, stop the app, retain the failed copy for diagnosis, restore the earlier full folder with its matching configuration, then repeat `/health` and the sample-report test.

## Codex AI maintainer handoff

Give Codex this project folder as the workspace root. Before any code or documentation edit, Codex must read `AGENTS.md`. The project rules require it to read these files before editing:

- `README.md`
- `AI EDIT CONFIG BY EACH COMPANY TO UPDATE.md`
- `app/main.py`
- `app/workflow.py`
- `app/runtime_config.py`
- `app/adapters/company_loader.py`

Key maintenance rules:

- Back up important files before each edit, using a nearby `.bak`, `.backup`, or dated `.bak` file.
- Keep runtime work focused on WalMartSummary and prefer small, clear changes.
- Ask before unrelated or assumption-changing edits.
- Do not create duplicate template versions, add unrequested multi-company logic, or move runtime folders before code is updated.
- Keep SQL-based Step 07 export behavior SQL-driven unless a change is explicitly approved.
- Keep UI, user documentation, workflow documentation, and append-only `edit log/ai edit history.txt` aligned when documentation changes.
- The AI history is append-only: never remove prior entries. Add date, time, purpose, changed files, backups, and route/button/doc impact.
- Do not copy `extras/` material into the active runtime without an explicit reason.

For company-clone work, use `AI EDIT CONFIG BY EACH COMPANY TO UPDATE.md` as the single source of configuration changes. The active company remains `companies/wallmartsummary/`.

## Current architecture and source of truth

```text
Browser -> FastAPI app/main.py -> company profile + processor -> formatted XLSX -> output/download
                     |                 |
                     |                 +-- companies/wallmartsummary/ (active profile/template)
                     +-- templates/index.html (UI)
                     +-- app_runtime.env (environment settings)
```

- `app/main.py` owns HTTP routes and coordinates processors.
- `app/workflow.py` contains the older/manual Step 01–03 workflow helpers and file preparation.
- `app/adapters/wallmart_python_processor.py` is the file-only Wal Mart report processor.
- `companies/wallmartsummary/company.json` identifies the active profile; its approved style asset is `companies/wallmartsummary/wallmart_status_style_template.xlsx`.
- `requirements.txt` is the Python dependency source of truth.
- `templates/index.html` is the active UI.
- `output/wallmartsummary_workflow_documentation.html` and `Docs/HUMAN_GUIDE.md` are user-facing workflow references.

## Troubleshooting

| Symptom | Check first |
| --- | --- |
| `Python was not found` | Set the exact `PYTHON_EXE` path in `app_runtime.env`; run `"path\to\python.exe" --version`. |
| App will not start | Run `"path\to\python.exe" -m pip install -r requirements.txt`; read the launcher window error; confirm port is unused. |
| `/health` does not open | Confirm the launcher window is still open, use the configured port, and check Windows Firewall only if accessing from another PC. |
| Browser cannot access from another PC | Set a valid target `HOST`, test locally first, then add a narrowly scoped inbound firewall rule for the chosen TCP port. |
| Wal Mart upload fails | Verify required raw columns and input file type. The normal route does not require SQL Server. |
| SQL/manual action fails | Verify SQL Server reachability, ODBC driver name, trusted-connection permissions, and the intended database; do not alter file-only report behavior to solve it. |
| Regular Accounts pivot creation fails | Confirm Microsoft Excel is installed and usable by the same Windows account that runs the app. |
| Output cannot be written | Check folder permissions and whether Excel or a sync tool has locked a file in `output/` or `duplicate_staging/`. |

## Security and data handling

- Do not transfer secrets, passwords, or API tokens in a shared ZIP, email, screenshots, or this documentation.
- Treat customer uploads, outputs, and logs as business data. Copy only through approved storage and restrict access.
- The supplied FastAPI launcher does not document built-in authentication or HTTPS termination. Keep it on a trusted internal network unless approved infrastructure adds TLS, authentication, and access controls.
- Use a dedicated Windows account with only the folder and optional SQL permissions required for the chosen workflow.

## Completion checklist

- [ ] Original project remains preserved until validation succeeds.
- [ ] Full workspace or clean runtime bundle copied to a permanent writable target folder.
- [ ] `app_runtime.env` reviewed for the target PC; `PYTHON_EXE`, `HOST`, and `PORT` set deliberately.
- [ ] Dependencies installed using the selected Python executable.
- [ ] `output/` and `duplicate_staging/` are writable.
- [ ] App started with `start_wallmartsummary_app.bat`.
- [ ] `/health` returned `status: ok`.
- [ ] One approved Wal Mart sample produced a valid workbook.
- [ ] Optional SQL, remote access, and Regular Accounts requirements were tested only if those features are needed.
- [ ] The old PC is not retired until the target copy is confirmed working and backed up.

## Related project documents

- `README.md` — current runtime behavior and input/report rules.
- `Docs/HUMAN_GUIDE.md` — operator workflow.
- `AGENTS.md` — mandatory rules for Codex AI edits.
- `AI EDIT CONFIG BY EACH COMPANY TO UPDATE.md` — company clone/update configuration map.
- `WallMartSummary Ready to Deploy/DEPLOYMENT_GUIDE.md` — concise clean-runtime deployment guide.
- `WallMartSummary Ready to Deploy/RELEASE_CHECKLIST.md` — release validation checklist.
- `WallMartSummary Ready to Deploy/CONFIGURATION_AND_SECURITY.md` — security and configuration controls.

Last updated: 2026-10-04 (Asia/Karachi).
