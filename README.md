# SQL Plan Viewer

A lightweight Windows GUI, written in PowerShell, for viewing and analysing SQL Server execution plans. Open a `.sqlplan` file, paste plan XML, or capture an estimated or actual plan straight from a server, then explore it as a graphical plan, a tree, a list of the most expensive operators, and a list of warnings and findings.

It also includes a **SQL browser** for registering instances and browsing their views and stored procedures, so you can go from an object to its plan in a couple of clicks.

<!-- Add a screenshot here, e.g. ![SQL Plan Viewer](docs/screenshot.png) -->

## Features

**Loading plans**
- Open `.sqlplan` / showplan XML files, or drag and drop them onto the window
- Paste plan XML from the clipboard
- Capture a plan from SQL Server: estimated (query is not executed) or actual (query is executed)
- Multiple batches (`GO` separators) and multi-statement plans are supported
- Save the current plan as a `.sqlplan` file

**Analysis**
- **Graphical plan** - operators shaded by cost, arrows scaled by row count, warning badges, hover tooltips, zoom (Ctrl+wheel), fit-to-window, export as PNG
- **Operator properties** - cost, rows (estimated and actual), runtime, seek predicates, predicates, output columns, warnings
- **Operator tree** and **Top operators** views, kept in sync with the graphical plan
- **Warnings and findings** - plan warnings (implicit conversions, spills, missing statistics, memory grants, ...), missing indexes, operators that account for 25% or more of the cost, large scans (10,000+ rows), key lookups that run 1,000+ times, and row estimates that are 10x or more off the actual rows
- **Missing indexes** - with a ready-to-copy `CREATE INDEX` script for each suggestion
- **Statement and parameters** - cost, DOP, memory grant, compile stats, query hash, and compiled vs runtime parameter values (useful for spotting parameter sniffing)

**SQL browser**
- Register SQL Server instances (Windows or SQL authentication)
- Browse instance > database > Views / Stored procedures
- View object definitions, copy them, and send an object to the plan dialog
  - views pre-fill `SELECT * FROM [schema].[view];`
  - stored procedures pre-fill `EXEC [schema].[proc] @param = NULL, ...;` (replace the NULLs with real values)

## Requirements

- Windows 10 / 11 (WinForms)
- Windows PowerShell 5.1 (included with Windows) or PowerShell 7 on Windows
- Network access to the SQL Server instances you want to query (not needed for opening plan files)
- SQL permissions: `SHOWPLAN` on the database to capture plans, and `VIEW DEFINITION` to see object definitions in the SQL browser

## Getting started

1. Download the files from this repository into one folder.
2. Files downloaded from the internet are blocked by Windows. Unblock them once:

   ```powershell
   Get-ChildItem . | Unblock-File
   ```

3. Run it:

   ```powershell
   .\SqlPlanViewer.ps1
   .\SqlPlanViewer.ps1 C:\temp\slow-query.sqlplan
   ```

The PowerShell console window is hidden automatically. If the app does not start, run it with `-ShowConsole` to see the error:

```powershell
.\SqlPlanViewer.ps1 -ShowConsole
```

For a launch with no console flash at all, double-click `SqlPlanViewer.vbs` (keep it in the same folder as the script). VBScript is being phased out of recent Windows 11 builds; if it is unavailable, use the `.ps1` directly.

A small example plan, `sample.sqlplan`, is included to try the viewer.

### Getting a plan from SQL Server

Use **Get plan from SQL Server...** (Ctrl+R), enter the server, authentication and database, paste the query, and choose:

- **Estimated plan** - the query is compiled but not executed. Safe to use anywhere.
- **Actual plan** - the query **is executed** and its results are discarded. By default it is wrapped in a transaction that is rolled back afterwards, and you are asked to confirm first. Rolling back does not undo side effects outside the transaction (emails, external calls, identity or sequence values, linked servers, ...).

`SET SHOWPLAN_XML` does not honour `USE`, so choose the database in the dialog instead.

### Keyboard shortcuts

| Shortcut | Action |
| --- | --- |
| Ctrl+O | Open a plan file |
| Ctrl+V | Load plan XML from the clipboard |
| Ctrl+R | Get a plan from SQL Server |
| Ctrl+wheel, Ctrl+`+` / Ctrl+`-` | Zoom the graphical plan |

## Building an .exe

`Build-SqlPlanViewerExe.ps1` compiles the script into `SqlPlanViewer.exe` (with the icon and no console window) using the open-source [ps2exe](https://github.com/MScholtes/PS2EXE) module.

1. Put `SqlPlanViewer.ps1`, `SqlPlanViewer.ico`, `Build-SqlPlanViewerExe.ps1` and `Build-SqlPlanViewerExe.cmd` in one folder (and unblock them as above).
2. Double-click `Build-SqlPlanViewerExe.cmd`, or run `.\Build-SqlPlanViewerExe.ps1` from Windows PowerShell 5.1.

The first run installs `ps2exe` for the current user from the PowerShell Gallery. The window stays open at the end, and everything is logged to `build.log`.

Options:

```powershell
.\Build-SqlPlanViewerExe.ps1 -Version 1.1.0.0
.\Build-SqlPlanViewerExe.ps1 -CertThumbprint <thumbprint>   # sign the exe with a code-signing certificate
```

Things to know about the exe:

- It is not standalone native code. It embeds the script (anyone can extract it) and runs it on the Windows PowerShell 5.1 that ships with Windows.
- Executables produced by ps2exe are sometimes flagged by antivirus or SmartScreen. Building the exe yourself, and signing it if you can, helps.

## Security notes

- Connections are always encrypted (`Encrypt=True`). **Trust server certificate** is off by default; only tick it for servers with self-signed certificates, because it disables certificate validation.
- Passwords are held in memory for the session only and are never written to disk. They are cleared when the app closes.
- Settings and the instance list are saved as plain JSON in `%APPDATA%\SqlPlanViewer\` (`settings.json`, `instances.json`): server, database, authentication type, login and options only. Query text is not saved.
- Plan files are parsed defensively: DTDs and external entities are blocked, files over 100 MB are rejected, and plans nested more than 200 operators deep are rejected. Plan content is only parsed and drawn; it is never executed.
- Metadata queries in the SQL browser are fixed or parameterized, and object names in generated scripts are bracket-escaped.
- **Plans contain sensitive information**: query text, parameter values, schema and server names. Treat saved plans, PNG exports and plan XML you share accordingly.
- The tool runs with your own SQL permissions. Use a least-privilege login, especially for actual plans.
- The scripts are unsigned. Keep them in a location only you can write to, and sign them if your policy requires it.

## Limitations

- Windows only.
- Uses the legacy `System.Data.SqlClient` provider.
- A running query cannot be cancelled from the UI; the timeout in the dialog is the limit.
- Cursor statements and statements without a query plan are not shown.
- Row-count and cost figures come from the plan XML as-is; "Cost %" is each operator's own cost as a share of the statement's total estimated cost.

## Repository contents

| File | Purpose |
| --- | --- |
| `SqlPlanViewer.ps1` | The application |
| `SqlPlanViewer.ico` | Application icon |
| `SqlPlanViewer.vbs` | Optional launcher with no console flash |
| `Build-SqlPlanViewerExe.ps1` / `.cmd` | Builds `SqlPlanViewer.exe` with ps2exe |
| `sample.sqlplan` | Example plan to try |

## Disclaimer

This is an independent project and is not affiliated with or endorsed by Microsoft. SQL Server is a trademark of Microsoft Corporation.
