# Allods Ratings Export

Project contains a single script for exporting Allods ratings to XLSX.

Source APIs:
- `https://allods.ru/ratings/#/hpi/`
- `https://allods.ru/ratings/#/hpi-astral/`

## Setup

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Microsoft Edge must be installed. Ratings are read through an isolated, headless
Edge session using Playwright, so the site's JavaScript security check can run.
No personal browser profile or saved cookies are used. One session is shared by
all ratings requests during each export and is closed when the export ends.

`run_daily.bat` uses the Python executable configured in its `PY` variable. Install
the dependencies with that interpreter too (installing them in `.venv` alone does
not configure the scheduled task):

```powershell
& 'C:\Users\Rais\AppData\Local\Python\pythoncore-3.14-64\python.exe' -m pip install -r requirements.txt
```

## Run

Default:

```powershell
python .\src\allods_hpi_to_xlsx.py
```

The default run creates two files: `allods_hpi_molodaya_gvardiya.xlsx` for shard `601` and `allods_hpi_nasledie_bogov.xlsx` for shard `101`.

Focus sheets use `Мгла` and `Ройс` for shard `601`, and `Нулёвый` for shard `101`.

Useful args:
- `--shard-id 601` or `--shard-id 101` to generate one server report
- `--output .\my_report.xlsx` with `--shard-id`
- `--timeout 30`
- [run_daily.bat](run_daily.bat)
- `--state-file .\allods_state_history.json` with `--shard-id`
- `--upload-to-yadisk` to publish both configured reports
- `--yadisk-path /allods/custom.xlsx` with `--shard-id` for a custom remote path
- `--yadisk-token-env YADISK_TOKEN`

## Browser security check

Only the official `allods.ru` API is used. A direct HTTP request can return
`403 Forbidden` with a JavaScript security check instead of ratings JSON. The
export opens API pages in Edge and waits for the site's script to finish normally.
`--timeout` limits the total wait for each API page, including its security check.
If the site requires manual CAPTCHA input, an unattended export cannot complete;
the script exits with an error instead of saving the check page as ratings.

The daily launcher accepts the same arguments as the Python script and returns
its exit code. Its log is written under `logs` beside the launcher. Example:

```powershell
.\run_daily.bat --shard-id 601 --output .\test_report.xlsx --state-file .\test_state.json
```

When `YADISK_TOKEN` is set, the daily launcher also uploads the generated reports.

## Yandex Disk upload

To upload both generated XLSX files to their stable Yandex Disk paths:

```powershell
$env:YADISK_TOKEN='your_oauth_token'
python .\src\allods_hpi_to_xlsx.py `
  --upload-to-yadisk
```

Persist the token for your Windows user profile:

```powershell
[Environment]::SetEnvironmentVariable('YADISK_TOKEN', 'your_oauth_token', 'User')
```

Script behavior:
- uploads each file to its configured Disk path with overwrite
- publishes each file if needed
- prints each public URL after upload
