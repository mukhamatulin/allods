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

## Run

Default:

```powershell
python .\src\allods_hpi_to_xlsx.py
```

Useful args:
- `--output .\my_report.xlsx`
- `--shard-id 601`
- `--timeout 30`[run_daily.bat](run_daily.bat)
- `--state-file .\allods_state_history.json`
- `--yadisk-path /allods/allods_hpi_molodaya_gvardiya.xlsx`
- `--yadisk-token-env YADISK_TOKEN`

## Yandex Disk

To upload the generated XLSX to a stable Yandex Disk path and keep one public link:

```powershell
$env:YADISK_TOKEN='your_oauth_token'
python .\src\allods_hpi_to_xlsx.py `
  --yadisk-path /allods/allods_hpi_molodaya_gvardiya.xlsx
```

Persist the token for your Windows user profile:

```powershell
[Environment]::SetEnvironmentVariable('YADISK_TOKEN', 'your_oauth_token', 'User')
```

Script behavior:
- uploads the file to the same Disk path with overwrite
- publishes the file if needed
- prints the public URL after upload
