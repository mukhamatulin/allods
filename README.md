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

The default run creates two files: `allods_hpi_molodaya_gvardiya.xlsx` for shard `601` and `allods_hpi_nasledie_bogov.xlsx` for shard `101`.

Useful args:
- `--shard-id 601` or `--shard-id 101` to generate one server report
- `--output .\my_report.xlsx` with `--shard-id`
- `--timeout 30`
- [run_daily.bat](run_daily.bat)
- `--state-file .\allods_state_history.json` with `--shard-id`
- `--upload-to-yadisk` to publish both configured reports
- `--yadisk-path /allods/custom.xlsx` with `--shard-id` for a custom remote path
- `--yadisk-token-env YADISK_TOKEN`

## Yandex Disk

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
