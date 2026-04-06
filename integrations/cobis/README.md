# COBIS → Veza OAA Integration

Push user-to-group membership data from a COBIS "Users By Group" Excel report into [Veza's Authorization Graph](https://www.veza.com) via the Open Authorization API (OAA).

## Overview

This connector reads a COBIS "Users By Group" Excel workbook (`.xls` or `.xlsx`) where **each sheet represents one security group** and the listed rows are users belonging to that group. It creates the following Veza entities:

| COBIS Source | Veza OAA Entity |
|---|---|
| Sheet → Group name (row 13, col I) | **Local Group** |
| User row (username in col A) | **Local User** |
| User listed on a sheet | **Group membership** (user ↔ group) |

No permission mapping is performed — only identity-to-group relationships are modeled.

### User Custom Properties

Each local user carries the following custom properties in Veza:

| Property | Type | Source Column |
|---|---|---|
| `department` | STRING | Column M (index 12) |
| `plant` | STRING | Column P (index 15) |
| `is_disabled` | BOOLEAN | Column S (index 18) |
| `disabled_date` | STRING | Column T (index 19) |
| `is_inactive` | BOOLEAN | Column W (index 22) |
| `inactivation_date` | STRING | Column Y (index 24) |
| `created_date` | TIMESTAMP | Column AB (index 27) |
| `created_by` | STRING | Column AC (index 28) |
| `last_modified` | TIMESTAMP | Column AD (index 29) |
| `modified_by` | STRING | Column AF (index 31) |
| `last_login` | TIMESTAMP | T27991126A report, column H (Laatste login) |

## How It Works

1. **Load configuration** — Reads CLI arguments, environment variables, and `.env` file (CLI args take precedence)
2. **Obtain Excel file** — Either from a local/mounted path or by downloading directly from an SMB/CIFS share
3. **Parse workbook** — Iterates over all sheets; for each, extracts the group name (row 13) and user records (rows 15+)
4. **Parse last-logon report** *(optional)* — Reads the T27991126A-style XLSX to extract each user's most recent login timestamp
5. **Build OAA payload** — Creates a `CustomApplication` with local users, local groups, custom properties, group memberships, and `last_login_at`
6. **Push to Veza** — Creates or updates the provider and pushes the application payload

## Prerequisites

- **OS**: Linux (RHEL/CentOS/Fedora or Ubuntu/Debian)
- **Python**: ≥ 3.8
- **Network**: Outbound HTTPS to your Veza tenant; access to the SMB share or local mount
- **Veza**: API key with OAA write permissions
- **SMB access**: Either a mounted CIFS share (via `cifs-utils`) or direct SMB access (via `smbprotocol`)

## Quick Start

### One-command installer

```bash
# From the repository directory:
bash install_cobis.sh

# Non-interactive (for CI/automation):
VEZA_URL=https://yourcompany.vezacloud.com \
VEZA_API_KEY=your_api_key \
COBIS_XLSX_PATH=/mnt/cobis/report.xls \
bash install_cobis.sh --non-interactive
```

## Manual Installation

### RHEL / CentOS / Fedora

```bash
sudo dnf install -y git python3 python3-pip python3-devel cifs-utils

# Clone / copy scripts
sudo mkdir -p /opt/cobis-veza/{scripts,logs}
sudo cp cobis.py requirements.txt /opt/cobis-veza/scripts/

# Virtual environment
sudo python3 -m venv /opt/cobis-veza/scripts/venv
sudo /opt/cobis-veza/scripts/venv/bin/pip install --upgrade pip
sudo /opt/cobis-veza/scripts/venv/bin/pip install -r /opt/cobis-veza/scripts/requirements.txt

# Configure
sudo cp .env.example /opt/cobis-veza/scripts/.env
sudo chmod 600 /opt/cobis-veza/scripts/.env
sudo vi /opt/cobis-veza/scripts/.env   # fill in real values
```

### Ubuntu / Debian

```bash
sudo apt-get update
sudo apt-get install -y git python3 python3-pip python3-venv python3-dev cifs-utils

# Same directory layout and venv steps as RHEL above
sudo mkdir -p /opt/cobis-veza/{scripts,logs}
sudo cp cobis.py requirements.txt /opt/cobis-veza/scripts/
sudo python3 -m venv /opt/cobis-veza/scripts/venv
sudo /opt/cobis-veza/scripts/venv/bin/pip install --upgrade pip
sudo /opt/cobis-veza/scripts/venv/bin/pip install -r /opt/cobis-veza/scripts/requirements.txt
sudo cp .env.example /opt/cobis-veza/scripts/.env
sudo chmod 600 /opt/cobis-veza/scripts/.env
```

### Mounting the SMB Share (recommended)

If you prefer a mounted share over direct SMB access:

```bash
# Install cifs-utils
sudo apt-get install -y cifs-utils   # or: sudo dnf install -y cifs-utils

# Create mount point and credentials file
sudo mkdir -p /mnt/cobis-share
echo "username=svc_cobis_reader" | sudo tee /etc/cobis-smb-credentials
echo "password=YOUR_PASSWORD" | sudo tee -a /etc/cobis-smb-credentials
sudo chmod 600 /etc/cobis-smb-credentials

# Add to /etc/fstab for persistent mount
echo "//fileserver.example.com/cobis-reports /mnt/cobis-share cifs credentials=/etc/cobis-smb-credentials,uid=cobis-veza,gid=cobis-veza,ro,file_mode=0440,dir_mode=0550 0 0" | sudo tee -a /etc/fstab

# Mount now
sudo mount /mnt/cobis-share
```

Then set `COBIS_XLSX_PATH=/mnt/cobis-share/reports/FBIC_Users_By_Group.xls` in your `.env`.

## Usage

```bash
cd /opt/cobis-veza/scripts
source venv/bin/activate
python3 cobis.py [OPTIONS]
```

### CLI Arguments

| Argument | Required | Values / Default | Description |
|---|---|---|---|
| `--env-file` | No | `.env` | Path to environment config file |
| `--log-level` | No | DEBUG/INFO/WARNING/ERROR (INFO) | Logging verbosity |
| `--dry-run` | No | flag | Build payload without pushing to Veza |
| `--save-json` | No | flag | Save OAA JSON payload to local file |
| `--veza-url` | Yes* | env: `VEZA_URL` | Veza tenant URL |
| `--veza-api-key` | Yes* | env: `VEZA_API_KEY` | Veza API key |
| `--provider-name` | No | `COBIS` | Provider name in Veza UI |
| `--datasource-name` | No | `COBIS` | Datasource name in Veza |
| `--xlsx-path` | Yes** | env: `COBIS_XLSX_PATH` | Local path to Excel file |
| `--last-logon-path` | No | env: `COBIS_LAST_LOGON_PATH` | Path to T27991126A last-logon XLSX |
| `--smb-server` | Yes** | env: `COBIS_SMB_SERVER` | SMB server hostname/IP |
| `--smb-share` | Yes** | env: `COBIS_SMB_SHARE` | SMB share name |
| `--smb-path` | Yes** | env: `COBIS_SMB_PATH` | File path within the share |
| `--smb-username` | Yes** | env: `COBIS_SMB_USERNAME` | SMB username |
| `--smb-password` | Yes** | env: `COBIS_SMB_PASSWORD` | SMB password |
| `--smb-port` | No | `445` | SMB port |

\* Not required with `--dry-run`
\** Provide EITHER `--xlsx-path` OR the `--smb-*` group

### Examples

```bash
# Dry run with local file
python3 cobis.py --xlsx-path /mnt/cobis/report.xls --dry-run --save-json

# Dry run with last-logon enrichment
python3 cobis.py --xlsx-path /mnt/cobis/report.xls \
  --last-logon-path /mnt/cobis/T27991126A.xlsx --dry-run --save-json

# Production run with .env file
python3 cobis.py --env-file .env

# Direct SMB access
python3 cobis.py \
  --smb-server fileserver.example.com \
  --smb-share cobis-reports \
  --smb-path "reports/FBIC_Users_By_Group.xls" \
  --smb-username svc_cobis_reader \
  --smb-password "$SMB_PASS" \
  --veza-url https://company.vezacloud.com \
  --veza-api-key "$VEZA_API_KEY"
```

## Deployment on Linux

### Service Account

```bash
sudo useradd -r -s /bin/bash -m -d /opt/cobis-veza cobis-veza
sudo chown -R cobis-veza:cobis-veza /opt/cobis-veza
sudo chmod 700 /opt/cobis-veza/scripts
sudo chmod 600 /opt/cobis-veza/scripts/.env
```

### SELinux (RHEL/CentOS)

```bash
if [[ "$(getenforce 2>/dev/null)" != "Disabled" ]]; then
    sudo restorecon -Rv /opt/cobis-veza/
fi
```

### Cron Scheduling

Create a wrapper script:

```bash
sudo tee /opt/cobis-veza/scripts/run_cobis.sh > /dev/null << 'EOF'
#!/usr/bin/env bash
set -euo pipefail
cd /opt/cobis-veza/scripts
source venv/bin/activate
python3 cobis.py --env-file .env >> /opt/cobis-veza/logs/cobis.log 2>&1
EOF
sudo chmod 700 /opt/cobis-veza/scripts/run_cobis.sh
sudo chown cobis-veza:cobis-veza /opt/cobis-veza/scripts/run_cobis.sh
```

Schedule via `/etc/cron.d/`:

```bash
echo "0 2 * * * cobis-veza /opt/cobis-veza/scripts/run_cobis.sh" | sudo tee /etc/cron.d/cobis-veza
sudo chmod 644 /etc/cron.d/cobis-veza
```

### Log Rotation

```bash
sudo tee /etc/logrotate.d/cobis-veza > /dev/null << 'EOF'
/opt/cobis-veza/logs/*.log {
    daily
    rotate 30
    compress
    delaycompress
    missingok
    notifempty
    create 640 cobis-veza cobis-veza
}
EOF
```

## Security Considerations

- **Credential storage**: The `.env` file contains sensitive credentials and must have `chmod 600` permissions owned by the service account
- **SMB credentials**: If mounting via `/etc/fstab`, store credentials in a separate file with `chmod 600` (see Mounting section above)
- **API key rotation**: Rotate the Veza API key periodically and update `.env` accordingly
- **Network**: Ensure the script can only reach the Veza API and the SMB share — no unnecessary outbound access
- **SELinux / AppArmor**: Verify contexts are correct after installation on hardened systems

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `ModuleNotFoundError: xlrd` | Dependencies not installed | `source venv/bin/activate && pip install -r requirements.txt` |
| `Unsupported file extension` | File is not `.xls` or `.xlsx` | Confirm the file format; rename if needed |
| `No user records found` | Workbook layout doesn't match expected format | Check that row 12 has "User Name" in col A and row 13 has "Group:" in col A |
| `Veza push failed: 401` | Invalid or expired API key | Regenerate the Veza API key and update `.env` |
| `Veza push failed: 403` | API key lacks OAA permissions | Ensure the key has the OAA Integration role |
| `SMB connection refused` | Firewall or wrong port | Verify network path and port 445 is open |
| `smbprotocol not installed` | Missing optional dependency | `pip install smbprotocol` |
| `mount.cifs: permission denied` | Bad SMB credentials or share perms | Verify credentials and share ACLs |

### Enabling Debug Logging

```bash
python3 cobis.py --env-file .env --log-level DEBUG --dry-run
```

## Changelog

- **v1.0** — Initial release: Excel (.xls/.xlsx) parsing, SMB support, user → group membership model
