#!/usr/bin/env bash
# install_cobis.sh — One-command installer for COBIS-Veza OAA integration
# Usage:
#   Interactive:     bash install_cobis.sh
#   Non-interactive: VEZA_URL=... VEZA_API_KEY=... bash install_cobis.sh --non-interactive
set -euo pipefail

# ── Defaults ───────────────────────────────────────────────────────────────
INSTALL_DIR="/opt/cobis-veza"
REPO_URL=""
BRANCH="main"
NON_INTERACTIVE=false
OVERWRITE_ENV=false
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# ── Colors ─────────────────────────────────────────────────────────────────
RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'; NC='\033[0m'
info()  { printf "${GREEN}[INFO]${NC}  %s\n" "$*"; }
warn()  { printf "${YELLOW}[WARN]${NC}  %s\n" "$*"; }
error() { printf "${RED}[ERROR]${NC} %s\n" "$*" >&2; }

# ── Parse flags ────────────────────────────────────────────────────────────
while [[ $# -gt 0 ]]; do
    case "$1" in
        --non-interactive) NON_INTERACTIVE=true; shift ;;
        --overwrite-env)   OVERWRITE_ENV=true; shift ;;
        --install-dir)     INSTALL_DIR="$2"; shift 2 ;;
        --repo-url)        REPO_URL="$2"; shift 2 ;;
        --branch)          BRANCH="$2"; shift 2 ;;
        -h|--help)
            echo "Usage: bash install_cobis.sh [OPTIONS]"
            echo ""
            echo "Options:"
            echo "  --non-interactive   Use env vars instead of prompts"
            echo "  --overwrite-env     Overwrite existing .env file"
            echo "  --install-dir DIR   Install directory (default: /opt/cobis-veza)"
            echo "  --repo-url URL      Git repository URL"
            echo "  --branch NAME       Git branch (default: main)"
            exit 0
            ;;
        *) error "Unknown option: $1"; exit 1 ;;
    esac
done

# ── Detect distro & package manager ───────────────────────────────────────
detect_pkg_manager() {
    if command -v dnf &>/dev/null; then
        PKG_MGR="dnf"
    elif command -v yum &>/dev/null; then
        PKG_MGR="yum"
    elif command -v apt-get &>/dev/null; then
        PKG_MGR="apt-get"
    else
        error "Unsupported package manager. Install git, python3, python3-pip, python3-venv manually."
        exit 1
    fi
}

install_packages() {
    local pkgs=("git" "curl" "python3" "python3-pip")
    case "$PKG_MGR" in
        dnf|yum)
            pkgs+=("python3-devel")
            info "Installing packages via $PKG_MGR: ${pkgs[*]}"
            sudo "$PKG_MGR" install -y "${pkgs[@]}"
            ;;
        apt-get)
            pkgs+=("python3-venv" "python3-dev")
            info "Installing packages via apt-get: ${pkgs[*]}"
            sudo apt-get update -qq
            sudo apt-get install -y "${pkgs[@]}"
            ;;
    esac

    # cifs-utils for mounting SMB shares (optional)
    if ! command -v mount.cifs &>/dev/null; then
        info "Installing cifs-utils for SMB share mounting..."
        sudo "$PKG_MGR" install -y cifs-utils 2>/dev/null || warn "cifs-utils install failed — SMB mount optional"
    fi
}

# ── Check Python version ──────────────────────────────────────────────────
check_python() {
    local py_version
    py_version=$(python3 -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')" 2>/dev/null || echo "0.0")
    local major minor
    major=$(echo "$py_version" | cut -d. -f1)
    minor=$(echo "$py_version" | cut -d. -f2)
    if [[ "$major" -lt 3 ]] || { [[ "$major" -eq 3 ]] && [[ "$minor" -lt 8 ]]; }; then
        error "Python >= 3.8 is required. Found: python3 $py_version"
        exit 1
    fi
    info "Python version: $py_version ✓"
}

# ── Prompt helper ──────────────────────────────────────────────────────────
prompt_value() {
    local var_name="$1" prompt_text="$2" default="${3:-}" secret="${4:-false}"
    local current_val="${!var_name:-$default}"

    if $NON_INTERACTIVE; then
        if [[ -z "$current_val" ]]; then
            error "Required: $var_name (set via environment variable)"
            exit 1
        fi
        printf -v "$var_name" '%s' "$current_val"
        return
    fi

    if [[ "$secret" == "true" ]]; then
        read -rsp "$prompt_text [hidden]: " "$var_name"
        echo
    else
        if [[ -n "$default" ]]; then
            read -rp "$prompt_text [$default]: " "$var_name"
            [[ -z "${!var_name}" ]] && printf -v "$var_name" '%s' "$default"
        else
            read -rp "$prompt_text: " "$var_name"
        fi
    fi
}

# ── Main ───────────────────────────────────────────────────────────────────
main() {
    echo "============================================================"
    echo "  COBIS → Veza OAA Integration Installer"
    echo "============================================================"
    echo

    detect_pkg_manager
    install_packages
    check_python

    # ── Create directory layout ────────────────────────────────────────────
    info "Creating directory layout at $INSTALL_DIR"
    sudo mkdir -p "$INSTALL_DIR/scripts" "$INSTALL_DIR/logs"

    # ── Copy or clone scripts ──────────────────────────────────────────────
    if [[ -n "$REPO_URL" ]]; then
        info "Cloning repository $REPO_URL (branch: $BRANCH)"
        if [[ -d "$INSTALL_DIR/scripts/.git" ]]; then
            cd "$INSTALL_DIR/scripts"
            sudo git fetch origin
            sudo git checkout "$BRANCH"
            sudo git pull origin "$BRANCH"
        else
            sudo git clone --branch "$BRANCH" "$REPO_URL" "$INSTALL_DIR/scripts"
        fi
    else
        # Copy from the directory where the installer lives
        info "Copying scripts from $SCRIPT_DIR"
        sudo cp -f "$SCRIPT_DIR/cobis.py" "$INSTALL_DIR/scripts/"
        sudo cp -f "$SCRIPT_DIR/requirements.txt" "$INSTALL_DIR/scripts/"
    fi

    # ── Python virtual environment ─────────────────────────────────────────
    info "Creating Python virtual environment"
    sudo python3 -m venv "$INSTALL_DIR/scripts/venv"
    sudo "$INSTALL_DIR/scripts/venv/bin/pip" install --upgrade pip -q
    sudo "$INSTALL_DIR/scripts/venv/bin/pip" install -r "$INSTALL_DIR/scripts/requirements.txt" -q
    info "Dependencies installed"

    # ── Collect credentials ────────────────────────────────────────────────
    ENV_FILE="$INSTALL_DIR/scripts/.env"

    if [[ -f "$ENV_FILE" ]] && ! $OVERWRITE_ENV; then
        warn ".env already exists at $ENV_FILE — skipping credential setup"
        warn "Use --overwrite-env to regenerate"
    else
        info "Configuring credentials..."
        echo

        prompt_value VEZA_URL       "Veza tenant URL (e.g. https://yourcompany.vezacloud.com)" ""
        prompt_value VEZA_API_KEY   "Veza API key" "" true

        echo
        info "Excel file source — choose ONE of the following:"
        echo "  1) Local / mounted path  (e.g. /mnt/cobis/report.xls)"
        echo "  2) Direct SMB access     (server/share/path + credentials)"
        echo

        if $NON_INTERACTIVE; then
            SOURCE_MODE="auto"
        else
            read -rp "Select [1/2]: " SOURCE_MODE
        fi

        COBIS_XLSX_PATH="${COBIS_XLSX_PATH:-}"
        COBIS_SMB_SERVER="${COBIS_SMB_SERVER:-}"
        COBIS_SMB_SHARE="${COBIS_SMB_SHARE:-}"
        COBIS_SMB_PATH="${COBIS_SMB_PATH:-}"
        COBIS_SMB_USERNAME="${COBIS_SMB_USERNAME:-}"
        COBIS_SMB_PASSWORD="${COBIS_SMB_PASSWORD:-}"

        if [[ "$SOURCE_MODE" == "1" ]] || { [[ "$SOURCE_MODE" == "auto" ]] && [[ -n "$COBIS_XLSX_PATH" ]]; }; then
            prompt_value COBIS_XLSX_PATH "Path to COBIS Excel file" "$COBIS_XLSX_PATH"
        else
            prompt_value COBIS_SMB_SERVER   "SMB server hostname/IP" "$COBIS_SMB_SERVER"
            prompt_value COBIS_SMB_SHARE    "SMB share name" "$COBIS_SMB_SHARE"
            prompt_value COBIS_SMB_PATH     "File path within the share" "$COBIS_SMB_PATH"
            prompt_value COBIS_SMB_USERNAME "SMB username" "$COBIS_SMB_USERNAME"
            prompt_value COBIS_SMB_PASSWORD "SMB password" "" true
        fi

        # Write .env
        sudo tee "$ENV_FILE" > /dev/null <<ENVEOF
# COBIS-Veza OAA Integration — Configuration
# Generated on $(date -u +"%Y-%m-%dT%H:%M:%SZ")

# Veza Configuration
VEZA_URL=${VEZA_URL}
VEZA_API_KEY=${VEZA_API_KEY}

# Excel File — Local / Mounted Path (option 1)
COBIS_XLSX_PATH=${COBIS_XLSX_PATH}

# Excel File — SMB Direct Access (option 2)
COBIS_SMB_SERVER=${COBIS_SMB_SERVER}
COBIS_SMB_SHARE=${COBIS_SMB_SHARE}
COBIS_SMB_PATH=${COBIS_SMB_PATH}
COBIS_SMB_USERNAME=${COBIS_SMB_USERNAME}
COBIS_SMB_PASSWORD=${COBIS_SMB_PASSWORD}

# OAA Provider Settings (optional)
# PROVIDER_NAME=COBIS
# DATASOURCE_NAME=COBIS
ENVEOF
        sudo chmod 600 "$ENV_FILE"
        info ".env written with 600 permissions"
    fi

    # ── Final summary ──────────────────────────────────────────────────────
    echo
    echo "============================================================"
    info "Installation complete!"
    echo "============================================================"
    echo
    echo "  Install directory:  $INSTALL_DIR"
    echo "  Scripts:            $INSTALL_DIR/scripts/"
    echo "  Logs:               $INSTALL_DIR/logs/"
    echo "  Config:             $ENV_FILE"
    echo "  Virtual env:        $INSTALL_DIR/scripts/venv/"
    echo
    echo "  Run a test (dry run):"
    echo "    cd $INSTALL_DIR/scripts"
    echo "    source venv/bin/activate"
    echo "    python3 cobis.py --env-file .env --dry-run"
    echo
    echo "  Schedule via cron (daily at 2 AM):"
    echo '    echo "0 2 * * * cobis-veza cd /opt/cobis-veza/scripts && venv/bin/python3 cobis.py --env-file .env >> /opt/cobis-veza/logs/cobis.log 2>&1" | sudo tee /etc/cron.d/cobis-veza'
    echo
}

main "$@"
