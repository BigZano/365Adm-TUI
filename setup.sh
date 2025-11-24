#!/bin/bash
# Setup script for M365 Admin TUI
# This script will install all necessary dependencies for Linux/macOS

set -e  # Exit on error

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

echo "Microsoft 365 Admin TUI Setup"
echo "===================================="
echo ""

# Check Python version
echo "Checking Python version..."
if ! command -v python3 &> /dev/null; then
    echo "Python 3 is not installed. Please install Python 3.12 or higher."
    exit 1
fi

PYTHON_VERSION=$(python3 --version 2>&1 | awk '{print $2}')
PYTHON_MAJOR=$(echo $PYTHON_VERSION | cut -d. -f1)
PYTHON_MINOR=$(echo $PYTHON_VERSION | cut -d. -f2)

if [ "$PYTHON_MAJOR" -lt 3 ] || ([ "$PYTHON_MAJOR" -eq 3 ] && [ "$PYTHON_MINOR" -lt 12 ]); then
    echo "Python 3.12 or higher is required. Current version: $PYTHON_VERSION"
    exit 1
fi

echo "Python $PYTHON_VERSION detected"
echo ""

# Check for PowerShell Core
echo "Checking PowerShell Core..."
if ! command -v pwsh &> /dev/null; then
    echo "PowerShell Core (pwsh) is not installed."
    echo ""
    echo "To install PowerShell Core:"
    echo ""
    
    if [[ "$OSTYPE" == "linux-gnu"* ]]; then
        echo "Ubuntu/Debian:"
        echo "  wget -q https://packages.microsoft.com/config/ubuntu/\$(lsb_release -rs)/packages-microsoft-prod.deb"
        echo "  sudo dpkg -i packages-microsoft-prod.deb"
        echo "  sudo apt-get update"
        echo "  sudo apt-get install -y powershell"
    elif [[ "$OSTYPE" == "darwin"* ]]; then
        echo "macOS:"
        echo "  brew install --cask powershell"
    fi
    echo ""
    read -p "Do you want to continue without PowerShell? (y/N) " -n 1 -r
    echo
    if [[ ! $REPLY =~ ^[Yy]$ ]]; then
        exit 1
    fi
else
    PWSH_VERSION=$(pwsh --version)
    echo "$PWSH_VERSION detected"
fi
echo ""

# Install Python dependencies
echo "Installing Python dependencies..."
if command -v uv &> /dev/null; then
    echo "Using uv for faster installation..."
    uv pip install -r requirements.txt
else
    python3 -m pip install -r requirements.txt
fi
echo "Python dependencies installed"
echo ""

# Check/Install PowerShell modules
if command -v pwsh &> /dev/null; then
    echo "Checking PowerShell modules..."
    
    # Create a temporary PowerShell script
    cat > /tmp/check_modules.ps1 << 'EOF'
$modules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Users", "ExchangeOnlineManagement")
$missing = @()

foreach ($module in $modules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        $missing += $module
    } else {
        Write-Host "$module is installed" -ForegroundColor Green
    }
}

if ($missing.Count -gt 0) {
    Write-Host "" 
    Write-Host "Missing PowerShell modules:" -ForegroundColor Yellow
    foreach ($module in $missing) {
        Write-Host "  - $module" -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "To install missing modules, run:" -ForegroundColor Cyan
    Write-Host "  pwsh" -ForegroundColor White
    foreach ($module in $missing) {
        if ($module -like "Microsoft.Graph.*") {
            Write-Host "  Install-Module Microsoft.Graph -Scope CurrentUser -Force" -ForegroundColor White
            break
        }
    }
    foreach ($module in $missing) {
        if ($module -eq "ExchangeOnlineManagement") {
            Write-Host "  Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force" -ForegroundColor White
        }
    }
} else {
    Write-Host "All PowerShell modules are installed" -ForegroundColor Green
}
EOF
    
    pwsh -NoProfile -File /tmp/check_modules.ps1
    rm /tmp/check_modules.ps1
    echo ""
fi

# Create necessary directories
echo "Creating directories..."
mkdir -p logs
mkdir -p ~/Documents/M365Reports
echo "Directories created"
echo ""

echo "===================================="
echo "Setup completed successfully!"
echo ""
echo "Next steps:"
echo " 1. If PowerShell modules are missing, install them"
echo " 2. Run the application: python3 main.py"
echo ""
echo "For more information, see README.md"
echo "===================================="
