"""
Configuration management for M365 Admin TUI.
Handles application settings and runtime configuration.
"""

import os
from pathlib import Path
from typing import Optional


class Config:
    """Application configuration."""
    
    def __init__(self):
        """Initialize configuration with default values."""
        self.base_dir = Path(__file__).parent.parent
        self.scripts_dir = self.base_dir / "Scripts"
        self.logs_dir = self.base_dir / "logs"
        self.output_dir = self._get_output_dir()
        
        # Logging settings
        self.log_level = os.getenv("M365_LOG_LEVEL", "INFO")
        
        # UI settings
        self.theme = os.getenv("M365_THEME", "dark")
        
    def _get_output_dir(self) -> Path:
        """Get the output directory for reports."""
        # Default to user's Documents/M365Reports
        if os.name == 'nt':  # Windows
            base = Path(os.environ.get('USERPROFILE', Path.home()))
        else:  # Linux/Mac
            base = Path.home()
        
        output_dir = base / "Documents" / "M365Reports"
        output_dir.mkdir(parents=True, exist_ok=True)
        return output_dir
    
    def get_script_path(self, script_name: str) -> Path:
        """Get full path to a PowerShell script."""
        return self.scripts_dir / script_name
    
    def get_log_file_pattern(self) -> str:
        """Get pattern for log files."""
        return str(self.logs_dir / "m365admin_*.log")
