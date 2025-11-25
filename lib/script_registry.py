"""
Script Registry - Auto-discovers and manages PowerShell scripts

Script descriptions are automatically extracted from the first comment line in each script.
You can override any auto-generated description by adding an entry to SCRIPT_DESCRIPTIONS.
"""
from pathlib import Path
from typing import Dict, List, Optional
import re
from dataclasses import dataclass


# Custom description overrides (optional - auto-extracted from script comments if not specified)
SCRIPT_DESCRIPTIONS = {
    # "Loop for Delegate access": "Custom description here",
}


@dataclass
class ScriptParameter:
    """Represents a script parameter"""
    name: str
    prompt: str
    default: str = ""
    required: bool = True
    password: bool = False


@dataclass
class ScriptInfo:
    """Information about a PowerShell script"""
    name: str
    path: Path
    description: str
    parameters: List[ScriptParameter]
    has_switches: bool = False
    switch_description: str = ""


class ScriptRegistry:
    """Discovers and manages PowerShell scripts from the Scripts directory"""
    
    def __init__(self, scripts_dir: Path):
        self.scripts_dir = scripts_dir
        self.scripts: Dict[str, ScriptInfo] = {}
        self._discover_scripts()
    
    def _discover_scripts(self):
        """Find all .ps1 files in the scripts directory"""
        if not self.scripts_dir.exists():
            return
        
        for script_path in self.scripts_dir.glob("*.ps1"):
            script_info = self._parse_script(script_path)
            if script_info:
                self.scripts[script_path.stem] = script_info
    
    def _parse_script(self, script_path: Path) -> Optional[ScriptInfo]:
        """Extract script metadata (parameters and description)"""
        try:
            content = script_path.read_text(encoding='utf-8')
            
            # Get description: check override dict first, then extract from script
            script_name = script_path.stem
            if script_name in SCRIPT_DESCRIPTIONS:
                description = SCRIPT_DESCRIPTIONS[script_name]
            else:
                description = self._extract_first_comment(content, script_name)
            
            # Extract parameters from param block
            parameters = self._extract_parameters(content)
            
            # Check for utility switches
            has_switches = '-ListLicenses' in content or 'switch]' in content
            switch_desc = "Supports utility switches (press 'S' for options)" if has_switches else ""
            
            return ScriptInfo(
                name=script_name,
                path=script_path,
                description=description,
                parameters=parameters,
                has_switches=has_switches,
                switch_description=switch_desc
            )
        except Exception as e:
            print(f"Error parsing {script_path}: {e}")
            return None
    
    def _extract_first_comment(self, content: str, script_name: str) -> str:
        """Extract description from first meaningful comment in script"""
        lines = content.split('\n')
        
        for line in lines[:30]:  # Check first 30 lines
            stripped = line.strip()
            
            # Skip shebang and empty lines
            if not stripped or stripped.startswith('#!'):
                continue
            
            # Found a comment line
            if stripped.startswith('#'):
                desc = stripped.lstrip('#').strip()
                # Must be reasonable length and not look like a separator
                if len(desc) > 10 and not desc.replace('-', '').replace('=', '').strip() == '':
                    return desc
            
            # Stop at first non-comment code line
            if not stripped.startswith('#') and not stripped.startswith('<#'):
                break
        
        # Fallback to generic description
        return f"PowerShell script: {script_name}"
    
    def _extract_parameters(self, content: str) -> List[ScriptParameter]:
        """Extract parameter information from PowerShell param block"""
        parameters = []
        
        # Find param block: param( ... )
        param_match = re.search(r'param\s*\((.*?)\n\)', content, re.DOTALL | re.IGNORECASE)
        if not param_match:
            return parameters
        
        param_block = param_match.group(1)
        
        # Split by comma (before opening bracket)
        param_sections = re.split(r',\s*(?=\[)', param_block)
        
        for section in param_sections:
            section = section.strip()
            if not section:
                continue
            
            # Check [Parameter(Mandatory=$true/$false)]
            mandatory_match = re.search(r'\[Parameter\([^)]*Mandatory\s*=\s*\$(\w+)', section, re.IGNORECASE)
            is_required = mandatory_match and mandatory_match.group(1).lower() == 'true'
            
            # Extract [type]$name = "default"
            type_name_match = re.search(r'\[(\w+)\]\s*\$(\w+)(?:\s*=\s*"?([^",\n]*)"?)?', section)
            if not type_name_match:
                continue
            
            param_type, param_name, default_value = type_name_match.groups()
            default_value = default_value or ""
            
            # Skip switch parameters
            if param_type.lower() == 'switch':
                continue
            
            # If no Mandatory attribute, infer from default value
            if mandatory_match is None:
                is_required = not default_value or default_value.strip() == ''
            
            # Generate friendly prompt
            prompt = self._param_name_to_prompt(param_name)
            
            # Check if password field
            is_password = 'password' in param_name.lower()
            
            parameters.append(ScriptParameter(
                name=param_name,
                prompt=prompt,
                default=default_value.strip(),
                required=is_required,
                password=is_password
            ))
        
        return parameters
    
    def _param_name_to_prompt(self, param_name: str) -> str:
        """Convert parameter name to user-friendly prompt"""
        # Insert space before capitals
        spaced = re.sub(r'([A-Z])', r' \1', param_name).strip()
        
        # Known special cases
        replacements = {
            'Upn': 'UPN (User Principal Name)',
            'Sku': 'SKU',
            'Mfa': 'MFA',
            'Display Name': 'Display Name (Full Name)',
            'User Principal Name': 'User Principal Name (Email)',
            'Usage Location': 'Usage Location (2-letter country code)',
            'Password': 'Password (min 8 characters)',
            'License Index': 'License Index (0 to skip, or number from list)',
            'Target User Email': 'Target User Email',
            'Mailbox Type': 'Mailbox Type (All, UserMailbox, SharedMailbox, etc.)'
        }
        
        return replacements.get(spaced, spaced)
    
    def get_script_list(self) -> List[str]:
        """Get list of available script names"""
        return sorted(self.scripts.keys())
    
    def get_script_info(self, script_name: str) -> Optional[ScriptInfo]:
        """Get information about a specific script"""
        return self.scripts.get(script_name)
    
    def get_display_name(self, script_name: str) -> str:
        """Format script name in PascalCase for display"""
        # Convert filename to readable name
        name = script_name.replace('_', ' ').replace('-', ' ')
        # Remove common suffixes
        name = re.sub(r'\b(script|ps1)\b', '', name, flags=re.IGNORECASE)
        name = name.strip()
        
        # Insert space before capitals (for camelCase names like MgGraphUserCreation)
        name = re.sub(r'(?<!^)(?=[A-Z])', ' ', name)
        
        # Convert to PascalCase
        words = name.split()
        
        # Preserve acronyms
        acronyms = {'mfa': 'MFA', 'mg': 'Mg', 'upn': 'UPN', 'sku': 'SKU', 'graph': 'Graph'}
        
        pascal_words = []
        for word in words:
            word_lower = word.lower()
            if word_lower in acronyms:
                pascal_words.append(acronyms[word_lower])
            else:
                pascal_words.append(word.capitalize())
        
        return ''.join(pascal_words)
