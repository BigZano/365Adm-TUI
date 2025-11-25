"""
Microsoft 365 Admin TUI
A Textual-based Terminal User Interface for managing Microsoft 365 via PowerShell scripts.
"""

from textual.app import App, ComposeResult
from textual.containers import Container, Horizontal, Vertical, ScrollableContainer, VerticalScroll
from textual.widgets import Header, Footer, Button, Static, Label, Input, LoadingIndicator, Rule
from textual.screen import Screen, ModalScreen
from textual.binding import Binding
from textual import work
from pathlib import Path
import asyncio
import os
import sys

# Add lib directory to path
sys.path.insert(0, str(Path(__file__).parent / "lib"))

from lib.logger import get_logger, setup_logging
from lib.config import Config

# Setup logging
config = Config()
log_file = setup_logging(config.logs_dir)
logger = get_logger(__name__)

# ============================================================================
# INPUT SCREENS FOR EACH SCRIPT
# ============================================================================

class CreateUserScreen(ModalScreen):
    """Screen for creating a new user."""
    
    BINDINGS = [("escape", "app.pop_screen", "Cancel")]
    
    def compose(self) -> ComposeResult:
        yield Container(
            Label("➕ Create New Microsoft 365 User", classes="dialog-title"),
            Rule(line_style="heavy"),
            
            Label("Display Name:", classes="field-label"),
            Label("Full name of the user (e.g., John Doe)", classes="field-hint"),
            Input(placeholder="John Doe", id="display_name"),
            
            Label("User Principal Name:", classes="field-label"),
            Label("Email address / login (e.g., john.doe@company.com)", classes="field-hint"),
            Input(placeholder="john.doe@company.com", id="upn"),
            
            Label("Usage Location:", classes="field-label"),
            Label("2-letter country code (e.g., US, GB, CA)", classes="field-hint"),
            Input(placeholder="US", id="location", max_length=2),
            
            Label("Password:", classes="field-label"),
            Label("Minimum 8 characters with upper, lower, and numbers", classes="field-hint"),
            Input(placeholder="Temp@Pass123", password=True, id="password"),
            
            Label("License Selection:", classes="field-label"),
            Label("Enter license index (1, 2, 3...) or 0 to skip licensing", classes="field-hint"),
            Label("💡 Tip: Run 'List Licenses' first to see available licenses and their index numbers", classes="field-hint"),
            Input(placeholder="0", id="license_index", value="0"),
            
            Horizontal(
                Button("Create User", variant="success", id="submit"),
                Button("List Licenses First", variant="default", id="list_licenses"),
                Button("Cancel", variant="default", id="cancel"),
                classes="button-row"
            ),
            id="dialog"
        )
    
    async def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "cancel":
            self.app.pop_screen()
        elif event.button.id == "list_licenses":
            # Run the script in list-only mode
            self.app.notify("📋 Listing available licenses...", severity="information")
            self.dismiss({"action": "list_licenses"})
        elif event.button.id == "submit":
            # Collect form data
            data = {
                "action": "create_user",
                "display_name": self.query_one("#display_name", Input).value.strip(),
                "upn": self.query_one("#upn", Input).value.strip(),
                "location": self.query_one("#location", Input).value.strip().upper(),
                "password": self.query_one("#password", Input).value,
                "license_index": self.query_one("#license_index", Input).value.strip(),
            }
            
            # Validate inputs
            if not all([data["display_name"], data["upn"], data["location"], data["password"]]):
                self.app.notify("⚠️ Please fill in all required fields", severity="error")
                return
            
            if len(data["location"]) != 2:
                self.app.notify("⚠️ Location must be 2-letter country code", severity="error")
                return
            
            if len(data["password"]) < 8:
                self.app.notify("⚠️ Password must be at least 8 characters", severity="error")
                return
            
            logger.info(f"Creating user: {data['upn']}")
            self.dismiss(data)


class DelegateAccessScreen(ModalScreen):
    """Screen for delegate access audit."""
    
    BINDINGS = [("escape", "app.pop_screen", "Cancel")]
    
    def compose(self) -> ComposeResult:
        yield Container(
            Label("🔍 Audit Delegate Access Permissions", classes="dialog-title"),
            Rule(line_style="heavy"),
            Label("Target User Email:", classes="field-label"),
            Label("Email of user to check permissions for", classes="field-hint"),
            Input(placeholder="user@company.com", id="target_email"),
            
            Horizontal(
                Button("Run Audit", variant="success", id="submit"),
                Button("Cancel", variant="default", id="cancel"),
                classes="button-row"
            ),
            id="dialog"
        )
    
    async def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "cancel":
            self.app.pop_screen()
        elif event.button.id == "submit":
            data = {
                "target_email": self.query_one("#target_email", Input).value.strip(),
            }
            
            if not data["target_email"]:
                self.app.notify("⚠️ Please enter target user email", severity="error")
                return
            
            logger.info(f"Running delegate access audit for: {data['target_email']}")
            self.dismiss(data)


class MailboxExportScreen(ModalScreen):
    """Screen for mailbox export."""
    
    BINDINGS = [("escape", "app.pop_screen", "Cancel")]
    
    def compose(self) -> ComposeResult:
        yield Container(
            Label("📊 Export Mailbox Report", classes="dialog-title"),
            Rule(line_style="heavy"),
            Label("Mailbox Type:", classes="field-label"),
            Label("Options: All, UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox", classes="field-hint"),
            Input(placeholder="All", id="mailbox_type", value="All"),
            
            Horizontal(
                Button("Export Report", variant="success", id="submit"),
                Button("Cancel", variant="default", id="cancel"),
                classes="button-row"
            ),
            id="dialog"
        )
    
    async def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "cancel":
            self.app.pop_screen()
        elif event.button.id == "submit":
            data = {
                "mailbox_type": self.query_one("#mailbox_type", Input).value.strip() or "All",
            }
            
            valid_types = ["All", "UserMailbox", "SharedMailbox", "RoomMailbox", "EquipmentMailbox"]
            if data["mailbox_type"] not in valid_types:
                self.app.notify(f"⚠️ Invalid mailbox type. Use: {', '.join(valid_types)}", severity="error")
                return
            
            logger.info(f"Exporting mailbox report for type: {data['mailbox_type']}")
            self.dismiss(data)


# ============================================================================
# MAIN APPLICATION
# ============================================================================

class M365AdminApp(App):
    """A professional Textual app for Microsoft 365 administration."""
    
    # Load external stylesheets
    CSS_PATH = [
        Path(__file__).parent / "themes" / "app.tcss",
        Path(__file__).parent / "themes" / "dialogs.tcss",
    ]
    
    BINDINGS = [
        Binding("q", "quit", "Quit", show=True),
        Binding("d", "toggle_dark", "Toggle Theme", show=True),
        Binding("c", "clear_output", "Clear Output", show=True),
        Binding("s", "toggle_switches", "Switches", show=True),
        Binding("l", "list_licenses", "List Licenses", show=False),
        Binding("up", "focus_previous", "Up", show=False),
        Binding("down", "focus_next", "Down", show=False),
    ]
    
    def __init__(self):
        super().__init__()
        self.config = Config()
        self.switches_visible = False
        logger.info("M365 Admin TUI started")
        logger.info(f"Log file: {log_file}")
        logger.info(f"Output directory: {self.config.output_dir}")
    
    def compose(self) -> ComposeResult:
        """Create child widgets for the app."""
        yield Header()
        
        with Vertical(id="main-container"):
            with Container(id="header-section"):
                yield Label("🔷 Microsoft 365 Admin TUI 🔷", classes="app-title")
                yield Label("Secure PowerShell Script Manager with OAuth2 & MFA Support", classes="app-subtitle")
                yield Rule(line_style="heavy")
                yield Label(f"📁 Reports will be saved to: {self.config.output_dir}", classes="info-text")
            
            with Container(id="menu-section"):
                yield Label("⚙️  Available Operations", classes="section-title")
                with Vertical(classes="menu-buttons"):
                    yield Button("👤 Create User (Microsoft Graph)", id="create_user", variant="primary")
                    yield Button("🔍 Audit Delegate Access", id="delegate_access", variant="default")
                    yield Button("📊 Export Mailbox Report", id="mailbox_export", variant="default")
                    yield Button("🔐 MFA Audit (All Users)", id="mfa_audit", variant="default")
                    yield Button("🔑 Authentication Method Report", id="auth_method", variant="default")
            
            with Container(id="switches-section", classes="hidden"):
                yield Label("🔧 Utility Switches", classes="section-title")
                yield Label("Press the key shown to execute the command", classes="switches-hint")
                with Vertical(classes="switches-menu"):
                    yield Label("[L] List Available Licenses", id="switch-list-licenses", classes="switch-item")
                    yield Label("[S] Close Switches Menu", id="switch-close", classes="switch-item-muted")
            
            with Container(id="output-section"):
                yield Label("📋 Script Output", id="output-title")
                with Container(id="output-container"):
                    with VerticalScroll(id="output-scroll"):
                        yield Static("Ready to execute commands...\nSelect an operation above to begin.", 
                                   id="output-content", classes="output-ready")
        
        yield Footer()
    
    def on_mount(self) -> None:
        """Called when app is mounted."""
        self.title = "M365 Admin TUI"
        self.sub_title = f"Log: {log_file.name}"
        self.theme = "textual-dark"  # Start with dark theme
        logger.info("Application mounted and ready")
    
    def action_toggle_dark(self) -> None:
        """Toggle dark mode."""
        # Toggle between dark and light themes
        if self.theme == "textual-dark":
            self.theme = "textual-light"
            logger.info("Theme changed to: light")
            self.notify("Theme: Light")
        else:
            self.theme = "textual-dark"
            logger.info("Theme changed to: dark")
            self.notify("Theme: Dark")
    
    def action_toggle_switches(self) -> None:
        """Toggle the switches menu visibility."""
        switches_section = self.query_one("#switches-section", Container)
        self.switches_visible = not self.switches_visible
        
        if self.switches_visible:
            switches_section.remove_class("hidden")
            self.notify("Switches menu opened - Press 'L' for licenses, 'S' to close", severity="information")
            logger.info("Switches menu opened")
        else:
            switches_section.add_class("hidden")
            self.notify("Switches menu closed")
            logger.info("Switches menu closed")
    
    def action_list_licenses(self) -> None:
        """Run the list licenses command."""
        # Only run if switches menu is visible
        if not self.switches_visible:
            return
        
        logger.info("List licenses command triggered from switches")
        self.notify("📋 Listing available licenses...", severity="information")
        self.run_list_licenses_script()
    
    @work(exclusive=True)
    async def run_list_licenses_script(self) -> None:
        """Execute the PowerShell script to list licenses."""
        output_panel = self.query_one("#output-content", Static)
        script_path = self.config.get_script_path("MgGraphUserCreation.ps1")
        display_name = "📋 List Available Licenses"
        
        logger.info("Starting license listing script")
        output_panel.update(f"🔄 Running {display_name}...\n\n⏳ Please wait, this may take a moment...\n\n🔐 You will be prompted to sign in with your admin credentials.\nMFA authentication is supported.")
        output_panel.remove_class("output-ready", "output-success", "output-error", "output-info")
        output_panel.set_class(True, "output-running")
        
        try:
            # Check if pwsh is available
            pwsh_check = await asyncio.create_subprocess_exec(
                "which", "pwsh",
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE
            )
            await pwsh_check.communicate()
            
            if pwsh_check.returncode != 0:
                raise Exception("PowerShell (pwsh) not found. Please install PowerShell Core.")
            
            # Run with -ListLicenses flag
            process = await asyncio.create_subprocess_exec(
                "pwsh",
                "-NoProfile",
                "-File",
                str(script_path),
                "-ListLicenses",
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE
            )
            
            stdout, stderr = await process.communicate()
            
            if process.returncode == 0:
                output_text = f"✅ {display_name} completed successfully!\n\n{'='*60}\n\n{stdout.decode()}"
                output_panel.update(output_text)
                output_panel.remove_class("output-ready", "output-running", "output-error", "output-info")
                output_panel.set_class(True, "output-success")
                self.notify(f"✅ Licenses listed!", severity="success")
                logger.info("License listing completed successfully")
            else:
                error_text = stderr.decode() if stderr else "Unknown error"
                output_text = f"❌ {display_name} failed!\n\n{'='*60}\n\nError:\n{error_text}\n\n{'='*60}\n\nStdout:\n{stdout.decode()}"
                output_panel.update(output_text)
                output_panel.remove_class("output-ready", "output-running", "output-success", "output-info")
                output_panel.set_class(True, "output-error")
                self.notify(f"❌ Failed - check output", severity="error")
                logger.error(f"License listing failed: {error_text}")
                
        except Exception as e:
            error_msg = str(e)
            output_panel.update(f"❌ Error executing script: {error_msg}\n\nPlease check:\n• PowerShell Core (pwsh) is installed\n• Required PowerShell modules are installed\n• Network connectivity is working")
            output_panel.remove_class("output-ready", "output-running", "output-success", "output-info")
            output_panel.set_class(True, "output-error")
            self.notify(f"❌ Error: {error_msg}", severity="error")
            logger.error(f"Exception running license list: {error_msg}")
    
    def action_clear_output(self) -> None:
        """Clear the output panel."""
        output = self.query_one("#output-content", Static)
        output.update("Output cleared.\nReady for next command...")
        # Clear all status classes first
        output.remove_class("output-running", "output-success", "output-error", "output-info")
        output.set_class(True, "output-ready")
        logger.info("Output cleared by user")
    
    def on_button_pressed(self, event: Button.Pressed) -> None:
        """Handle button press events."""
        button_id = event.button.id
        logger.info(f"Button pressed: {button_id}")
        
        # MFA Audit and Auth Method don't need input - run directly
        if button_id == "mfa_audit":
            self.run_script_no_input("mfa_audit.ps1", "🔐 MFA Audit")
        elif button_id == "auth_method":
            self.run_script_no_input("MFA_AuthMethod.ps1", "🔑 Authentication Method Report")
        elif button_id in ["create_user", "delegate_access", "mailbox_export"]:
            # Scripts that need input - show appropriate screen
            self.show_input_screen(button_id)
    
    @work(exclusive=True)
    async def show_input_screen(self, button_id: str) -> None:
        """Show input screen and run script with parameters."""
        screen_map = {
            "create_user": CreateUserScreen(),
            "delegate_access": DelegateAccessScreen(),
            "mailbox_export": MailboxExportScreen(),
        }
        
        if button_id in screen_map:
            result = await self.push_screen_wait(screen_map[button_id])
            if result:
                await self.run_script_with_params(button_id, result)
    
    @work(exclusive=True)
    async def run_script_no_input(self, script_name: str, display_name: str) -> None:
        """Execute a PowerShell script that doesn't require input."""
        output_panel = self.query_one("#output-content", Static)
        script_path = self.config.get_script_path(script_name)
        
        logger.info(f"Starting script: {script_name}")
        output_panel.update(f"🔄 Running {display_name}...\n\n⏳ Please wait, this may take a moment...\n\n🔐 You will be prompted to sign in with your admin credentials.\nMFA authentication is supported.")
        output_panel.remove_class("output-ready", "output-success", "output-error", "output-info")
        output_panel.set_class(True, "output-running")
        self.notify(f"Starting {display_name}...", severity="information")
        
        try:
            # Check if pwsh is available
            pwsh_check = await asyncio.create_subprocess_exec(
                "which", "pwsh",
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE
            )
            await pwsh_check.communicate()
            
            if pwsh_check.returncode != 0:
                raise Exception("PowerShell (pwsh) not found. Please install PowerShell Core.")
            
            process = await asyncio.create_subprocess_exec(
                "pwsh",
                "-NoProfile",
                "-File",
                str(script_path),
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE
            )
            
            stdout, stderr = await process.communicate()
            
            if process.returncode == 0:
                output_text = f"✅ {display_name} completed successfully!\n\n{'='*60}\n\n{stdout.decode()}"
                output_panel.update(output_text)
                output_panel.remove_class("output-ready", "output-running", "output-error", "output-info")
                output_panel.set_class(True, "output-success")
                self.notify(f"✅ {display_name} completed!", severity="success")
                logger.info(f"Script completed successfully: {script_name}")
            else:
                error_text = stderr.decode() if stderr else "Unknown error"
                output_text = f"❌ {display_name} failed!\n\n{'='*60}\n\nError:\n{error_text}\n\n{'='*60}\n\nStdout:\n{stdout.decode()}"
                output_panel.update(output_text)
                output_panel.remove_class("output-ready", "output-running", "output-success", "output-info")
                output_panel.set_class(True, "output-error")
                self.notify(f"❌ {display_name} failed - check output", severity="error")
                logger.error(f"Script failed: {script_name} - {error_text}")
                
        except Exception as e:
            error_msg = str(e)
            output_panel.update(f"❌ Error executing script: {error_msg}\n\nPlease check:\n• PowerShell Core (pwsh) is installed\n• Required PowerShell modules are installed\n• Network connectivity is working")
            output_panel.remove_class("output-ready", "output-running", "output-success", "output-info")
            output_panel.set_class(True, "output-error")
            self.notify(f"❌ Error: {error_msg}", severity="error")
            logger.error(f"Exception running script {script_name}: {error_msg}")
    
    @work(exclusive=True)
    async def run_script_with_params(self, script_type: str, params: dict) -> None:
        """Execute a PowerShell script with parameters."""
        output_panel = self.query_one("#output-content", Static)
        
        # Handle special case: List licenses only
        if script_type == "create_user" and params.get("action") == "list_licenses":
            script_path = self.config.get_script_path("MgGraphUserCreation.ps1")
            display_name = "📋 List Available Licenses"
            
            logger.info("Listing available licenses")
            output_panel.update(f"🔄 Running {display_name}...\n\n⏳ Please wait, this may take a moment...\n\n🔐 You will be prompted to sign in with your admin credentials.\nMFA authentication is supported.")
            output_panel.remove_class("output-ready", "output-success", "output-error", "output-info")
            output_panel.set_class(True, "output-running")
            self.notify(f"Starting {display_name}...", severity="information")
            
            try:
                pwsh_check = await asyncio.create_subprocess_exec(
                    "which", "pwsh",
                    stdout=asyncio.subprocess.PIPE,
                    stderr=asyncio.subprocess.PIPE
                )
                await pwsh_check.communicate()
                
                if pwsh_check.returncode != 0:
                    raise Exception("PowerShell (pwsh) not found. Please install PowerShell Core.")
                
                # Run with -ListLicenses flag
                cmd = ["pwsh", "-NoProfile", "-File", str(script_path), "-ListLicenses"]
                
                process = await asyncio.create_subprocess_exec(
                    *cmd,
                    stdout=asyncio.subprocess.PIPE,
                    stderr=asyncio.subprocess.PIPE
                )
                
                stdout, stderr = await process.communicate()
                
                if process.returncode == 0:
                    output_text = f"✅ {display_name} completed successfully!\n\n{'='*60}\n\n{stdout.decode()}"
                    output_panel.update(output_text)
                    output_panel.remove_class("output-ready", "output-running", "output-error", "output-info")
                    output_panel.set_class(True, "output-success")
                    self.notify(f"✅ {display_name} completed!", severity="success")
                    logger.info("License listing completed successfully")
                else:
                    error_text = stderr.decode() if stderr else "Unknown error"
                    output_text = f"❌ {display_name} failed!\n\n{'='*60}\n\nError:\n{error_text}\n\n{'='*60}\n\nStdout:\n{stdout.decode()}"
                    output_panel.update(output_text)
                    output_panel.remove_class("output-ready", "output-running", "output-success", "output-info")
                    output_panel.set_class(True, "output-error")
                    self.notify(f"❌ {display_name} failed - check output", severity="error")
                    logger.error(f"License listing failed: {error_text}")
                    
            except Exception as e:
                error_msg = str(e)
                output_panel.update(f"❌ Error executing script: {error_msg}\n\nPlease check:\n• PowerShell Core (pwsh) is installed\n• Required PowerShell modules are installed\n• Network connectivity is working")
                output_panel.remove_class("output-ready", "output-running", "output-success", "output-info")
                output_panel.set_class(True, "output-error")
                self.notify(f"❌ Error: {error_msg}", severity="error")
                logger.error(f"Exception running license list: {error_msg}")
            
            return
        
        # Map script types to files and parameter construction
        script_config = {
            "create_user": {
                "file": "MgGraphUserCreation.ps1",
                "display_name": "👤 Create User",
                "args": [
                    "-DisplayName", params["display_name"],
                    "-UserPrincipalName", params["upn"],
                    "-UsageLocation", params["location"],
                    "-Password", params["password"],
                    "-LicenseIndex", params["license_index"],
                ]
            },
            "delegate_access": {
                "file": "Loop for Delegate access.ps1",
                "display_name": "🔍 Delegate Access Audit",
                "args": [
                    "-TargetUserEmail", params["target_email"],
                ]
            },
            "mailbox_export": {
                "file": "Mailbox export.ps1",
                "display_name": "📊 Mailbox Export",
                "args": [
                    "-MailboxType", params["mailbox_type"],
                ]
            },
        }
        
        if script_type not in script_config:
            self.notify("Unknown script type", severity="error")
            logger.error(f"Unknown script type: {script_type}")
            return
        
        config = script_config[script_type]
        script_path = self.config.get_script_path(config["file"])
        display_name = config["display_name"]
        
        logger.info(f"Starting script: {config['file']} with params: {params}")
        output_panel.update(f"🔄 Running {display_name}...\n\n⏳ Please wait, this may take a moment...\n\n🔐 You will be prompted to sign in with your admin credentials.\nMFA authentication is supported.")
        output_panel.remove_class("output-ready", "output-success", "output-error", "output-info")
        output_panel.set_class(True, "output-running")
        self.notify(f"Starting {display_name}...", severity="information")
        
        try:
            # Check if pwsh is available
            pwsh_check = await asyncio.create_subprocess_exec(
                "which", "pwsh",
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE
            )
            await pwsh_check.communicate()
            
            if pwsh_check.returncode != 0:
                raise Exception("PowerShell (pwsh) not found. Please install PowerShell Core.")
            
            # Build command
            cmd = ["pwsh", "-NoProfile", "-File", str(script_path)] + config["args"]
            
            process = await asyncio.create_subprocess_exec(
                *cmd,
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE
            )
            
            stdout, stderr = await process.communicate()
            
            if process.returncode == 0:
                output_text = f"✅ {display_name} completed successfully!\n\n{'='*60}\n\n{stdout.decode()}"
                output_panel.update(output_text)
                output_panel.remove_class("output-ready", "output-running", "output-error", "output-info")
                output_panel.set_class(True, "output-success")
                self.notify(f"✅ {display_name} completed!", severity="success")
                logger.info(f"Script completed successfully: {config['file']}")
            else:
                error_text = stderr.decode() if stderr else "Unknown error"
                output_text = f"❌ {display_name} failed!\n\n{'='*60}\n\nError:\n{error_text}\n\n{'='*60}\n\nStdout:\n{stdout.decode()}"
                output_panel.update(output_text)
                output_panel.remove_class("output-ready", "output-running", "output-success", "output-info")
                output_panel.set_class(True, "output-error")
                self.notify(f"❌ {display_name} failed - check output", severity="error")
                logger.error(f"Script failed: {config['file']} - {error_text}")
                
        except Exception as e:
            error_msg = str(e)
            output_panel.update(f"❌ Error executing script: {error_msg}\n\nPlease check:\n• PowerShell Core (pwsh) is installed\n• Required PowerShell modules are installed\n• Network connectivity is working")
            output_panel.remove_class("output-ready", "output-running", "output-success", "output-info")
            output_panel.set_class(True, "output-error")
            self.notify(f"❌ Error: {error_msg}", severity="error")
            logger.error(f"Exception running script {config['file']}: {error_msg}")


def main():
    """Main entry point."""
    app = M365AdminApp()
    try:
        app.run()
    except Exception as e:
        logger.error(f"Application error: {e}")
        raise


if __name__ == "__main__":
    main()
