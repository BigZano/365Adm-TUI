"""
Microsoft 365 Admin TUI
A Textual-based Terminal User Interface for managing Microsoft 365 via PowerShell scripts.
"""

from textual.app import App, ComposeResult
from textual.containers import Container, Horizontal, Vertical, ScrollableContainer
from textual.widgets import Header, Footer, Button, Static, Label, Input
from textual.screen import Screen
from textual.binding import Binding
import asyncio
import os

# ============================================================================
# INPUT SCREENS FOR EACH SCRIPT
# ============================================================================

class CreateUserScreen(Screen):
    """Screen for creating a new user."""
    
    CSS = """
    CreateUserScreen {
        align: center middle;
    }
    
    #dialog {
        width: 60;
        height: auto;
        border: thick $background 80%;
        background: $surface;
        padding: 1 2;
    }
    
    Input {
        margin: 1 0;
    }
    
    .label {
        margin: 1 0 0 0;
        color: $text-muted;
    }
    """
    
    BINDINGS = [("escape", "app.pop_screen", "Back")]
    
    def compose(self) -> ComposeResult:
        yield Container(
            Label("Create New User", classes="title"),
            Label("Display Name:", classes="label"),
            Input(placeholder="John Doe", id="display_name"),
            Label("User Principal Name:", classes="label"),
            Input(placeholder="john.doe@company.com", id="upn"),
            Label("Usage Location:", classes="label"),
            Input(placeholder="US", id="location"),
            Label("Password:", classes="label"),
            Input(placeholder="Min 8 chars, upper, lower, number", password=True, id="password"),
            Label("License Index (0 to skip):", classes="label"),
            Input(placeholder="0", id="license_index", value="0"),
            Horizontal(
                Button("Create User", variant="primary", id="submit"),
                Button("Cancel", variant="default", id="cancel"),
            ),
            id="dialog"
        )
    
    async def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "cancel":
            self.app.pop_screen()
        elif event.button.id == "submit":
            # Collect form data
            data = {
                "display_name": self.query_one("#display_name", Input).value,
                "upn": self.query_one("#upn", Input).value,
                "location": self.query_one("#location", Input).value,
                "password": self.query_one("#password", Input).value,
                "license_index": self.query_one("#license_index", Input).value,
            }
            
            # Validate inputs
            if not all([data["display_name"], data["upn"], data["location"], data["password"]]):
                self.app.notify("Please fill in all required fields", severity="error")
                return
            
            self.dismiss(data)


class DelegateAccessScreen(Screen):
    """Screen for delegate access audit."""
    
    CSS = """
    DelegateAccessScreen {
        align: center middle;
    }
    
    #dialog {
        width: 60;
        height: auto;
        border: thick $background 80%;
        background: $surface;
        padding: 1 2;
    }
    
    Input {
        margin: 1 0;
    }
    
    .label {
        margin: 1 0 0 0;
        color: $text-muted;
    }
    """
    
    BINDINGS = [("escape", "app.pop_screen", "Back")]
    
    def compose(self) -> ComposeResult:
        yield Container(
            Label("Audit Delegate Access", classes="title"),
            Label("Admin Email (for Exchange Online):", classes="label"),
            Input(placeholder="admin@company.com", id="admin_email"),
            Label("Target User Email:", classes="label"),
            Input(placeholder="user@company.com", id="target_email"),
            Horizontal(
                Button("Run Audit", variant="primary", id="submit"),
                Button("Cancel", variant="default", id="cancel"),
            ),
            id="dialog"
        )
    
    async def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "cancel":
            self.app.pop_screen()
        elif event.button.id == "submit":
            data = {
                "admin_email": self.query_one("#admin_email", Input).value,
                "target_email": self.query_one("#target_email", Input).value,
            }
            
            if not all(data.values()):
                self.app.notify("Please fill in all fields", severity="error")
                return
            
            self.dismiss(data)


class MailboxExportScreen(Screen):
    """Screen for mailbox export."""
    
    CSS = """
    MailboxExportScreen {
        align: center middle;
    }
    
    #dialog {
        width: 60;
        height: auto;
        border: thick $background 80%;
        background: $surface;
        padding: 1 2;
    }
    
    Input {
        margin: 1 0;
    }
    
    .label {
        margin: 1 0 0 0;
        color: $text-muted;
    }
    """
    
    BINDINGS = [("escape", "app.pop_screen", "Back")]
    
    def compose(self) -> ComposeResult:
        yield Container(
            Label("Export Mailbox Report", classes="title"),
            Label("Admin Email (for Exchange Online):", classes="label"),
            Input(placeholder="admin@company.com", id="admin_email"),
            Label("Output Path (optional):", classes="label"),
            Input(placeholder="./MailboxReport.csv", id="output_path", value="./MailboxReport.csv"),
            Horizontal(
                Button("Export", variant="primary", id="submit"),
                Button("Cancel", variant="default", id="cancel"),
            ),
            id="dialog"
        )
    
    async def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "cancel":
            self.app.pop_screen()
        elif event.button.id == "submit":
            data = {
                "admin_email": self.query_one("#admin_email", Input).value,
                "output_path": self.query_one("#output_path", Input).value,
            }
            
            if not data["admin_email"]:
                self.app.notify("Please enter admin email", severity="error")
                return
            
            self.dismiss(data)


class AuthMethodScreen(Screen):
    """Screen for authentication method report."""
    
    CSS = """
    AuthMethodScreen {
        align: center middle;
    }
    
    #dialog {
        width: 60;
        height: auto;
        border: thick $background 80%;
        background: $surface;
        padding: 1 2;
    }
    
    Input {
        margin: 1 0;
    }
    
    .label {
        margin: 1 0 0 0;
        color: $text-muted;
    }
    """
    
    BINDINGS = [("escape", "app.pop_screen", "Back")]
    
    def compose(self) -> ComposeResult:
        yield Container(
            Label("Authentication Method Report", classes="title"),
            Label("Admin Email (for Exchange Online):", classes="label"),
            Input(placeholder="admin@company.com", id="admin_email"),
            Label("Output Path:", classes="label"),
            Input(placeholder="./UserAuthPolicies.csv", id="output_path", value="./UserAuthPolicies.csv"),
            Horizontal(
                Button("Generate Report", variant="primary", id="submit"),
                Button("Cancel", variant="default", id="cancel"),
            ),
            id="dialog"
        )
    
    async def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "cancel":
            self.app.pop_screen()
        elif event.button.id == "submit":
            data = {
                "admin_email": self.query_one("#admin_email", Input).value,
                "output_path": self.query_one("#output_path", Input).value,
            }
            
            if not data["admin_email"]:
                self.app.notify("Please enter admin email", severity="error")
                return
            
            self.dismiss(data)


# ============================================================================
# MAIN APPLICATION
# ============================================================================

class M365AdminApp(App):
    """A Textual app for Microsoft 365 administration."""
    
    CSS = """
    Screen {
        background: $surface;
    }
    
    Container {
        height: auto;
        margin: 1 2;
    }
    
    Button {
        margin: 1 2;
        min-width: 40;
    }
    
    .title {
        content-align: center middle;
        text-style: bold;
        color: $accent;
        margin: 1;
        text-style: bold underline;
    }
    
    #subtitle {
        margin: 0 2 1 2;
        color: $text-muted;
    }
    
    #output-label {
        margin: 2 2 0 2;
        text-style: bold;
        color: $text;
    }
    
    #output-panel {
        height: 20;
        border: solid $primary;
        margin: 0 2 1 2;
        padding: 1;
        background: $panel;
    }
    
    .success {
        color: $success;
    }
    
    .error {
        color: $error;
    }
    """
    
    BINDINGS = [
        Binding("q", "quit", "Quit", show=True),
        Binding("d", "toggle_dark", "Toggle Dark Mode", show=True),
    ]
    
    def compose(self) -> ComposeResult:
        """Create child widgets for the app."""
        yield Header()
        yield Container(
            Label("Microsoft 365 Admin TUI", classes="title"),
            Label("Select an operation:", id="subtitle"),
            Vertical(
                Button("👤 Create User (MgGraph)", id="create_user", variant="primary"),
                Button("📧 Audit Delegate Access", id="delegate_access", variant="default"),
                Button("📊 Export Mailbox Report", id="mailbox_export", variant="default"),
                Button("🔐 MFA Audit (No Input)", id="mfa_audit", variant="default"),
                Button("🔑 Auth Method Report", id="auth_method", variant="default"),
            ),
            Label("Output:", id="output-label"),
            ScrollableContainer(Static("Ready to process commands...", id="output-panel")),
        )
        yield Footer()
    
    def action_toggle_dark(self) -> None:
        """Toggle dark mode."""
        self.dark = not self.dark
    
    async def on_button_pressed(self, event: Button.Pressed) -> None:
        """Handle button press events."""
        button_id = event.button.id
        
        # MFA Audit doesn't need input - run directly
        if button_id == "mfa_audit":
            await self.run_script_no_input("mfa_audit.ps1")
            return
        
        # Other scripts need input - show appropriate screen
        screen_map = {
            "create_user": CreateUserScreen(),
            "delegate_access": DelegateAccessScreen(),
            "mailbox_export": MailboxExportScreen(),
            "auth_method": AuthMethodScreen(),
        }
        
        if button_id in screen_map:
            result = await self.push_screen_wait(screen_map[button_id])
            if result:
                await self.run_script_with_params(button_id, result)
    
    async def run_script_no_input(self, script_name: str) -> None:
        """Execute a PowerShell script that doesn't require input."""
        output_panel = self.query_one("#output-panel", Static)
        script_path = f"./Scripts/{script_name}"
        
        output_panel.update(f"🔄 Running {script_name}...\n")
        self.notify(f"Running {script_name}...", severity="information")
        
        try:
            process = await asyncio.create_subprocess_exec(
                "pwsh",
                "-NoProfile",
                "-File",
                script_path,
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE,
                cwd=os.path.dirname(os.path.abspath(__file__))
            )
            
            stdout, stderr = await process.communicate()
            
            if process.returncode == 0:
                output_text = f"✅ {script_name} completed successfully!\n\n{stdout.decode()}"
                output_panel.update(output_text)
                self.notify("Script completed successfully!", severity="success")
            else:
                output_text = f"❌ {script_name} failed!\n\nError:\n{stderr.decode()}"
                output_panel.update(output_text)
                self.notify("Script failed - check output", severity="error")
                
        except Exception as e:
            output_panel.update(f"❌ Error executing script: {str(e)}")
            self.notify(f"Error: {str(e)}", severity="error")
    
    async def run_script_with_params(self, script_type: str, params: dict) -> None:
        """Execute a PowerShell script with parameters."""
        output_panel = self.query_one("#output-panel", Static)
        
        # Map script types to files and parameter construction
        script_config = {
            "create_user": {
                "file": "MgGraphUserCreation.ps1",
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
                "args": [
                    "-AdminEmail", params["admin_email"],
                    "-TargetUserEmail", params["target_email"],
                ]
            },
            "mailbox_export": {
                "file": "Mailbox export.ps1",
                "args": [
                    "-AdminEmail", params["admin_email"],
                    "-OutputPath", params["output_path"],
                ]
            },
            "auth_method": {
                "file": "MFA_AuthMethod.ps1",
                "args": [
                    "-AdminEmail", params["admin_email"],
                    "-OutputPath", params["output_path"],
                ]
            },
        }
        
        if script_type not in script_config:
            self.notify("Unknown script type", severity="error")
            return
        
        config = script_config[script_type]
        script_path = f"./Scripts/{config['file']}"
        
        output_panel.update(f"🔄 Running {config['file']}...\n\nNote: Scripts need to be modified to accept parameters (see TEXTUAL_SETUP_GUIDE.md)")
        self.notify(f"Running {config['file']}...", severity="information")
        
        try:
            # Build command
            cmd = ["pwsh", "-NoProfile", "-File", script_path] + config["args"]
            
            process = await asyncio.create_subprocess_exec(
                *cmd,
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE,
                cwd=os.path.dirname(os.path.abspath(__file__))
            )
            
            stdout, stderr = await process.communicate()
            
            if process.returncode == 0:
                output_text = f"✅ {config['file']} completed successfully!\n\n{stdout.decode()}"
                output_panel.update(output_text)
                self.notify("Script completed successfully!", severity="success")
            else:
                output_text = f"❌ {config['file']} failed!\n\nError:\n{stderr.decode()}"
                output_panel.update(output_text)
                self.notify("Script failed - check output", severity="error")
                
        except Exception as e:
            output_panel.update(f"❌ Error executing script: {str(e)}")
            self.notify(f"Error: {str(e)}", severity="error")


if __name__ == "__main__":
    app = M365AdminApp()
    app.run()
