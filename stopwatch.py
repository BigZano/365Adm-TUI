from textual.app import App, ComposeResult
from textual.widgets import Footer, Header

class StopwatchApp(App):
    """Simple Stopwatch - trying out Textual."""
    
    BINDINGS = [("d", "toggle_dark", "Toggle dark mode")]

    def compose(self) -> ComposeResult:
        """create child widgets for the app."""
        yield Header()
        yield Footer()

    def action_toggle_dark(self) -> None:
        """Toggle dark mode."""
        self.theme = (
            "textual-dark" if self.theme == "textual-light" else "textual-light"
        )



if __name__ == "__main__":
    app = StopwatchApp()
    app.run()

    