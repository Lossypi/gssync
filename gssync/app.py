from pathlib import Path
from typing import Optional

import gspread
from textual import on, work
from textual.app import App, ComposeResult
from textual.binding import Binding
from textual.containers import Horizontal, Vertical
from textual.screen import Screen
from textual.widgets import Button, Footer, Header, Input, Label, ListItem, ListView, Static

from .config import CREDENTIALS_FILE, Config, load_config, save_config
from .sheets import get_client, list_sheet_names, open_spreadsheet
from .sync import pull_all, pull_sheet, push_all, push_sheet


class SetupScreen(Screen):
    CSS = """
    SetupScreen {
        align: center middle;
    }
    #setup-box {
        width: 60;
        height: auto;
        border: solid $primary;
        padding: 1 2;
    }
    #setup-box Label { margin-bottom: 1; }
    #setup-box Input { margin-bottom: 1; }
    """

    def compose(self) -> ComposeResult:
        with Vertical(id="setup-box"):
            yield Label("GSSync — First Run Setup", id="title")
            yield Label("Spreadsheet URL:")
            yield Input(placeholder="https://docs.google.com/spreadsheets/d/...", id="url")
            yield Label("Local file path:")
            yield Input(placeholder="C:\\data\\report.xlsx", id="filepath")
            yield Button("Continue →", variant="primary", id="continue")

    @on(Button.Pressed, "#continue")
    def on_continue(self) -> None:
        url = self.query_one("#url", Input).value.strip()
        filepath = self.query_one("#filepath", Input).value.strip()
        if not url or not filepath:
            return
        cfg = load_config()
        cfg.spreadsheet_url = url
        cfg.file_path = filepath
        save_config(cfg)
        self.app.switch_screen(MainScreen())


class MainScreen(Screen):
    BINDINGS = [
        Binding("r", "refresh", "Refresh"),
        Binding("right", "pull_sheet", "Pull →"),
        Binding("left", "push_sheet", "← Push"),
        Binding("p", "pull_all", "Pull All"),
        Binding("u", "push_all", "Push All"),
        Binding("f", "change_format", "Format"),
        Binding("e", "edit_paths", "Edit"),
        Binding("q", "quit", "Quit"),
        Binding("tab", "switch_panel", "Switch"),
    ]

    CSS = """
    #panels { height: 1fr; }
    #google-panel {
        width: 1fr;
        border: solid $success;
        padding: 0 1;
    }
    #local-panel {
        width: 1fr;
        border: solid $primary;
        padding: 0 1;
    }
    #google-panel.active-panel { border: double $success; }
    #local-panel.active-panel { border: double $primary; }
    .panel-title { text-align: center; text-style: bold; margin-bottom: 1; }
    #status { height: 1; padding: 0 1; background: $surface; }
    #status.error { color: $error; }
    """

    def __init__(self) -> None:
        super().__init__()
        self._config = load_config()
        self._spreadsheet: Optional[gspread.Spreadsheet] = None
        self._active_panel = "google"

    def compose(self) -> ComposeResult:
        yield Header()
        with Horizontal(id="panels"):
            with Vertical(id="google-panel"):
                yield Label("Google Sheets", classes="panel-title")
                yield ListView(id="google-list")
            with Vertical(id="local-panel"):
                yield Label("Local File", classes="panel-title")
                yield ListView(id="local-list")
        yield Static("Press R to load sheets", id="status")
        yield Footer()

    def on_mount(self) -> None:
        self.title = "GSSync"
        self.sub_title = self._config.spreadsheet_url.split("/d/")[1].split("/")[0][:20] if "/d/" in self._config.spreadsheet_url else ""
        self.query_one("#google-panel").add_class("active-panel")
        self._load_sheets()

    @work(exclusive=True, thread=True)
    def _load_sheets(self) -> None:
        self._set_status("Connecting to Google Sheets…")
        try:
            client = get_client()
            self._spreadsheet = open_spreadsheet(client, self._config.spreadsheet_url)
            google_names = list_sheet_names(self._spreadsheet)
        except Exception as e:
            self._set_status(f"Error: {e}", error=True)
            return

        from .storage import list_local_sheets
        local_names = list_local_sheets(Path(self._config.file_path), self._config.file_format)

        self.call_from_thread(self._populate_lists, google_names, local_names)
        self._set_status("Ready.")

    def _populate_lists(self, google_names: list, local_names: list) -> None:
        gl = self.query_one("#google-list", ListView)
        ll = self.query_one("#local-list", ListView)
        gl.clear()
        ll.clear()
        for name in google_names:
            gl.append(ListItem(Label(name)))
        for name in local_names:
            ll.append(ListItem(Label(name)))

    def _set_status(self, msg: str, error: bool = False) -> None:
        def _update():
            status = self.query_one("#status", Static)
            status.update(msg)
            status.set_class(error, "error")
        self.call_from_thread(_update)

    def _active_selected_name(self) -> Optional[str]:
        list_id = "#google-list" if self._active_panel == "google" else "#local-list"
        lv = self.query_one(list_id, ListView)
        if lv.highlighted_child is None:
            return None
        label = lv.highlighted_child.query_one(Label)
        return str(label.renderable)

    def action_switch_panel(self) -> None:
        self._active_panel = "local" if self._active_panel == "google" else "google"
        self.query_one("#google-panel").set_class(self._active_panel == "google", "active-panel")
        self.query_one("#local-panel").set_class(self._active_panel == "local", "active-panel")

    def action_refresh(self) -> None:
        self._load_sheets()

    @work(exclusive=True, thread=True)
    def action_pull_sheet(self) -> None:
        name = self._active_selected_name()
        if not name or not self._spreadsheet:
            return
        self._set_status(f"Pulling '{name}'…")
        try:
            pull_sheet(self._spreadsheet, name, Path(self._config.file_path), self._config.file_format)
            self._set_status(f"Pulled '{name}' successfully.")
        except Exception as e:
            self._set_status(f"Error: {e}", error=True)
        self._load_sheets()

    @work(exclusive=True, thread=True)
    def action_push_sheet(self) -> None:
        name = self._active_selected_name()
        if not name or not self._spreadsheet:
            return
        self._set_status(f"Pushing '{name}'…")
        try:
            push_sheet(self._spreadsheet, name, Path(self._config.file_path), self._config.file_format)
            self._set_status(f"Pushed '{name}' successfully.")
        except Exception as e:
            self._set_status(f"Error: {e}", error=True)
        self._load_sheets()

    @work(exclusive=True, thread=True)
    def action_pull_all(self) -> None:
        if not self._spreadsheet:
            return
        self._set_status("Pulling all sheets…")
        try:
            pull_all(self._spreadsheet, Path(self._config.file_path), self._config.file_format)
            self._set_status("Pulled all sheets successfully.")
        except Exception as e:
            self._set_status(f"Error: {e}", error=True)
        self._load_sheets()

    @work(exclusive=True, thread=True)
    def action_push_all(self) -> None:
        if not self._spreadsheet:
            return
        self._set_status("Pushing all sheets…")
        try:
            push_all(self._spreadsheet, Path(self._config.file_path), self._config.file_format)
            self._set_status("Pushed all sheets successfully.")
        except Exception as e:
            self._set_status(f"Error: {e}", error=True)

    def action_change_format(self) -> None:
        formats = ["xlsx", "json", "csv"]
        current = self._config.file_format
        next_fmt = formats[(formats.index(current) + 1) % len(formats)]
        self._config.file_format = next_fmt
        save_config(self._config)
        self._set_status(f"Format changed to {next_fmt}. Changes affect next pull/push.")

    def action_edit_paths(self) -> None:
        self.app.push_screen(SetupScreen())
