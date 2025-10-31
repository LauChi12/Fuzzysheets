import gspread
import tkinter as tk
from tkinter import messagebox
from pathlib import Path

# Where the token will be stored on the user's computer
TOKEN_PATH = Path.home() / ".config" / "gspread" / "token.json"

def get_gspread_client(scopes=None):
    """
    Returns an authenticated gspread client.
    Opens a browser for first-time login if needed.
    Stores a token locally for subsequent runs.
    """

    scopes = scopes or [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    try:
        # gspread.oauth() will open the browser only if token.json doesn't exist
        gc = gspread.oauth(
            credentials_filename="client_secret.json",  
            authorized_user_filename=str(TOKEN_PATH),
            scopes=scopes
        )
        return gc

    except Exception as e:
        _show_auth_error(e)
        raise

def _show_auth_error(e):
    """Show a Tkinter popup with authentication error details."""
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Google Authentication Error",
            f"Could not authenticate with Google Sheets.\n\n"
            f"{type(e).__name__}: {e}\n\n"
            "Make sure you have internet access and complete the login in the browser."
        )
    finally:
        if "root" in locals():
            root.destroy()
