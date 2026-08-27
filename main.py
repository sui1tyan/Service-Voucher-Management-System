"""SVMS application entry point."""

from tkinter import messagebox

from config import APP_NAME, logger
from database import init_db
from gui import VoucherApp

def main() -> None:
    try:
        init_db()
        app = VoucherApp()
        if app.ready:
            app.mainloop()
    except Exception as exc:
        logger.exception("Application crashed")
        messagebox.showerror(APP_NAME, f"The application could not continue:\n\n{exc}")


if __name__ == "__main__":
    main()
