import sys
from database import init_db
from gui import VoucherApp
from config import logger

if __name__ == "__main__":
    try:
        init_db()
        app = VoucherApp()
        app.mainloop()
    except Exception as e:
        logger.exception("App crashed")
        import tkinter.messagebox
        tkinter.messagebox.showerror("Crash", str(e))
