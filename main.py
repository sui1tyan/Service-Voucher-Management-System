import sys
import os
from database import init_db
from gui import VoucherApp
from config import logger

if __name__ == "__main__":
    try:
        # Initialize Database
        init_db()
        
        # Start GUI
        app = VoucherApp()
        app.mainloop()
        
    except Exception as e:
        logger.exception("Application crashed")
        # Ensure user sees error if GUI fails to start
        try:
            import tkinter.messagebox
            tkinter.messagebox.showerror("Fatal Error", f"App crashed:\n{e}")
        except:
            print(f"CRITICAL: {e}")
