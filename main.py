import tkinter as tk
from login_gui import LoginGUI

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    mainwin = tk.Toplevel(root)
    app = LoginGUI(mainwin)
    app.pack(fill="both", expand=True)
    mainwin.protocol("WM_DELETE_WINDOW", root.quit)  # <-- This ensures full exit!
    mainwin.mainloop()