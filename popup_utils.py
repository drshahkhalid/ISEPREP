import tkinter as tk
from tkinter import ttk

"""
popup_utils.py
Unified Windows 11–style popup utilities (dynamic sizing & centering)
with accessible entry styling helpers.

Exports:
    custom_popup(parent, title, message, kind="info")
    custom_askyesno(parent, title, message)
    custom_dialog(parent, title, message, buttons, kind="question") -> returns button key
    show_toast(parent, message, kind="info", duration=2500)
"""

# ---------------------------------------------------------------------------
# Accent / Theme Configuration
# ---------------------------------------------------------------------------

_ACCENTS = {
    "info":     {"fg": "#2563EB", "emoji": "\u2139"},   # ℹ
    "success":  {"fg": "#1B873F", "emoji": "\u2714"},   # ✔
    "warning":  {"fg": "#B45309", "emoji": "\u26A0"},   # ⚠
    "error":    {"fg": "#B91C1C", "emoji": "\u26A0"},   # ⚠
    "question": {"fg": "#2563EB", "emoji": "\u2753"},   # ❓
}

_BG_MAIN        = "#FFFFFF"
_BORDER_COLOR   = "#D0D7DE"
_TEXT_COLOR     = "#1F2328"
_FONT_FAMILY    = "Segoe UI"
_MIN_WIDTH      = 320
_MAX_WIDTH      = 620
_PADDING_X      = 22
_PADDING_Y      = 18
_WRAP_MARGIN    = 40
_MARGIN_HEIGHT  = 14

# Accessible Entry Style Colors
_ENTRY_BORDER_NORMAL = "#93A1B5"
_ENTRY_BORDER_FOCUS  = "#2563EB"
_ENTRY_BORDER_ERROR  = "#DC2626"
_ENTRY_BG            = "#FFFFFF"
_ENTRY_FONT          = (_FONT_FAMILY, 10)

# ---------------------------------------------------------------------------
# Helper Functions (Styles / Geometry)
# ---------------------------------------------------------------------------

def _apply_style(context: tk.Toplevel):
    style = ttk.Style(context)
    try:
        if style.theme_use() not in style.theme_names():
            style.theme_use("default")
    except:
        pass

    style.configure("Popup.TFrame", background=_BG_MAIN)
    style.configure("Popup.Icon.TLabel", background=_BG_MAIN, font=(_FONT_FAMILY, 30, "bold"))
    style.configure("Popup.Title.TLabel", background=_BG_MAIN, foreground=_TEXT_COLOR, font=(_FONT_FAMILY, 13, "bold"))
    style.configure("Popup.Msg.TLabel", background=_BG_MAIN, foreground=_TEXT_COLOR, font=(_FONT_FAMILY, 10))

    # Generic button style
    style.configure("Popup.TButton", font=(_FONT_FAMILY, 10), padding=(14, 6))
    style.map("Popup.TButton", background=[("active", "#F1F4F8")])

    # Primary action button style (e.g., Yes, OK, Adopt)
    style.configure("Primary.Popup.TButton", font=(_FONT_FAMILY, 10, "bold"))
    style.map("Primary.Popup.TButton", background=[("active", "#E3F2E9")])

    # Secondary/Cancel action button style
    style.configure("Secondary.Popup.TButton", font=(_FONT_FAMILY, 10))
    style.map("Secondary.Popup.TButton", background=[("active", "#F8E5E5")])

def _center_window(win: tk.Toplevel, parent: tk.Widget = None):
    win.update_idletasks()
    w, h = win.winfo_width(), win.winfo_height()

    if parent and parent.winfo_exists():
        try:
            px, py = parent.winfo_rootx(), parent.winfo_rooty()
            pw, ph = parent.winfo_width(), parent.winfo_height()
            x = px + (pw // 2) - (w // 2)
            y = py + (ph // 2) - (h // 2)
        except:
            sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
            x, y = (sw // 2) - (w // 2), (sh // 2) - (h // 2)
    else:
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        x, y = (sw // 2) - (w // 2), (sh // 2) - (h // 2)

    win.geometry(f"+{max(0, x)}+{max(0, y)}")

def _auto_size(top: tk.Toplevel, container: ttk.Frame, text_widgets):
    top.update_idletasks()
    req_w = max(w.winfo_reqwidth() for w in text_widgets) + (_PADDING_X * 2)
    width = max(_MIN_WIDTH, min(req_w, _MAX_WIDTH))
    for w in text_widgets:
        if isinstance(w, ttk.Label):
            w.configure(wraplength=width - _WRAP_MARGIN)
    top.update_idletasks()
    height = container.winfo_reqheight()
    top.geometry(f"{width}x{height}")

# ---------------------------------------------------------------------------
# Public Popups
# ---------------------------------------------------------------------------

def custom_popup(parent, title, message, kind="info"):
    if kind not in _ACCENTS: kind = "info"

    buttons = [{"key": "ok", "text": "OK", "style": "Primary.Popup.TButton"}]

    custom_dialog(parent, title, message, buttons, kind=kind)

def custom_askyesno(parent, title, message):
    buttons = [
        {"key": "yes", "text": "Yes", "style": "Primary.Popup.TButton"},
        {"key": "no", "text": "No", "style": "Secondary.Popup.TButton"}
    ]
    return custom_dialog(parent, title, message, buttons, kind="question")

def custom_dialog(parent, title, message, buttons, kind="question"):
    if not buttons: return None
    if kind not in _ACCENTS: kind = "question"

    top = tk.Toplevel(parent if parent and parent.winfo_exists() else None)
    top.withdraw()
    top.title(str(title))
    top.configure(bg=_BG_MAIN, highlightbackground=_BORDER_COLOR, highlightthickness=1)
    top.transient(parent)
    top.grab_set()
    top.resizable(False, False)

    result = {"value": buttons[-1]["key"]} # Default to last button on close

    def _finish(val):
        result["value"] = val
        _close()

    def _close():
        try: top.grab_release()
        except: pass
        top.destroy()

    top.protocol("WM_DELETE_WINDOW", _close)
    _apply_style(top)

    container = ttk.Frame(top, style="Popup.TFrame", padding=(_PADDING_X, _PADDING_Y))
    container.pack(fill="both", expand=True)

    content_frame = ttk.Frame(container, style="Popup.TFrame")
    content_frame.pack(fill="x", expand=True, pady=(0, 18))

    accent = _ACCENTS[kind]
    ttk.Label(content_frame, text=accent["emoji"], style="Popup.Icon.TLabel", foreground=accent["fg"]).pack()
    title_lbl = ttk.Label(content_frame, text=str(title), style="Popup.Title.TLabel", justify="center")
    title_lbl.pack(pady=(0, 6))
    msg_lbl = ttk.Label(content_frame, text=str(message), style="Popup.Msg.TLabel", justify="left")
    msg_lbl.pack()

    btn_row = ttk.Frame(container, style="Popup.TFrame")
    btn_row.pack()

    for i, btn_info in enumerate(buttons):
        btn = ttk.Button(btn_row, text=btn_info["text"], style=btn_info.get("style", "Popup.TButton"),
                         command=lambda v=btn_info["key"]: _finish(v))
        btn.pack(side="left", padx=6)
        if i == 0:
            btn.focus_set()
            top.bind("<Return>", lambda e, v=btn_info["key"]: _finish(v))

    top.bind("<Escape>", lambda e: _close())

    _auto_size(top, container, [title_lbl, msg_lbl])
    _center_window(top, parent)
    top.deiconify()
    top.wait_window()
    return result["value"]

def show_toast(parent, message, kind="info", duration=2500):
    # This function remains unchanged
    if kind not in _ACCENTS: kind = "info"
    toast = tk.Toplevel(parent if parent and parent.winfo_exists() else None)
    toast.overrideredirect(True)
    toast.attributes("-topmost", True)
    toast.configure(bg=_BG_MAIN, highlightbackground=_BORDER_COLOR, highlightthickness=1)
    _apply_style(toast)
    frame = ttk.Frame(toast, style="Popup.TFrame", padding=(14, 10))
    frame.pack(fill="both", expand=True)
    accent = _ACCENTS[kind]
    lbl = ttk.Label(frame, text=f"{accent['emoji']}  {message}", style="Popup.Msg.TLabel", foreground=accent["fg"])
    lbl.pack()
    toast.update_idletasks()
    req_w = min(max(frame.winfo_reqwidth() + 10, 240), 400)
    lbl.configure(wraplength=req_w - 30)
    toast.update_idletasks()
    toast.geometry(f"{req_w}x{frame.winfo_reqheight()+6}")
    if parent and parent.winfo_exists():
        px, py = parent.winfo_rootx(), parent.winfo_rooty()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        x, y = px + pw - toast.winfo_width() - 24, py + ph - toast.winfo_height() - 24
    else:
        sw, sh = toast.winfo_screenwidth(), toast.winfo_screenheight()
        x, y = sw - toast.winfo_width() - 30, sh - toast.winfo_height() - 45
    toast.geometry(f"+{x}+{y}")
    toast.after(duration, toast.destroy)

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("860x560+200+120")
    root.title("Popup Utils Demo")

    def demo_custom_dialog():
        buttons = [
            {"key": "adopt", "text": "Adopt Shortest", "style": "Primary.Popup.TButton"},
            {"key": "cancel", "text": "Cancel", "style": "Secondary.Popup.TButton"}
        ]
        result = custom_dialog(root, "Expiry Required", "An item needs an expiry date. What would you like to do?", buttons)
        custom_popup(root, "Result", f"You clicked: '{result}'", "info")

    ttk.Button(root, text="Info Popup", command=lambda: custom_popup(root, "Information", "This is a standard info message.")).pack(pady=5)
    ttk.Button(root, text="Yes/No Question", command=lambda: custom_askyesno(root, "Confirmation", "Are you sure you want to proceed?")).pack(pady=5)
    ttk.Button(root, text="Custom Dialog", command=demo_custom_dialog).pack(pady=5)
    ttk.Button(root, text="Show Toast", command=lambda: show_toast(root, "Operation completed.", "success")).pack(pady=5)
    root.mainloop()
