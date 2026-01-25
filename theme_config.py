"""
Centralized Theme Configuration for ISEPREP Application
=======================================================
Single source of truth for all UI styling, colors, fonts, and ttk configurations.
All modules should import from this file to ensure consistency.

Usage in modules:
    from theme_config import AppTheme, apply_global_style
    
    # At module initialization:
    apply_global_style()  # Call once per Toplevel window
    
    # Use colors:
    frame = tk.Frame(parent, bg=AppTheme.BG_MAIN)
    button = tk.Button(parent, bg=AppTheme.BTN_SUCCESS, fg=AppTheme.TEXT_WHITE)
"""

import tkinter as tk
from tkinter import ttk


# ============================================================
# CENTRALIZED COLOR PALETTE
# ============================================================
class AppTheme:
    """
    Master color palette for entire application.
    DO NOT define colors elsewhere - use these constants.
    """
    
    # -------------------- Background Colors --------------------
    BG_MAIN = "#F0F4F8"          # Main background (light blue-gray)
    BG_PANEL = "#FFFFFF"         # Panel/card backgrounds (white)
    BG_LOGIN = "#F0F4F8"         # Login screen background
    CONTENT_BG = "#F5F5F5"       # Content area background
    ENTRY_BG = "#FFFFFF"         # Text entry background (always white)
    COMBO_DISABLED_BG = "#F0F0F0"  # Combobox disabled background (light gray)
    
    # -------------------- Text Colors --------------------
    COLOR_PRIMARY = "#2C3E50"    # Primary text (dark blue-gray)
    COLOR_SECONDARY = "#7F8C8D"  # Secondary text (medium gray)
    TEXT_DARK = "#1E3A8A"        # Dark text (deep blue)
    TEXT_WHITE = "#FFFFFF"       # White text
    ENTRY_FG = "#1E3A8A"         # Entry field text color
    
    # -------------------- Accent & Highlight Colors --------------------
    COLOR_ACCENT = "#2563EB"     # Primary accent (bright blue)
    COLOR_BORDER = "#D0D7DE"     # Border color (light gray)
    COLOR_SELECTED = "#2563EB"   # Selected item highlight
    
    # -------------------- Button Colors --------------------
    BTN_PRIMARY = "#5DADE2"      # Primary button (sky blue)
    BTN_PRIMARY_HOVER = "#3498DB"  # Primary button hover
    BTN_SUCCESS = "#27AE60"      # Success/Save button (green)
    BTN_DANGER = "#C0392B"       # Delete/Danger button (red)
    BTN_WARNING = "#2980B9"      # Edit/Warning button (blue)
    BTN_NEUTRAL = "#7F8C8D"      # Clear/Cancel button (gray)
    BTN_DISABLED = "#94A3B8"     # Disabled button (light gray)
    BTN_EXPORT = "#2980B9"       # Export button (blue)
    BTN_REFRESH = "#2563EB"      # Refresh button (bright blue)
    BTN_TOGGLE = "#8E44AD"       # Toggle button (purple)
    
    # -------------------- Special UI Element Colors --------------------
    BOTTOM_BAR_BG = "#1D4ED8"    # Bottom navigation bar
    ROW_NORM = "#FFFFFF"         # Normal table row
    ROW_ALT = "#F7FAFC"          # Alternate table row
    ROW_SELECTED = "#2563EB"     # Selected row
    
    # -------------------- Kit/Module/Item Colors --------------------
    KIT_COLOR = "#228B22"        # Kit background (forest green)
    KIT_HEADER_BG = "#E3F6E1"    # Kit header background (light green)
    KIT_DATA_BG = "#C5EDC1"      # Kit data row background
    
    MODULE_COLOR = "#ADD8E6"     # Module background (light blue)
    MODULE_HEADER_BG = "#E1ECFC" # Module header background
    MODULE_DATA_BG = "#C9E2FA"   # Module data row background
    
    ITEM_COLOR = "#222222"       # Item text color (dark)
    
    # -------------------- Status Colors --------------------
    EXPIRED_COLOR = "#FFF5F5"    # Expired items background (light red)
    WARNING_COLOR = "#FFF9C4"    # Warning highlight (light yellow)
    SUCCESS_COLOR = "#E3F6E1"    # Success highlight (light green)
    ERROR_COLOR = "#FFEBEE"      # Error highlight (light red)
    
    # -------------------- Graph Colors --------------------
    LINE_COLOR_IN = "#2E8B57"    # Stock IN graph line (green)
    LINE_COLOR_OUT = "#C0392B"   # Stock OUT graph line (red)
    
    # -------------------- Font Configuration --------------------
    FONT_FAMILY = "Helvetica"
    FONT_SIZE_SMALL = 9
    FONT_SIZE_NORMAL = 10
    FONT_SIZE_HEADING = 11
    FONT_SIZE_LARGE = 12
    FONT_SIZE_TITLE = 14
    FONT_SIZE_HUGE = 20
    FONT_SIZE_ICON = 26
    
    # -------------------- Spacing & Sizing --------------------
    PADDING_SMALL = 4
    PADDING_NORMAL = 8
    PADDING_LARGE = 12
    TREE_ROW_HEIGHT = 26
    BOTTOM_BAR_HEIGHT = 78
    ICON_SIZE = 70
    ICON_HEIGHT = 54


# ============================================================
# GLOBAL STYLE APPLICATION
# ============================================================
def apply_global_style(root=None):
    """
    Apply consistent ttk.Style configuration globally.
    
    This should be called ONCE at application startup (in LoginGUI.__init__)
    or when creating new Toplevel windows.
    
    Args:
        root: Optional root window. If None, applies to default root.
    """
    style = ttk.Style(root) if root else ttk.Style()
    
    # Set base theme
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass  # Theme not available, use default
    
    # -------------------- Treeview Styling --------------------
    style.configure(
        "Treeview",
        background=AppTheme.BG_PANEL,
        fieldbackground=AppTheme.BG_PANEL,
        foreground=AppTheme.COLOR_PRIMARY,
        rowheight=AppTheme.TREE_ROW_HEIGHT,
        font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL),
        bordercolor=AppTheme.COLOR_BORDER,
        relief="flat"
    )
    
    style.map(
        "Treeview",
        background=[("selected", AppTheme.ROW_SELECTED)],
        foreground=[("selected", AppTheme.TEXT_WHITE)]
    )
    
    style.configure(
        "Treeview.Heading",
        background="#E5E8EB",
        foreground=AppTheme.COLOR_PRIMARY,
        font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_HEADING, "bold"),
        relief="flat",
        bordercolor=AppTheme.COLOR_BORDER
    )
    
    # -------------------- Button Styling --------------------
    style.configure(
        "TButton",
        font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL),
        padding=(10, 5),
        borderwidth=0,
        relief="flat"
    )
    
    style.map(
        "TButton",
        background=[("active", AppTheme.BTN_PRIMARY_HOVER)],
        foreground=[("disabled", AppTheme.COLOR_SECONDARY)]
    )
    
    # -------------------- Entry Styling (Always White) --------------------
    style.configure(
        "TEntry",
        font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL),
        fieldbackground=AppTheme.ENTRY_BG,  # Always white
        foreground=AppTheme.ENTRY_FG,
        bordercolor=AppTheme.COLOR_BORDER,
        relief="flat"
    )
    
    # Keep Entry always white (active look)
    style.map(
        "TEntry",
        fieldbackground=[
            ("disabled", AppTheme.COMBO_DISABLED_BG),  # Light gray only when disabled
            ("", AppTheme.ENTRY_BG)                    # White always (active look)
        ],
        bordercolor=[
            ("focus", AppTheme.COLOR_ACCENT),  # Blue border when focused
            ("", AppTheme.COLOR_BORDER)        # Gray border by default
        ]
    )
    
    # -------------------- Combobox Styling (CORRECTED: Always White Like Search Bars) --------------------
    style.configure(
        "TCombobox",
        font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL),
        fieldbackground=AppTheme.ENTRY_BG,     # WHITE by default (active look)
        foreground=AppTheme.ENTRY_FG,
        bordercolor=AppTheme.COLOR_BORDER,
        relief="flat",
        background=AppTheme.ENTRY_BG,          # Dropdown button background also white
        arrowcolor=AppTheme.COLOR_PRIMARY,     # Arrow color (dark)
        arrowsize=12
    )
    
    # CRITICAL FIX: All active dropdowns WHITE, only disabled ones GRAY
    # This matches the search bar behavior (always white/active)
    style.map(
        "TCombobox",
        fieldbackground=[
            ("disabled", AppTheme.COMBO_DISABLED_BG),  # GRAY only when disabled
            ("", AppTheme.ENTRY_BG)                    # WHITE always (active look) ✅
        ],
        foreground=[
            ("disabled", AppTheme.COLOR_SECONDARY),  # Gray text when disabled
            ("", AppTheme.ENTRY_FG)                  # Normal text always
        ],
        bordercolor=[
            ("focus", AppTheme.COLOR_ACCENT),   # Blue border when focused
            ("", AppTheme.COLOR_BORDER)         # Gray border otherwise
        ],
        background=[
            ("disabled", AppTheme.COMBO_DISABLED_BG),  # Gray dropdown button when disabled
            ("", AppTheme.ENTRY_BG)                    # White dropdown button always
        ]
    )
    
    # -------------------- Frame Styling --------------------
    style.configure(
        "TFrame",
        background=AppTheme.BG_MAIN,
        bordercolor=AppTheme.COLOR_BORDER,
        relief="flat"
    )
    
    # -------------------- Label Styling --------------------
    style.configure(
        "TLabel",
        background=AppTheme.BG_MAIN,
        foreground=AppTheme.COLOR_PRIMARY,
        font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL)
    )
    
    # -------------------- Popup-Specific Styles --------------------
    style.configure(
        "Popup.TFrame",
        background=AppTheme.BG_MAIN
    )
    
    style.configure(
        "Popup.Icon.TLabel",
        background=AppTheme.BG_MAIN,
        font=(AppTheme.FONT_FAMILY, 30, "bold")
    )
    
    style.configure(
        "Popup.Title.TLabel",
        background=AppTheme.BG_MAIN,
        foreground=AppTheme.TEXT_DARK,
        font=(AppTheme.FONT_FAMILY, 13, "bold")
    )
    
    style.configure(
        "Popup.Msg.TLabel",
        background=AppTheme.BG_MAIN,
        foreground=AppTheme.COLOR_PRIMARY,
        font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL)
    )
    
    style.configure(
        "Popup.TButton",
        font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL),
        padding=(14, 6)
    )
    
    style.map(
        "Popup.TButton",
        background=[("active", "#F1F4F8")]
    )
    
    style.configure(
        "Primary.Popup.TButton",
        font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL, "bold")
    )


# ============================================================
# HELPER FUNCTIONS
# ============================================================
def get_button_style(button_type="primary"):
    """
    Get consistent button styling dictionary.
    
    Args:
        button_type: One of 'primary', 'success', 'danger', 'warning', 'neutral', 'disabled'
    
    Returns:
        Dictionary with 'bg', 'fg', 'activebackground' keys
    """
    styles = {
        "primary": {
            "bg": AppTheme.BTN_PRIMARY,
            "fg": AppTheme.TEXT_WHITE,
            "activebackground": AppTheme.BTN_PRIMARY_HOVER
        },
        "success": {
            "bg": AppTheme.BTN_SUCCESS,
            "fg": AppTheme.TEXT_WHITE,
            "activebackground": "#229954"
        },
        "danger": {
            "bg": AppTheme.BTN_DANGER,
            "fg": AppTheme.TEXT_WHITE,
            "activebackground": "#A93226"
        },
        "warning": {
            "bg": AppTheme.BTN_WARNING,
            "fg": AppTheme.TEXT_WHITE,
            "activebackground": "#21618C"
        },
        "neutral": {
            "bg": AppTheme.BTN_NEUTRAL,
            "fg": AppTheme.TEXT_WHITE,
            "activebackground": "#5D6D7E"
        },
        "disabled": {
            "bg": AppTheme.BTN_DISABLED,
            "fg": AppTheme.COLOR_SECONDARY,
            "activebackground": AppTheme.BTN_DISABLED
        },
        "export": {
            "bg": AppTheme.BTN_EXPORT,
            "fg": AppTheme.TEXT_WHITE,
            "activebackground": "#21618C"
        },
        "refresh": {
            "bg": AppTheme.BTN_REFRESH,
            "fg": AppTheme.TEXT_WHITE,
            "activebackground": AppTheme.BTN_PRIMARY_HOVER
        }
    }
    return styles.get(button_type, styles["primary"])


def configure_tree_tags(tree):
    """
    Configure standard tree tags for Kit/Module/Item coloring.
    
    Args:
        tree: ttk.Treeview widget
    """
    # Row alternating colors
    tree.tag_configure("norm", background=AppTheme.ROW_NORM)
    tree.tag_configure("alt", background=AppTheme.ROW_ALT)
    
    # Kit tags
    tree.tag_configure("header_kit", 
                      background=AppTheme.KIT_HEADER_BG,
                      font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL, "bold"))
    tree.tag_configure("kit_data", background=AppTheme.KIT_DATA_BG)
    tree.tag_configure("kitrow", 
                      background=AppTheme.KIT_COLOR,
                      foreground=AppTheme.TEXT_WHITE)
    
    # Module tags
    tree.tag_configure("header_module",
                      background=AppTheme.MODULE_HEADER_BG,
                      font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL, "bold"))
    tree.tag_configure("module_header",
                      background=AppTheme.MODULE_HEADER_BG,
                      font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL, "bold"))
    tree.tag_configure("module_data", background=AppTheme.MODULE_DATA_BG)
    
    # Item tags
    tree.tag_configure("item_row", foreground=AppTheme.ITEM_COLOR)
    
    # Editable/Status tags
    tree.tag_configure("editable_row", foreground="#000000")
    tree.tag_configure("non_editable", foreground="#666666")
    tree.tag_configure("Kit_module_highlight", background=AppTheme.WARNING_COLOR)
    
    # Expiry tags
    tree.tag_configure("expired_light", background=AppTheme.EXPIRED_COLOR)


def create_styled_button(parent, text, command, button_type="primary", **kwargs):
    """
    Create a consistently styled button.
    
    Args:
        parent: Parent widget
        text: Button text
        command: Button command
        button_type: Style type (primary/success/danger/warning/neutral)
        **kwargs: Additional button options (will override defaults)
    
    Returns:
        tk.Button widget
    """
    style = get_button_style(button_type)
    
    defaults = {
        "text": text,
        "command": command,
        "font": (AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL, "bold"),
        "relief": "flat",
        "bd": 0,
        "cursor": "hand2",
        "padx": 14,
        "pady": 6
    }
    defaults.update(style)
    defaults.update(kwargs)
    
    return tk.Button(parent, **defaults)


# ============================================================
# MODULE INFO
# ============================================================
__all__ = [
    'AppTheme',
    'apply_global_style',
    'get_button_style',
    'configure_tree_tags',
    'create_styled_button'
]