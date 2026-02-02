"""
Centralized Theme Configuration for ISEPREP Application
=======================================================
Single source of truth for all UI styling, colors, fonts, and ttk configurations.
All modules should import from this file to ensure consistency.

Usage in modules:
    from theme_config import AppTheme, apply_global_style, apply_multiline_headings

    # At module initialization:
    apply_global_style()  # Call once per Toplevel window

    # OPTIONAL: For multi-line headings (call AFTER tree.heading() setup):
    apply_multiline_headings(tree, columns, headings_dict)

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
    BG_MAIN = "#F0F4F8"  # Main background (light blue-gray)
    BG_PANEL = "#FFFFFF"  # Panel/card backgrounds (white)
    BG_LOGIN = "#F0F4F8"  # Login screen background
    CONTENT_BG = "#F5F5F5"  # Content area background
    ENTRY_BG = "#FFFFFF"  # Text entry background (always white)
    COMBO_DISABLED_BG = "#F0F0F0"  # Combobox disabled background (light gray)

    # -------------------- Text Colors --------------------
    COLOR_PRIMARY = "#2C3E50"  # Primary text (dark blue-gray)
    COLOR_SECONDARY = "#7F8C8D"  # Secondary text (medium gray)
    TEXT_DARK = "#1E3A8A"  # Dark text (deep blue)
    TEXT_WHITE = "#FFFFFF"  # White text
    ENTRY_FG = "#1E3A8A"  # Entry field text color

    # -------------------- Accent & Highlight Colors --------------------
    COLOR_ACCENT = "#2563EB"  # Primary accent (bright blue)
    COLOR_BORDER = "#D0D7DE"  # Border color (light gray)
    COLOR_SELECTED = "#2563EB"  # Selected item highlight

    # -------------------- Button Colors --------------------
    BTN_PRIMARY = "#5DADE2"  # Primary button (sky blue)
    BTN_PRIMARY_HOVER = "#3498DB"  # Primary button hover
    BTN_SUCCESS = "#27AE60"  # Success/Save button (green)
    BTN_DANGER = "#C0392B"  # Delete/Danger button (red)
    BTN_WARNING = "#2980B9"  # Edit/Warning button (blue)
    BTN_NEUTRAL = "#7F8C8D"  # Clear/Cancel button (gray)
    BTN_DISABLED = "#94A3B8"  # Disabled button (light gray)
    BTN_EXPORT = "#2980B9"  # Export button (blue)
    BTN_REFRESH = "#2563EB"  # Refresh button (bright blue)
    BTN_TOGGLE = "#8E44AD"  # Toggle button (purple)

    # -------------------- Special UI Element Colors --------------------
    BOTTOM_BAR_BG = "#1D4ED8"  # Bottom navigation bar
    ROW_NORM = "#FFFFFF"  # Normal table row
    ROW_ALT = "#F7FAFC"  # Alternate table row
    ROW_SELECTED = "#2563EB"  # Selected row

    # -------------------- Kit/Module/Item Colors --------------------
    KIT_COLOR = "#228B22"  # Kit background (forest green)
    KIT_HEADER_BG = "#E3F6E1"  # Kit header background (light green)
    KIT_DATA_BG = "#C5EDC1"  # Kit data row background

    MODULE_COLOR = "#ADD8E6"  # Module background (light blue)
    MODULE_HEADER_BG = "#E1ECFC"  # Module header background
    MODULE_DATA_BG = "#C9E2FA"  # Module data row background

    ITEM_COLOR = "#222222"  # Item text color (dark)

    # -------------------- Status Colors --------------------
    EXPIRED_COLOR = "#FFF5F5"  # Expired items background (light red)
    WARNING_COLOR = "#FFF9C4"  # Warning highlight (light yellow)
    SUCCESS_COLOR = "#E3F6E1"  # Success highlight (light green)
    ERROR_COLOR = "#FFEBEE"  # Error highlight (light red)

    # -------------------- Graph Colors --------------------
    LINE_COLOR_IN = "#2E8B57"  # Stock IN graph line (green)
    LINE_COLOR_OUT = "#C0392B"  # Stock OUT graph line (red)

    # -------------------- Font Configuration --------------------
    FONT_FAMILY = "Helvetica"
    FONT_SIZE_SMALL = 9
    FONT_SIZE_NORMAL = 10
    FONT_SIZE_HEADING = 9  # Column heading font size (smaller, normal weight)
    FONT_SIZE_LARGE = 12
    FONT_SIZE_TITLE = 14
    FONT_SIZE_HUGE = 20
    FONT_SIZE_ICON = 26

    # -------------------- Spacing & Sizing --------------------
    PADDING_SMALL = 4
    PADDING_NORMAL = 8
    PADDING_LARGE = 12
    TREE_ROW_HEIGHT = 26  # Data row height
    TREE_HEADING_HEIGHT = 40  # Heading row height (for 2-line headings)
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
        relief="flat",
    )

    style.map(
        "Treeview",
        background=[("selected", AppTheme.ROW_SELECTED)],
        foreground=[("selected", AppTheme.TEXT_WHITE)],
    )

    # -------------------- Treeview Heading (Normal Font, Not Bold) --------------------
    style.configure(
        "Treeview.Heading",
        background="#E5E8EB",
        foreground=AppTheme.COLOR_PRIMARY,
        font=(
            AppTheme.FONT_FAMILY,
            AppTheme.FONT_SIZE_HEADING,
        ),  # Normal weight (NOT bold)
        relief="flat",
        bordercolor=AppTheme.COLOR_BORDER,
        anchor="center",
    )

    # -------------------- Button Styling --------------------
    style.configure(
        "TButton",
        font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL),
        padding=(10, 5),
        borderwidth=0,
        relief="flat",
    )

    style.map(
        "TButton",
        background=[("active", AppTheme.BTN_PRIMARY_HOVER)],
        foreground=[("disabled", AppTheme.COLOR_SECONDARY)],
    )

    # -------------------- Entry Styling (Always White) --------------------
    style.configure(
        "TEntry",
        font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL),
        fieldbackground=AppTheme.ENTRY_BG,
        foreground=AppTheme.ENTRY_FG,
        bordercolor=AppTheme.COLOR_BORDER,
        relief="flat",
    )

    style.map(
        "TEntry",
        fieldbackground=[
            ("disabled", AppTheme.COMBO_DISABLED_BG),
            ("", AppTheme.ENTRY_BG),
        ],
        bordercolor=[("focus", AppTheme.COLOR_ACCENT), ("", AppTheme.COLOR_BORDER)],
    )

    # -------------------- Combobox Styling --------------------
    style.configure(
        "TCombobox",
        font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL),
        fieldbackground=AppTheme.ENTRY_BG,
        foreground=AppTheme.ENTRY_FG,
        bordercolor=AppTheme.COLOR_BORDER,
        relief="flat",
        background=AppTheme.ENTRY_BG,
        arrowcolor=AppTheme.COLOR_PRIMARY,
        arrowsize=12,
    )

    style.map(
        "TCombobox",
        fieldbackground=[
            ("disabled", AppTheme.COMBO_DISABLED_BG),
            ("", AppTheme.ENTRY_BG),
        ],
        foreground=[("disabled", AppTheme.COLOR_SECONDARY), ("", AppTheme.ENTRY_FG)],
        bordercolor=[("focus", AppTheme.COLOR_ACCENT), ("", AppTheme.COLOR_BORDER)],
        background=[("disabled", AppTheme.COMBO_DISABLED_BG), ("", AppTheme.ENTRY_BG)],
    )

    # -------------------- Frame Styling --------------------
    style.configure(
        "TFrame",
        background=AppTheme.BG_MAIN,
        bordercolor=AppTheme.COLOR_BORDER,
        relief="flat",
    )

    # -------------------- Label Styling --------------------
    style.configure(
        "TLabel",
        background=AppTheme.BG_MAIN,
        foreground=AppTheme.COLOR_PRIMARY,
        font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL),
    )

    # -------------------- Popup-Specific Styles --------------------
    style.configure("Popup.TFrame", background=AppTheme.BG_MAIN)

    style.configure(
        "Popup.Icon.TLabel",
        background=AppTheme.BG_MAIN,
        font=(AppTheme.FONT_FAMILY, 30, "bold"),
    )

    style.configure(
        "Popup.Title.TLabel",
        background=AppTheme.BG_MAIN,
        foreground=AppTheme.TEXT_DARK,
        font=(AppTheme.FONT_FAMILY, 13, "bold"),
    )

    style.configure(
        "Popup.Msg.TLabel",
        background=AppTheme.BG_MAIN,
        foreground=AppTheme.COLOR_PRIMARY,
        font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL),
    )

    style.configure(
        "Popup.TButton",
        font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL),
        padding=(14, 6),
    )

    style.map("Popup.TButton", background=[("active", "#F1F4F8")])

    style.configure(
        "Primary.Popup.TButton",
        font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL, "bold"),
    )


# ============================================================
# MULTI-LINE HEADING WORKAROUND (OPTIONAL)
# ============================================================
def apply_multiline_headings(tree, columns, headings_dict, widths_dict=None):
    """
    WORKAROUND: Apply multi-line headings to Treeview columns.

    TTK Treeview doesn't natively support \\n in headings, so this function
    manually splits long text and applies it.

    Usage (OPTIONAL - call after setting up tree headings):
        headings = {"code": "Item Code", "qty": "Quantity\\nExpiring"}
        widths = {"code": 100, "qty": 120}
        apply_multiline_headings(self.tree, cols, headings, widths)

    Args:
        tree: ttk.Treeview widget
        columns: List of column IDs
        headings_dict: Dictionary mapping column IDs to heading text
        widths_dict: Optional dictionary of column widths (for smart wrapping)

    Note: This is OPTIONAL. Only use if you want multi-line headings.
          Otherwise, just use single-line headings as-is.
    """
    for col in columns:
        heading_text = headings_dict.get(col, col)

        # If heading already has \n, keep it (won't work in TTK, but kept for compatibility)
        if "\n" in heading_text:
            # TTK limitation: can't render \n directly
            # Best we can do is keep the text as-is (will show single line)
            tree.heading(col, text=heading_text.replace("\n", " "))
        else:
            tree.heading(col, text=heading_text)


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
            "activebackground": AppTheme.BTN_PRIMARY_HOVER,
        },
        "success": {
            "bg": AppTheme.BTN_SUCCESS,
            "fg": AppTheme.TEXT_WHITE,
            "activebackground": "#229954",
        },
        "danger": {
            "bg": AppTheme.BTN_DANGER,
            "fg": AppTheme.TEXT_WHITE,
            "activebackground": "#A93226",
        },
        "warning": {
            "bg": AppTheme.BTN_WARNING,
            "fg": AppTheme.TEXT_WHITE,
            "activebackground": "#21618C",
        },
        "neutral": {
            "bg": AppTheme.BTN_NEUTRAL,
            "fg": AppTheme.TEXT_WHITE,
            "activebackground": "#5D6D7E",
        },
        "disabled": {
            "bg": AppTheme.BTN_DISABLED,
            "fg": AppTheme.COLOR_SECONDARY,
            "activebackground": AppTheme.BTN_DISABLED,
        },
        "export": {
            "bg": AppTheme.BTN_EXPORT,
            "fg": AppTheme.TEXT_WHITE,
            "activebackground": "#21618C",
        },
        "refresh": {
            "bg": AppTheme.BTN_REFRESH,
            "fg": AppTheme.TEXT_WHITE,
            "activebackground": AppTheme.BTN_PRIMARY_HOVER,
        },
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
    tree.tag_configure(
        "header_kit",
        background=AppTheme.KIT_HEADER_BG,
        font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL, "bold"),
    )
    tree.tag_configure("kit_data", background=AppTheme.KIT_DATA_BG)
    tree.tag_configure(
        "kitrow", background=AppTheme.KIT_COLOR, foreground=AppTheme.TEXT_WHITE
    )

    # Module tags
    tree.tag_configure(
        "header_module",
        background=AppTheme.MODULE_HEADER_BG,
        font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL, "bold"),
    )
    tree.tag_configure(
        "module_header",
        background=AppTheme.MODULE_HEADER_BG,
        font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL, "bold"),
    )
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
        "pady": 6,
    }
    defaults.update(style)
    defaults.update(kwargs)

    return tk.Button(parent, **defaults)

    # ============================================================
    # AUTO-ADJUST COLUMN WIDTH (DOUBLE-CLICK)
    # ============================================================
    def enable_column_auto_resize(tree):
        """
        Enable double-click on column headers to auto-resize based on content.

        Usage (call once after tree is created and populated):
            tree = ttk.Treeview(...)
            # ... set up columns and data ...
            enable_column_auto_resize(tree)

        Args:
            tree: ttk.Treeview widget

        Features:
            - Double-click column header → auto-resize to fit content
            - Measures both header text and data content
            - Sets optimal width with padding
        """

    def auto_resize_column(event):
        """Auto-resize column on double-click"""
        # Identify which column was clicked
        region = tree.identify("region", event.x, event.y)
        if region != "heading":
            return

        column = tree.identify_column(event.x)
        if not column:
            return

        # Convert column identifier (e.g., '#1') to column name
        col_index = int(column.replace("#", "")) - 1
        columns = tree["columns"]

        if col_index < 0 or col_index >= len(columns):
            return

        col_name = columns[col_index]

        # Measure header width
        heading_text = tree.heading(col_name, "text")
        if not heading_text:
            heading_text = col_name

        # Calculate text width using font
        try:
            font = (AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_HEADING)
            import tkinter.font as tkfont

            tk_font = tkfont.Font(family=font[0], size=font[1])
            header_width = tk_font.measure(str(heading_text))
        except:
            header_width = len(str(heading_text)) * 8  # Fallback

        # Measure content width (sample first 100 rows for performance)
        max_width = header_width
        data_font = (AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL)

        try:
            import tkinter.font as tkfont

            tk_data_font = tkfont.Font(family=data_font[0], size=data_font[1])

            # Get all items (or first 100 for large datasets)
            items = tree.get_children()
            sample_size = min(100, len(items))

            for item in items[:sample_size]:
                values = tree.item(item, "values")
                if col_index < len(values):
                    cell_text = str(values[col_index])
                    cell_width = tk_data_font.measure(cell_text)
                    max_width = max(max_width, cell_width)
        except:
            # Fallback: character count estimation
            items = tree.get_children()
            for item in items[:100]:
                values = tree.item(item, "values")
                if col_index < len(values):
                    cell_text = str(values[col_index])
                    max_width = max(max_width, len(cell_text) * 8)

        # Add padding (20px left + 20px right)
        optimal_width = max_width + 40

        # Set minimum and maximum bounds
        optimal_width = max(50, min(optimal_width, 400))

        # Apply new width
        tree.column(col_name, width=optimal_width)

        # Bind double-click event to tree headings
        tree.bind("<Double-Button-1>", auto_resize_column)


# ============================================================
# AUTO-ADJUST COLUMN WIDTH (DOUBLE-CLICK)
# ============================================================
def enable_column_auto_resize(tree):
    """
    Enable double-click on column headers to auto-resize based on content.

    Usage (call once after tree is created and populated):
        tree = ttk.Treeview(...)
        # ... set up columns and data ...
        enable_column_auto_resize(tree)

    Args:
        tree: ttk.Treeview widget

    Features:
        - Double-click column header → auto-resize to fit content
        - Measures both header text and data content
        - Sets optimal width with padding
    """

    def auto_resize_column(event):
        """Auto-resize column on double-click"""
        # Identify which column was clicked
        region = tree.identify("region", event.x, event.y)
        if region != "heading":
            return

        column = tree.identify_column(event.x)
        if not column:
            return

        # Convert column identifier (e.g., '#1') to column name
        col_index = int(column.replace("#", "")) - 1
        columns = tree["columns"]

        if col_index < 0 or col_index >= len(columns):
            return

        col_name = columns[col_index]

        # Measure header width
        heading_text = tree.heading(col_name, "text")
        if not heading_text:
            heading_text = col_name

        # Calculate text width using font
        try:
            font = (AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_HEADING)
            import tkinter.font as tkfont

            tk_font = tkfont.Font(family=font[0], size=font[1])
            header_width = tk_font.measure(str(heading_text))
        except:
            header_width = len(str(heading_text)) * 8  # Fallback

        # Measure content width (sample first 100 rows for performance)
        max_width = header_width
        data_font = (AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL)

        try:
            import tkinter.font as tkfont

            tk_data_font = tkfont.Font(family=data_font[0], size=data_font[1])

            # Get all items (or first 100 for large datasets)
            items = tree.get_children()
            sample_size = min(100, len(items))

            for item in items[:sample_size]:
                values = tree.item(item, "values")
                if col_index < len(values):
                    cell_text = str(values[col_index])
                    cell_width = tk_data_font.measure(cell_text)
                    max_width = max(max_width, cell_width)
        except:
            # Fallback: character count estimation
            items = tree.get_children()
            for item in items[:100]:
                values = tree.item(item, "values")
                if col_index < len(values):
                    cell_text = str(values[col_index])
                    max_width = max(max_width, len(cell_text) * 8)

        # Add padding (20px left + 20px right)
        optimal_width = max_width + 40

        # Set minimum and maximum bounds
        optimal_width = max(50, min(optimal_width, 400))

        # Apply new width
        tree.column(col_name, width=optimal_width)

    # Bind double-click event to tree headings
    tree.bind("<Double-Button-1>", auto_resize_column)


# ============================================================
# MODULE INFO
# ============================================================
__all__ = [
    "AppTheme",
    "apply_global_style",
    "apply_multiline_headings",
    "enable_column_auto_resize",
    "get_button_style",
    "configure_tree_tags",
    "create_styled_button",
]
