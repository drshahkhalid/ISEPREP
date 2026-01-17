import os
import zipfile
from datetime import datetime
import shutil
from tkinter import filedialog, Tk, messagebox
import sqlite3
from language_manager import lang

# SQLite database file
DB_FILE = os.path.join(os.path.dirname(__file__), "iseprep.db")

# Base list of source files to ensure key files are always included
BASE_SOURCE_FILES = [
    "item_utils.py", "db.py", "manage_items.py", "login.py", "end_users.py",
    "language_manager.py", "manage_parties.py", "in_.py", "kits_Composition.py",
    "menu_bar.py", "out.py", "stock_data.py", "stock_transactions. py",
    "reports.py", "scenarios.py", "stock_inv.py", "standard_list.py",
    "manage_users.py", "project_details.py", "transaction_utils.py", "translations.py",
    "backup_restore.py"
]

def t(key, fallback=None, **kwargs):
    """Translation helper for backup_restore module."""
    return lang.t(f"backup_restore.{key}", fallback=fallback, **kwargs)

def get_all_source_files():
    """Dynamically detect all .py files in the current directory, excluding subfolders."""
    current_dir_files = [f for f in os.listdir() if f.endswith(".py") and not f.startswith((".", "__"))]
    return list(set(BASE_SOURCE_FILES + current_dir_files))  # Ensure no duplicates

def get_project_code():
    """Fetch the latest project code from the database using SQLite."""
    try:
        conn = sqlite3.connect(DB_FILE)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute("SELECT project_code FROM project_details ORDER BY id DESC LIMIT 1")
        result = cursor.fetchone()
        cursor.close()
        conn.close()
        return result["project_code"] if result else "NO_PROJECT"
    except sqlite3.Error as e:
        print(f"{t('error_fetching_project', fallback='Error fetching project code')}: {e}")
        return "NO_PROJECT"

def create_backup_zip():
    """Create a zip file containing the SQLite database and source files, prompting for save location."""
    # Create a hidden Tk root for the file dialog
    root = Tk()
    root.withdraw()  # Hide the main window

    timestamp = datetime. now().strftime("%Y%m%d_%H%M%S")
    project_code = get_project_code()
    default_filename = f"iseprep_backup_{project_code}_{timestamp}.zip"
    
    # Prompt user for save location
    zip_filename = filedialog.asksaveasfilename(
        title=t("save_backup_title", fallback="Save Backup"),
        defaultextension=".zip",
        filetypes=[(t("zip_files", fallback="Zip Files"), "*.zip")],
        initialfile=default_filename,
        initialdir=os.getcwd()  # Default to project root
    )
    root.destroy()  # Close the hidden root

    if not zip_filename:  # User canceled the dialog
        print(t("backup_canceled", fallback="Backup canceled by user."))
        messagebox.showinfo(
            t("backup_title", fallback="Backup"),
            t("backup_canceled", fallback="Backup canceled by user.")
        )
        return None

    # Get source files and data folder contents
    source_files = get_all_source_files()
    data_folder = os. path.join(os.getcwd(), "data")
    data_files = []
    if os. path.exists(data_folder):
        for root_dir, dirs, files in os.walk(data_folder):
            if "__pycache__" in dirs: 
                dirs.remove("__pycache__")  # Exclude __pycache__
            for file in files:
                full_path = os.path.join(root_dir, file)
                arc_name = os.path.relpath(full_path, os. getcwd())
                data_files.append(arc_name)
                print(t("adding_to_backup", fallback="Adding {file} to backup (from data folder)").format(file=arc_name))

    # Create zip file
    try: 
        with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Add SQLite database
            if os.path.exists(DB_FILE):
                zipf.write(DB_FILE, os.path.basename(DB_FILE))
                print(t("added_db_to_backup", fallback="Added {db} to backup").format(db=DB_FILE))
            else:
                print(t("warning_db_not_found", fallback="Warning: SQLite database {db} not found").format(db=DB_FILE))

            # Add source files
            for file in source_files:
                if os.path.exists(file):
                    zipf.write(file, os.path.basename(file))
                    print(t("added_source_file", fallback="Added {file} to backup (source file)").format(file=file))

            # Add data folder contents
            for file in data_files:
                full_path = os.path.join(os.getcwd(), file)
                zipf.write(full_path, file)
                print(t("added_data_file", fallback="Added {file} to backup (data file)").format(file=file))

        print(t("backup_created", fallback="Backup created: {filename}").format(filename=zip_filename))
        messagebox.showinfo(
            t("backup_title", fallback="Backup"),
            t("backup_success", fallback="Backup created successfully: {filename}").format(filename=zip_filename)
        )
        return zip_filename
    except Exception as e:
        print(t("error_creating_backup", fallback="Error creating backup: {error}").format(error=str(e)))
        messagebox.showerror(
            t("backup_error_title", fallback="Backup Error"),
            t("backup_failed", fallback="Failed to create backup: {error}").format(error=str(e))
        )
        return None

def restore_backup(zip_file):
    """Restore the backup from a zip file, preserving folder structure and showing restore details."""
    try:
        # Create temporary directory for extraction
        temp_dir = "temp_restore"
        os.makedirs(temp_dir, exist_ok=True)

        # Extract zip contents
        with zipfile.ZipFile(zip_file, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # Restore SQLite database
        db_file = os.path.join(temp_dir, "iseprep.db")
        restored_tables = []
        restored_counts = {}
        if os. path.exists(db_file):
            shutil.copy(db_file, DB_FILE)
            print(t("restored_db", fallback="Restored SQLite database to {db}").format(db=DB_FILE))

            # Verify restored data
            conn = sqlite3.connect(DB_FILE)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()

            # Check key tables
            tables = ["scenarios", "compositions", "items_list"]
            for table in tables:
                try:
                    cursor.execute(f"SELECT COUNT(*) AS count FROM {table}")
                    count = cursor.fetchone()["count"]
                    restored_counts[table] = count
                    restored_tables.append(table)
                except sqlite3.Error as e:
                    print(t("error_counting_rows", fallback="Error counting rows in {table}: {error}").format(table=table, error=str(e)))

            cursor.close()
            conn.close()
        else:
            print(t("warning_db_not_in_backup", fallback="Warning:  SQLite database not found in backup"))

        # Restore source files and data folder
        source_files = get_all_source_files()
        data_folder = os. path.join(os.getcwd(), "data")
        os.makedirs(data_folder, exist_ok=True)
        restored_files = []

        for root_dir, dirs, files in os.walk(temp_dir):
            if "__pycache__" in dirs:
                dirs.remove("__pycache__")  # Exclude __pycache__
            for file in files: 
                full_path = os.path.join(root_dir, file)
                rel_path = os.path.relpath(full_path, temp_dir)
                dest_path = os.path. join(os.getcwd(), rel_path)

                # Handle source files (in main directory)
                if rel_path in source_files or (not rel_path.startswith("data/") and file.endswith(".py")):
                    dest_path = os.path.join(os.getcwd(), file)  # Place in main directory
                # Handle data folder files (e.g., translations)
                elif rel_path.startswith("data/"):
                    dest_path = os.path.join(os.getcwd(), rel_path)  # Preserve data folder structure
                else:
                    continue  # Skip unexpected files

                os.makedirs(os.path. dirname(dest_path), exist_ok=True)
                shutil.copy(full_path, dest_path)
                restored_files.append(rel_path)
                print(t("restored_file", fallback="Restored {file} to {dest}").format(file=rel_path, dest=dest_path))

        # Clean up
        shutil.rmtree(temp_dir)

        # Build confirmation message
        details = [t("restore_success", fallback="Restore completed successfully!")]
        details.append(t("restored_database", fallback="Restored database:  {db}").format(db=DB_FILE))
        for table, count in restored_counts.items():
            details.append(t("table_rows", fallback=" - {table}: {count} rows").format(table=table, count=count))
        details.append(t("restored_files_count", fallback="Restored files: {count}").format(count=len(restored_files)))
        for file in restored_files[: 5]:  # Show up to 5 files for brevity
            details.append(f" - {file}")
        if len(restored_files) > 5:
            details.append(t("more_files", fallback=" - ... and {count} more files").format(count=len(restored_files) - 5))

        confirmation = "\n".join(details)
        print(confirmation)
        messagebox.showinfo(
            t("restore_complete_title", fallback="Restore Complete"),
            confirmation
        )
    except Exception as e:
        print(t("error_during_restore", fallback="Error during restore: {error}").format(error=str(e)))
        messagebox.showerror(
            t("restore_error_title", fallback="Restore Error"),
            t("restore_failed", fallback="Failed to restore backup: {error}").format(error=str(e))
        )
        shutil.rmtree(temp_dir, ignore_errors=True)

def select_and_restore_backup():
    """Prompt user to select a backup file and restore it."""
    root = Tk()
    root.withdraw()  # Hide the main window
    
    zip_file = filedialog.askopenfilename(
        title=t("select_backup_title", fallback="Select Backup File"),
        filetypes=[(t("zip_files", fallback="Zip Files"), "*.zip")],
        initialdir=os.getcwd()
    )
    root.destroy()
    
    if not zip_file:
        print(t("restore_canceled", fallback="Restore canceled by user."))
        messagebox.showinfo(
            t("restore_title", fallback="Restore"),
            t("restore_canceled", fallback="Restore canceled by user.")
        )
        return
    
    # Confirm restore
    confirm = messagebox.askyesno(
        t("confirm_restore_title", fallback="Confirm Restore"),
        t("confirm_restore_message", fallback="This will overwrite current data. Continue?")
    )
    
    if confirm:
        restore_backup(zip_file)
    else:
        print(t("restore_canceled", fallback="Restore canceled by user."))
        messagebox.showinfo(
            t("restore_title", fallback="Restore"),
            t("restore_canceled", fallback="Restore canceled by user.")
        )

if __name__ == "__main__":
    # Perform backup
    create_backup_zip()

    # Example restore (uncomment to test)
    # select_and_restore_backup()