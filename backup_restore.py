import os
import zipfile
from datetime import datetime
import shutil
from tkinter import filedialog, Tk, messagebox
import sqlite3

# SQLite database file
DB_FILE = os.path.join(os.path.dirname(__file__), "iseprep.db")

# Base list of source files to ensure key files are always included
BASE_SOURCE_FILES = [
    "item_utils.py", "db.py", "manage_items.py", "login.py", "end_users.py",
    "language_manager.py", "manage_parties.py", "in_.py", "kits_Composition.py",
    "menu_bar.py", "out.py", "stock_data.py", "stock_transactions.py",
    "reports.py", "scenarios.py", "stock_inv.py", "standard_list.py",
    "manage_users.py", "project_details.py", "transaction_utils.py", "translations.py"
]

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
        print(f"Error fetching project code: {e}")
        return "NO_PROJECT"

def create_backup_zip():
    """Create a zip file containing the SQLite database and source files, prompting for save location."""
    # Create a hidden Tk root for the file dialog
    root = Tk()
    root.withdraw()  # Hide the main window

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    project_code = get_project_code()
    default_filename = f"iseprep_backup_{project_code}_{timestamp}.zip"
    
    # Prompt user for save location
    zip_filename = filedialog.asksaveasfilename(
        defaultextension=".zip",
        filetypes=[("Zip Files", "*.zip")],
        initialfile=default_filename,
        initialdir=os.getcwd()  # Default to project root
    )
    root.destroy()  # Close the hidden root

    if not zip_filename:  # User canceled the dialog
        print("Backup canceled by user.")
        messagebox.showinfo("Backup", "Backup canceled by user.")
        return None

    # Get source files and data folder contents
    source_files = get_all_source_files()
    data_folder = os.path.join(os.getcwd(), "data")
    data_files = []
    if os.path.exists(data_folder):
        for root_dir, dirs, files in os.walk(data_folder):
            if "__pycache__" in dirs:
                dirs.remove("__pycache__")  # Exclude __pycache__
            for file in files:
                full_path = os.path.join(root_dir, file)
                arc_name = os.path.relpath(full_path, os.getcwd())
                data_files.append(arc_name)
                print(f"Adding {arc_name} to backup (from data folder)")

    # Create zip file
    try:
        with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Add SQLite database
            if os.path.exists(DB_FILE):
                zipf.write(DB_FILE, os.path.basename(DB_FILE))
                print(f"Added {DB_FILE} to backup")
            else:
                print(f"Warning: SQLite database {DB_FILE} not found")

            # Add source files
            for file in source_files:
                if os.path.exists(file):
                    zipf.write(file, os.path.basename(file))
                    print(f"Added {file} to backup (source file)")

            # Add data folder contents
            for file in data_files:
                full_path = os.path.join(os.getcwd(), file)
                zipf.write(full_path, file)
                print(f"Added {file} to backup (data file)")

        print(f"Backup created: {zip_filename}")
        messagebox.showinfo("Backup", f"Backup created successfully: {zip_filename}")
        return zip_filename
    except Exception as e:
        print(f"Error creating backup: {e}")
        messagebox.showerror("Backup Error", f"Failed to create backup: {str(e)}")
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
        if os.path.exists(db_file):
            shutil.copy(db_file, DB_FILE)
            print(f"Restored SQLite database to {DB_FILE}")

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
                    print(f"Error counting rows in {table}: {e}")

            cursor.close()
            conn.close()
        else:
            print(f"Warning: SQLite database not found in backup")

        # Restore source files and data folder
        source_files = get_all_source_files()
        data_folder = os.path.join(os.getcwd(), "data")
        os.makedirs(data_folder, exist_ok=True)
        restored_files = []

        for root_dir, dirs, files in os.walk(temp_dir):
            if "__pycache__" in dirs:
                dirs.remove("__pycache__")  # Exclude __pycache__
            for file in files:
                full_path = os.path.join(root_dir, file)
                rel_path = os.path.relpath(full_path, temp_dir)
                dest_path = os.path.join(os.getcwd(), rel_path)

                # Handle source files (in main directory)
                if rel_path in source_files or (not rel_path.startswith("data/") and file.endswith(".py")):
                    dest_path = os.path.join(os.getcwd(), file)  # Place in main directory
                # Handle data folder files (e.g., translations)
                elif rel_path.startswith("data/"):
                    dest_path = os.path.join(os.getcwd(), rel_path)  # Preserve data folder structure
                else:
                    continue  # Skip unexpected files

                os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                shutil.copy(full_path, dest_path)
                restored_files.append(rel_path)
                print(f"Restored {rel_path} to {dest_path}")

        # Clean up
        shutil.rmtree(temp_dir)

        # Build confirmation message
        details = ["Restore completed successfully!"]
        details.append(f"Restored database: {DB_FILE}")
        for table, count in restored_counts.items():
            details.append(f" - {table}: {count} rows")
        details.append(f"Restored files: {len(restored_files)}")
        for file in restored_files[:5]:  # Show up to 5 files for brevity
            details.append(f" - {file}")
        if len(restored_files) > 5:
            details.append(f" - ... and {len(restored_files) - 5} more files")

        confirmation = "\n".join(details)
        print(confirmation)
        messagebox.showinfo("Restore Complete", confirmation)
    except Exception as e:
        print(f"Error during restore: {e}")
        messagebox.showerror("Restore Error", f"Failed to restore backup: {str(e)}")
        shutil.rmtree(temp_dir, ignore_errors=True)

if __name__ == "__main__":
    # Perform backup
    create_backup_zip()

    # Example restore (uncomment and provide zip file path to test)
    # restore_backup("iseprep_backup_NO_PROJECT_20250817_152000.zip")