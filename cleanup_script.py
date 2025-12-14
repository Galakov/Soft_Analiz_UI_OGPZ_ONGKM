import os
import shutil
import sys

def cleanup():
    # List of files to delete
    files_to_delete = [
        "check_rules.py",
        "debug_rules.py",
        "inspect_rules.py",
        "verify_changes.py"
    ]

    # List of directories to delete
    dirs_to_delete = [
        "legacy",
        "scripts",
        "logs",
        "dist",
        "analytics_ui_ogpz.egg-info"
    ]

    base_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"Cleaning up in: {base_dir}")

    # Delete files
    for filename in files_to_delete:
        filepath = os.path.join(base_dir, filename)
        if os.path.exists(filepath):
            try:
                os.remove(filepath)
                print(f"Deleted file: {filename}")
            except Exception as e:
                print(f"Error deleting {filename}: {e}")
        else:
            print(f"File not found (already deleted?): {filename}")

    # Delete directories
    for dirname in dirs_to_delete:
        dirpath = os.path.join(base_dir, dirname)
        if os.path.exists(dirpath):
            try:
                shutil.rmtree(dirpath)
                print(f"Deleted directory: {dirname}")
            except Exception as e:
                print(f"Error deleting {dirname}: {e}")
        else:
            print(f"Directory not found (already deleted?): {dirname}")

    print("\nCleanup finished!")
    input("Press Enter to exit...")

if __name__ == "__main__":
    cleanup()
