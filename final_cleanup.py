import os
import shutil
import glob

def clean_project():
    print("Starting cleanup...")
    
    # Files to remove
    files_to_remove = [
        "cleanup_script.py", # Remove self/old script
        "dist/*",
        "build/*",
        "*.spec",
    ]
    
    # Directories to remove
    dirs_to_remove = [
        "__pycache__",
        "*.egg-info",
        "analytics_ui/__pycache__",
        ".pytest_cache"
    ]

    base_dir = os.path.dirname(os.path.abspath(__file__))

    for pattern in files_to_remove:
        path_pattern = os.path.join(base_dir, pattern)
        for file in glob.glob(path_pattern):
             try:
                if os.path.isfile(file):
                    os.remove(file)
                    print(f"Removed file: {file}")
             except Exception as e:
                print(f"Error removing file {file}: {e}")

    for pattern in dirs_to_remove:
        path_pattern = os.path.join(base_dir, pattern)
        for dir_path in glob.glob(path_pattern):
            try:
                if os.path.isdir(dir_path):
                    shutil.rmtree(dir_path)
                    print(f"Removed directory: {dir_path}")
            except Exception as e:
                print(f"Error removing directory {dir_path}: {e}")

    print("Cleanup complete.")

if __name__ == "__main__":
    clean_project()
