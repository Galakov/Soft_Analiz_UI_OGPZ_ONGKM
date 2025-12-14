import os
import sys
import shutil
from pathlib import Path

def create_shortcuts():
    """Creates desktop shortcuts for the application."""
    print("Setting up shortcuts for Analytics UI...")

    # Define paths
    home = Path.home()
    desktop_dir = home / "Desktop"
    if not desktop_dir.exists():
        desktop_dir = home / "Рабочий стол"
    
    applications_dir = home / ".local" / "share" / "applications"
    
    # Ensure directories exist
    applications_dir.mkdir(parents=True, exist_ok=True)
    
    # Find the executable path
    # When installed via pip, the executable 'analytics-ui' should be in the path
    # We can try to find it using shutil.which
    exec_path = shutil.which("analytics-ui")
    
    if not exec_path:
        print("Warning: Could not find 'analytics-ui' in PATH.")
        print("Assuming it is in ~/.local/bin/analytics-ui")
        exec_path = str(home / ".local" / "bin" / "analytics-ui")
    
    print(f"Executable path: {exec_path}")

    # Find icon path
    icon_path = Path(__file__).parent / "icon.png"
    if not icon_path.exists():
        icon_path = "utilities-terminal" # Fallback
    else:
        icon_path = str(icon_path)

    # Desktop Entry content
    desktop_entry = f"""[Desktop Entry]
Version=1.0
Type=Application
Name=Аналитика УИ ОГПЗ
Comment=Программа для объединения и анализа Excel файлов
Exec={exec_path}
Icon={icon_path}
Terminal=false
Categories=Utility;Office;
StartupNotify=true
"""

    # Create shortcut in Applications menu
    app_shortcut = applications_dir / "analytics_ui.desktop"
    try:
        with open(app_shortcut, "w", encoding="utf-8") as f:
            f.write(desktop_entry)
        print(f"Created menu shortcut: {app_shortcut}")
        
        # Make executable
        app_shortcut.chmod(0o755)
    except Exception as e:
        print(f"Error creating menu shortcut: {e}")

    # Create shortcut on Desktop
    if desktop_dir.exists():
        desktop_shortcut = desktop_dir / "analytics_ui.desktop"
        try:
            with open(desktop_shortcut, "w", encoding="utf-8") as f:
                f.write(desktop_entry)
            print(f"Created desktop shortcut: {desktop_shortcut}")
            
            # Make executable (important for GNOME/RedOS)
            desktop_shortcut.chmod(0o755)
            
            # Allow launching (trusted) - specific to some DEs, but chmod usually helps
        except Exception as e:
            print(f"Error creating desktop shortcut: {e}")
    else:
        print("Desktop directory not found, skipping desktop shortcut.")

    print("\nSetup complete! You may need to restart your session or refresh the desktop.")

def remove_shortcuts():
    """Removes desktop shortcuts for the application."""
    print("Removing shortcuts for Analytics UI...")
    
    home = Path.home()
    
    # Paths to remove
    paths_to_remove = [
        home / ".local" / "share" / "applications" / "analytics_ui.desktop",
        home / "Desktop" / "analytics_ui.desktop",
        home / "Рабочий стол" / "analytics_ui.desktop"
    ]
    
    for path in paths_to_remove:
        if path.exists():
            try:
                path.unlink()
                print(f"Removed: {path}")
            except Exception as e:
                print(f"Error removing {path}: {e}")
        else:
            pass # File doesn't exist, which is fine

    print("\nShortcuts removed.")

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--remove":
        remove_shortcuts()
    else:
        create_shortcuts()
