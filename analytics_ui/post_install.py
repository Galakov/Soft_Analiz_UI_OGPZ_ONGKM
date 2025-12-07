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

    # Desktop Entry content
    desktop_entry = f"""[Desktop Entry]
Version=1.0
Type=Application
Name=Аналитика УИ ОГПЗ
Comment=Программа для объединения и анализа Excel файлов
Exec={exec_path}
Icon=utilities-terminal
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

if __name__ == "__main__":
    create_shortcuts()
