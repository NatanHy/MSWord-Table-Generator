import os
import PyInstaller.__main__

# List of GUI Python scripts to compile
files_to_build = [
    "scripts/generation_gui.py",
    "scripts/sync_gui.py"
]

# Build settings (shared)
common_options = [
    "--onefile",        # Create a single .exe
    "--noconsole",      # No console window (for GUI apps)
]

# Paths
RESOURCES_DIR = "resources"
BACKUPS_DIR = "backups"

def build_exe(script):
    print(f"Building {script}...")

    # Add resources folder to exe
    add_data = []
    if os.path.isdir(RESOURCES_DIR):
        add_data = [f"--add-data={RESOURCES_DIR};{RESOURCES_DIR}"]
    else:
        raise RuntimeError(f"Could not find {RESOURCES_DIR}")

    PyInstaller.__main__.run(common_options + add_data + [script])


def ensure_backups_folder():
    dist_dir = "dist"
    backups_path = os.path.join(dist_dir, BACKUPS_DIR)

    if not os.path.exists(backups_path):
        print("Creating backups folder in dist...")
        os.makedirs(backups_path, exist_ok=True)
    else:
        print("Backups folder already exists in dist.")


def main():
    for file in files_to_build:
        if os.path.exists(file):
            build_exe(file)
        else:
            print(f"Error: {file} not found!")

    ensure_backups_folder()
    print("\n✅ Build complete! EXEs are in the 'dist' folder.")


if __name__ == "__main__":
    main()
