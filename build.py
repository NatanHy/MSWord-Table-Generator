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
    "--noconsole",      # No console window 
    "--clean"
]

# Paths
RESOURCE_DIRS = [
    "resources", 
    "config"     
]
BACKUPS_DIR = "backups"

def build_exe(script):
    print(f"Building {script}...")

    add_data = []

    for rdir in RESOURCE_DIRS:
        if os.path.isdir(rdir):
            add_data.append(f"--add-data={rdir};{rdir}")
        else:
            raise RuntimeError(f"Could not find {rdir}")


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
