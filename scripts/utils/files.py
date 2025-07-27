import os
import shutil
from datetime import datetime
import glob

def create_backup(file_path, backup_dir="backups", max_backups=2):
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    # Ensure the backup directory exists
    os.makedirs(backup_dir, exist_ok=True)

    # Extract filename info
    base_name = os.path.basename(file_path)
    name, ext = os.path.splitext(base_name)

    # Create timestamped backup filename
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_name = f"{name}_{timestamp}{ext}"
    backup_path = os.path.join(backup_dir, backup_name)

    # Copy file to create backup
    shutil.copy2(file_path, backup_path)
    print(f"Backup created: {backup_path}")

    # Clean up old backups
    backup_pattern = os.path.join(backup_dir, f"{name}_*{ext}")
    all_backups = sorted(
        glob.glob(backup_pattern),
        key=os.path.getmtime,
        reverse=True  # Newest first
    )

    # Delete older backups beyond max_backups
    for old_backup in all_backups[max_backups:]:
        os.remove(old_backup)
        print(f"Old backup deleted: {old_backup}")

    return backup_path
