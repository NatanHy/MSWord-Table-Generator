import os
import shutil
import sys
from datetime import datetime
import glob

def _backup_path(file_path, backup_dir="backups") -> str:
    # Extract filename info
    base_name = os.path.basename(file_path)
    name, ext = os.path.splitext(base_name)

    # Create timestamped backup filename
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_name = f"{name}_{timestamp}{ext}"
    backup_path = os.path.join(backup_dir, backup_name)
    return backup_path

def _get_old_backups(file_path, backup_dir="backups"):
    # Extract filename info
    base_name = os.path.basename(file_path)
    name, ext = os.path.splitext(base_name)

    backup_pattern = os.path.join(backup_dir, f"{name}_*{ext}")
    return sorted(
        glob.glob(backup_pattern),
        key=os.path.getmtime,
        reverse=True  # Newest first
    )

def revert_changes_from_backup(original_file_path, backup_dir="backups"):
    backups = _get_old_backups(original_file_path, backup_dir=backup_dir)
    newest_backup = backups[0]

    shutil.copy2(newest_backup, original_file_path)

def create_backup(file_path, backup_dir="backups", max_backups=2):
    # Ensure the backup directory exists
    os.makedirs(backup_dir, exist_ok=True)

    backup_path = _backup_path(file_path, backup_dir=backup_dir)

    # Copy file to create backup
    shutil.copy2(file_path, backup_path)
    print(f"Backup created: {backup_path}")

    # Clean up old backups
    all_backups = _get_old_backups(file_path, backup_dir=backup_dir)

    # Delete older backups beyond max_backups
    for old_backup in all_backups[max_backups:]:
        os.remove(old_backup)
        print(f"Old backup deleted: {old_backup}")

    return backup_path

def resource_path(relative_path):
    """Return absolute path to resource, works in dev and PyInstaller exe"""
    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base_path, relative_path)
