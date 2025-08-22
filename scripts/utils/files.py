import os
import shutil
import sys
from datetime import datetime
import glob
from abc import ABC, abstractmethod
import zipfile
import shutil
import re
from pathlib import Path

from docx import Document
import openpyxl
import pandas as pd

class FileManager(ABC):
    def __init__(self, file_path : str):
        self.file_path = file_path

    @abstractmethod
    def save(self):
        pass

    def backup_and_save(self):
        create_backup(self.file_path)
        self.save()

class ExcelFileManager(FileManager):
    def __init__(self, file_path : str):
        super().__init__(file_path)
        self.xls = pd.ExcelFile(file_path)
        self.wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
        self.updates = {}

    def write(self, sheet_name : str, cell : str, value):
        self.updates[(sheet_name, cell)] = value

    def save(self):
        self._patch_excel_values()

    def _patch_excel_values(self):
        """
        Patch cached values in an .xlsx file, using inlineStr for text.
        updates: {("Sheet1","A1"): value, ...} value can be str or number
        """
        tmp_dir = Path("tmp_extract")
        if tmp_dir.exists():
            shutil.rmtree(tmp_dir)
        tmp_dir.mkdir()

        # Extract .xlsx
        with zipfile.ZipFile(self.file_path, "r") as zf:
            zf.extractall(tmp_dir)

        # Map sheet names -> sheetN.xml
        import xml.etree.ElementTree as ET
        wb_tree = ET.parse(tmp_dir / "xl" / "workbook.xml")
        ns = {"ns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        sheets = {}
        for s in wb_tree.getroot().findall("ns:sheets/ns:sheet", ns):
            sheets[s.attrib["name"]] = f"sheet{s.attrib['sheetId']}.xml"

        for (sheet_name, cell), new_val in self.updates.items():
            sheet_file = tmp_dir / "xl" / "worksheets" / sheets[sheet_name]
            text = sheet_file.read_text(encoding="utf-8")

            if isinstance(new_val, str):
                # Match the full <c ...> ... </c> block
                pattern = re.compile(
                    rf'(<c[^>]*\br="{re.escape(cell)}"[^>]*>).*?(</c>)',
                    flags=re.DOTALL
                )

                def str_repl(m):
                    opening_tag = m.group(1)

                    # Check if t="..." exists; replace or add t="inlineStr"
                    if 't=' in opening_tag:
                        # Replace existing t="..." with t="inlineStr"
                        opening_tag = re.sub(r't="[^"]*"', 't="inlineStr"', opening_tag)
                    else:
                        # Add t="inlineStr" before closing >
                        opening_tag = opening_tag.rstrip('>') + ' t="inlineStr">'

                    # Build the new cell content with inlineStr
                    return f'{opening_tag}<is><t>{new_val}</t></is>{m.group(2)}'

                text, n = pattern.subn(str_repl, text)
            else:
                # Numeric → usual <v> replacement
                pat_replace = re.compile(
                    rf'(?s)(<c[^>]*\br="{re.escape(cell)}"[^>]*>.*?<v>)(.*?)(</v>)'
                )
                def repl(m):
                    return m.group(1) + str(new_val) + m.group(3)
                text, n = pat_replace.subn(repl, text)
                if n == 0:
                    # Insert <v> if missing
                    pat_insert = re.compile(
                        rf'(?s)(<c[^>]*\br="{re.escape(cell)}"[^>]*>)(.*?)(</c>)'
                    )
                    text, n = pat_insert.subn(rf'\1<v>{new_val}</v>\3', text)

            sheet_file.write_text(text, encoding="utf-8")

        # Repack into .xlsx
        shutil.make_archive(self.file_path.replace(".xlsx",""), "zip", tmp_dir)
        shutil.move(self.file_path.replace(".xlsx",".zip"), self.file_path)
        shutil.rmtree(tmp_dir)


class WordFileManager(FileManager):
    def __init__(self, file_path : str):
        super().__init__(file_path)
        self.doc = Document(file_path)

    def save(self):
        self.doc.save(self.file_path)


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
