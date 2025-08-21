"""
Clean Test Files

This script removes all non-PDF files from the testFiles directory,
except for Excel database files which are needed for testing the application.
"""

import os
from pathlib import Path
import shutil

# Define the test directory
test_dir = Path("D:/rsdan/Documents/VSC_Projects/FileSorter/testFiles")

# Counter for deleted files
deleted_count = 0
preserved_count = 0

# Create a backup directory
backup_dir = test_dir.parent / "testFiles_backup"
if not backup_dir.exists():
    backup_dir.mkdir()
    print(f"Created backup directory: {backup_dir}")

# Process each file in the test directory
for file_path in test_dir.glob("*"):
    if file_path.is_file():
        # Keep PDF files
        if file_path.suffix.lower() == ".pdf":
            preserved_count += 1
            print(f"Preserving PDF file: {file_path.name}")
            continue

        # Keep Excel database files
        if file_path.suffix.lower() in [".xlsx", ".xls"] and "database" in file_path.name.lower():
            preserved_count += 1
            print(f"Preserving database file: {file_path.name}")
            continue

        # Backup and delete all other files
        try:
            # Copy to backup first
            backup_path = backup_dir / file_path.name
            shutil.copy2(file_path, backup_path)

            # Then delete the original
            os.remove(file_path)
            deleted_count += 1
            print(f"Deleted file: {file_path.name}")
        except Exception as e:
            print(f"Error processing {file_path.name}: {str(e)}")

print(f"\nCleanup complete!")
print(f"Files preserved: {preserved_count}")
print(f"Files deleted: {deleted_count}")
print(f"Backup saved to: {backup_dir}")
