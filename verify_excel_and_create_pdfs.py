#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Verify Excel Database and Create Additional PDF Test Files

This script:
1. Verifies the Excel database files
2. Creates additional PDF test files
"""

import os
import pandas as pd
from pathlib import Path
import random
from datetime import datetime, timedelta
from fpdf import FPDF
import shutil

# Test directory
TEST_DIR = Path("testFiles")
EXCEL_FILES = ["project_database.xlsx", "comprehensive_project_database.xlsx"]

# Document types for new files
DOCUMENT_TYPES = [
    "Report", "Specification", "Proposal", "Contract", "Invoice",
    "Memo", "Minutes", "Plan", "Presentation", "Analysis",
    "Assessment", "Brief", "Case Study", "Compliance", "Statement",
    "Technical Documentation", "User Guide", "Manual", "Whitepaper", "Review"
]

# Company departments
DEPARTMENTS = [
    "HR", "Finance", "Marketing", "Sales", "IT",
    "Operations", "Legal", "R&D", "Customer Service", "Product Development"
]

# Project prefixes
PROJECT_PREFIXES = [
    "PRJ", "PROJ", "P", "CP", "DP", "SP", "MP", "IP", "TP", "FP"
]


def verify_excel_files():
    """Verify Excel database files and return their contents"""
    print("\n=== Verifying Excel Database Files ===")

    excel_data = {}

    for excel_file in EXCEL_FILES:
        file_path = TEST_DIR / excel_file
        if not file_path.exists():
            print(
                f"Warning: Excel file {excel_file} does not exist in {TEST_DIR}")
            continue

        try:
            # Read Excel file
            df = pd.read_excel(file_path)

            # Display basic info
            print(f"\nExcel file: {excel_file}")
            print(f"Shape: {df.shape[0]} rows, {df.shape[1]} columns")
            print(f"Columns: {', '.join(df.columns)}")

            # Check for missing values
            missing_values = df.isnull().sum().sum()
            if missing_values > 0:
                print(f"Warning: Found {missing_values} missing values")

            # Store dataframe
            excel_data[excel_file] = df

            print(f"✓ Successfully verified {excel_file}")

        except Exception as e:
            print(f"Error reading {excel_file}: {str(e)}")

    return excel_data


def generate_pdf(output_path, title, content, metadata=None):
    """Generate a PDF file with content"""
    pdf = FPDF()
    pdf.add_page()

    # Set title
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, title, ln=True)

    # Add some space
    pdf.ln(5)

    # Set content
    pdf.set_font("Arial", "", 12)

    # Split content into lines and add to PDF
    for line in content.split('\n'):
        pdf.multi_cell(0, 10, line)

    # Add metadata if provided
    if metadata:
        pdf.ln(10)
        pdf.set_font("Arial", "I", 10)
        for key, value in metadata.items():
            pdf.cell(0, 8, f"{key}: {value}", ln=True)

    # Save the PDF
    pdf.output(output_path)


def create_pdf_test_files(excel_data, num_files=10):
    """Create additional PDF test files based on Excel data and random generation"""
    print("\n=== Creating Additional PDF Test Files ===")

    created_files = []

    # Check if we have Excel data to work with
    if not excel_data:
        print("No Excel data available. Creating generic test files.")

        # Create generic test files
        for i in range(1, num_files + 1):
            # Generate random document attributes
            doc_type = random.choice(DOCUMENT_TYPES)
            department = random.choice(DEPARTMENTS)
            project_id = f"{random.choice(PROJECT_PREFIXES)}{random.randint(1000, 9999)}"
            date = (datetime.now() -
                    timedelta(days=random.randint(0, 365))).strftime("%Y-%m-%d")

            # Create file name
            file_name = f"{project_id} - {department} {doc_type} {date}.pdf"
            file_path = TEST_DIR / file_name

            # Create PDF content
            title = f"{department} {doc_type}"
            content = f"This is a test file for {project_id}.\n\n"
            content += f"Department: {department}\n"
            content += f"Document Type: {doc_type}\n"
            content += f"Date: {date}\n\n"
            content += "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed at magna "
            content += "vel risus aliquet gravida. Maecenas eget metus vitae nunc pharetra malesuada."

            metadata = {
                "Project ID": project_id,
                "Department": department,
                "Document Type": doc_type,
                "Date": date
            }

            # Generate PDF
            generate_pdf(file_path, title, content, metadata)
            created_files.append(file_name)

            print(f"Created: {file_name}")
    else:
        # Use Excel data to create test files
        for excel_file, df in excel_data.items():
            print(f"\nCreating files based on {excel_file}:")

            # Determine ID and Name columns
            id_col = next(
                (col for col in df.columns if 'id' in col.lower()), None)
            name_col = next(
                (col for col in df.columns if 'name' in col.lower()), None)

            if not id_col or not name_col:
                print(
                    f"Warning: Could not identify ID or Name columns in {excel_file}")
                continue

            # Create a subset of rows to use for file creation
            if len(df) > num_files:
                subset = df.sample(num_files)
            else:
                subset = df

            # Create files based on subset
            for _, row in subset.iterrows():
                try:
                    project_id = str(row[id_col])
                    name = str(row[name_col])

                    # Create random document attributes
                    doc_type = random.choice(DOCUMENT_TYPES)

                    # Create file name
                    file_name = f"{name} - {doc_type}.pdf"
                    file_path = TEST_DIR / file_name

                    # Create PDF content
                    title = f"{name} - {doc_type}"
                    content = f"Project ID: {project_id}\n"
                    content += f"Project Name: {name}\n"
                    content += f"Document Type: {doc_type}\n\n"

                    # Add other columns as content
                    for col in df.columns:
                        if col != id_col and col != name_col:
                            content += f"{col}: {row[col]}\n"

                    content += "\nLorem ipsum dolor sit amet, consectetur adipiscing elit. "
                    content += "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua."

                    metadata = {
                        "Project ID": project_id,
                        "Project Name": name,
                        "Document Type": doc_type,
                        "Date": datetime.now().strftime("%Y-%m-%d")
                    }

                    # Generate PDF
                    generate_pdf(file_path, title, content, metadata)
                    created_files.append(file_name)

                    print(f"Created: {file_name}")
                except Exception as e:
                    print(f"Error creating file for {project_id}: {str(e)}")

    print(f"\nSuccessfully created {len(created_files)} PDF test files")
    return created_files


def main():
    """Main function"""
    print("Starting Excel verification and PDF test file creation...")

    # Create backup of current test files
    backup_dir = Path("testFiles_new_backup")
    if not backup_dir.exists():
        shutil.copytree(TEST_DIR, backup_dir)
        print(f"Created backup of test files in {backup_dir}")

    # Verify Excel files
    excel_data = verify_excel_files()

    # Create additional PDF test files
    created_files = create_pdf_test_files(excel_data, num_files=15)

    print("\nProcess completed successfully!")
    print(f"You can find all files in the {TEST_DIR} directory")


if __name__ == "__main__":
    main()
