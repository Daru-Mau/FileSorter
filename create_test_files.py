import pandas as pd
import os
from pathlib import Path
import random
import datetime
import string
import shutil

# Create test directory if it doesn't exist
test_dir = Path("D:/rsdan/Documents/VSC_Projects/FileSorter/testFiles")
test_dir.mkdir(exist_ok=True)

# Clear existing files
for file in test_dir.glob("*"):
    if file.is_file() and not file.name.endswith(".xlsx"):
        file.unlink()

# Define data for test files
project_data = [
    {
        "ID": "PRJ001",
        "ProjectName": "Marketing Campaign 2025",
        "ClientName": "Global Foods Inc.",
        "Department": "Marketing",
        "Category": "Advertising",
        "StartDate": "2025-06-15",
        "EndDate": "2025-09-30",
        "Manager": "John Smith",
        "Status": "Active",
        "Priority": "High",
        "Budget": 125000,
        "FileCount": 8
    },
    {
        "ID": "PRJ002",
        "ProjectName": "Website Redesign",
        "ClientName": "TechSolutions Corp",
        "Department": "IT",
        "Category": "Development",
        "StartDate": "2025-07-01",
        "EndDate": "2025-11-15",
        "Manager": "Sarah Johnson",
        "Status": "Active",
        "Priority": "Medium",
        "Budget": 85000,
        "FileCount": 6
    },
    {
        "ID": "PRJ003",
        "ProjectName": "Annual Financial Audit",
        "ClientName": "Internal",
        "Department": "Finance",
        "Category": "Compliance",
        "StartDate": "2025-08-15",
        "EndDate": "2025-09-15",
        "Manager": "Michael Brown",
        "Status": "Planning",
        "Priority": "High",
        "Budget": 45000,
        "FileCount": 5
    },
    {
        "ID": "PRJ004",
        "ProjectName": "Product Launch - XYZ Device",
        "ClientName": "InnovateNow",
        "Department": "Product",
        "Category": "Launch",
        "StartDate": "2025-07-20",
        "EndDate": "2025-10-05",
        "Manager": "Emily Davis",
        "Status": "On Hold",
        "Priority": "Critical",
        "Budget": 250000,
        "FileCount": 10
    },
    {
        "ID": "PRJ005",
        "ProjectName": "Employee Training Program",
        "ClientName": "Internal",
        "Department": "HR",
        "Category": "Training",
        "StartDate": "2025-09-01",
        "EndDate": "2025-11-30",
        "Manager": "Robert Wilson",
        "Status": "Planning",
        "Priority": "Medium",
        "Budget": 35000,
        "FileCount": 4
    },
    {
        "ID": "PRJ006",
        "ProjectName": "Supply Chain Optimization",
        "ClientName": "Global Logistics Partner",
        "Department": "Operations",
        "Category": "Optimization",
        "StartDate": "2025-06-30",
        "EndDate": "2025-12-15",
        "Manager": "Jessica Martinez",
        "Status": "Active",
        "Priority": "High",
        "Budget": 180000,
        "FileCount": 7
    },
    {
        "ID": "PRJ007",
        "ProjectName": "Market Research Study",
        "ClientName": "New Ventures LLC",
        "Department": "Research",
        "Category": "Analysis",
        "StartDate": "2025-08-10",
        "EndDate": "2025-10-10",
        "Manager": "David Thompson",
        "Status": "Active",
        "Priority": "Medium",
        "Budget": 65000,
        "FileCount": 6
    },
    {
        "ID": "PRJ008",
        "ProjectName": "Mobile App Development",
        "ClientName": "HealthTech Innovations",
        "Department": "IT",
        "Category": "Development",
        "StartDate": "2025-07-15",
        "EndDate": "2025-12-01",
        "Manager": "Lisa Anderson",
        "Status": "Active",
        "Priority": "High",
        "Budget": 120000,
        "FileCount": 8
    },
    {
        "ID": "PRJ009",
        "ProjectName": "Office Relocation",
        "ClientName": "Internal",
        "Department": "Facilities",
        "Category": "Infrastructure",
        "StartDate": "2025-09-15",
        "EndDate": "2025-11-15",
        "Manager": "Kevin Miller",
        "Status": "Planning",
        "Priority": "Medium",
        "Budget": 75000,
        "FileCount": 5
    },
    {
        "ID": "PRJ010",
        "ProjectName": "Cybersecurity Upgrade",
        "ClientName": "Internal",
        "Department": "IT Security",
        "Category": "Security",
        "StartDate": "2025-08-01",
        "EndDate": "2025-09-30",
        "Manager": "Olivia Wilson",
        "Status": "Active",
        "Priority": "Critical",
        "Budget": 95000,
        "FileCount": 6
    }
]

# Create Excel file with project data
df = pd.DataFrame(project_data)
df.to_excel(test_dir / "comprehensive_project_database.xlsx", index=False)

print("Created Excel file with project data")

# Define file types and their extensions
file_types = {
    "Document": [".docx", ".doc", ".pdf", ".txt"],
    "Spreadsheet": [".xlsx", ".xls", ".csv"],
    "Presentation": [".pptx", ".ppt"],
    "Image": [".jpg", ".png", ".gif"],
    "Archive": [".zip", ".rar"],
    "Data": [".json", ".xml"]
}

# Document types and their descriptions
document_types = {
    "Report": ["Analysis", "Summary", "Overview", "Metrics", "Results"],
    "Plan": ["Strategy", "Timeline", "Roadmap", "Schedule", "Framework"],
    "Meeting": ["Minutes", "Notes", "Agenda", "Summary"],
    "Proposal": ["Draft", "Final", "Revision"],
    "Contract": ["Agreement", "Terms", "Conditions", "Legal"],
    "Budget": ["Forecast", "Expenses", "Costs", "Financial"],
    "Specs": ["Requirements", "Technical", "Design", "Specifications"]
}

# Date format function


def format_date(date_str):
    date_obj = datetime.datetime.strptime(date_str, "%Y-%m-%d")
    return date_obj.strftime("%Y%m%d")


# Generate test files
file_count = 0
for project in project_data:
    project_name = project["ProjectName"]
    project_id = project["ID"]

    # Create different variations of filenames based on the project
    file_formats = [
        # Standard name formats
        f"{project_name} Documentation",
        f"Project {project_name}",
        f"{project_name} - Status Update",

        # Complex name formats with dates
        f"{format_date(project['StartDate'])} - {project_name}",
        f"{project_name} Plan ({format_date(project['StartDate'])} to {format_date(project['EndDate'])})",

        # Inconsistent formats
        f"{project_name.upper()}",
        f"{project_name.replace(' ', '_')}",

        # Names with special characters
        f"{project_name} (Draft)",
        f"{project_name} - {project['ClientName']}",

        # Names with status or department
        f"{project['Department']} - {project_name}",
        f"{project_name} [{project['Status']}]",

        # Names with manager information
        f"{project_name} - {project['Manager']}",

        # Names with no clear connection
        f"File for {project_id}",

        # Files with cryptic names (to test difficult matching scenarios)
        f"DOC{project_id.replace('PRJ', '')}_{random.randint(1000, 9999)}"
    ]

    # Take a random subset of file formats based on project's FileCount
    num_files = min(project["FileCount"], len(file_formats))
    selected_formats = random.sample(file_formats, num_files)

    for format_idx, filename_base in enumerate(selected_formats):
        # Select random document type and its description
        doc_type = random.choice(list(document_types.keys()))
        doc_description = random.choice(document_types[doc_type])

        # Select file category and extension
        file_category = random.choice(list(file_types.keys()))
        file_extension = random.choice(file_types[file_category])

        # Create full filename
        final_filename = f"{filename_base} - {doc_type} {doc_description}{file_extension}"

        # Sanitize filename (remove characters that are not allowed in filenames)
        final_filename = final_filename.replace(
            ":", "-").replace("/", "-").replace("\\", "-")
        final_filename = ''.join(
            c for c in final_filename if c not in '<>:"/\\|?*')

        # Create file content
        content = f"Test file for {project_name} (ID: {project_id})\n"
        content += f"Manager: {project['Manager']}\n"
        content += f"Client: {project['ClientName']}\n"
        content += f"Status: {project['Status']}\n"
        content += f"Period: {project['StartDate']} to {project['EndDate']}\n"
        content += f"\nThis is a {doc_type} {doc_description} file created for testing the Excel-Based File Renamer application."

        # Write file to disk
        file_path = test_dir / final_filename
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(content)

        file_count += 1

print(f"Created {file_count} test files across {len(project_data)} projects")

# Create some special test cases
special_cases = [
    # Files with no matching entry in Excel
    "Unknown Project Documentation.docx",
    "Miscellaneous Files.pdf",

    # Files with special characters
    "Project (with) {special} [characters].docx",

    # Very long filenames
    "This is an extremely long filename that contains a lot of words and might be difficult to display in some interfaces or file systems depending on their limitations.docx",

    # Files with numbers in the name
    "Project123 - Version 2.5 Documentation.pdf",

    # Files with similar names to test matching accuracy
    "Marketing Campaign 2024.docx",  # Similar to PRJ001 but different year
    "Website Redesign Phase 2.docx",  # Similar to PRJ002

    # Files with dates
    "20250820_Meeting_Notes.docx",
    "Report_2025-08-21.pdf"
]

for special_file in special_cases:
    content = f"This is a special test case file: {special_file}\n"
    content += "Created for testing edge cases in the Excel-Based File Renamer application."

    file_path = test_dir / special_file
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(content)

    file_count += 1

print(f"Added {len(special_cases)} special test case files")
print(f"Total files created: {file_count}")

print("\nTest files and Excel database created successfully!")
