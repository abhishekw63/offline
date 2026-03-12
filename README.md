# GT Mass Dump Automation

A Python-based desktop utility that automates the extraction of **BC Code** and **Order Quantity** from multiple Sales Order Excel files and generates a consolidated dump file.

This tool removes the need for repetitive manual copying from Excel sheets and significantly speeds up operational workflows.

---

# Problem

Operations teams often receive **30–40 Sales Order Excel files daily**.

To prepare a consolidated dump, the manual process requires:

1. Opening each Excel file
2. Locating **BC Code** and **Order Qty**
3. Copying item numbers and quantities
4. Combining them into a single sheet

This process is repetitive, time-consuming, and prone to human error.

---

# Solution

**GT Mass Dump Automation** processes multiple Sales Order files automatically and produces a consolidated dump within seconds.

The tool:

* Reads multiple Excel files
* Extracts **BC Code** and **Order Qty**
* Cleans and validates data
* Reconstructs **SO number format**
* Generates a ready-to-use dump file

---

# Features

* Simple **desktop UI**
* Select multiple Excel files at once
* Automatic header detection
* Ignores summary rows
* Cleans numeric formats (`1,000 → 1000`)
* Converts filenames to original SO format
* Automatic output folder creation
* Date-based dump file naming
* Logging for processing transparency

---

# Example Output

| SO Number   | Item No | Qty |
| ----------- | ------- | --- |
| SO/GTM/5954 | 200173  | 120 |
| SO/GTM/5954 | 200165  | 30  |
| SO/GTM/5955 | 201249  | 240 |

---

# Output File

Generated files are stored in the **output** folder.

Example:

```
output/
    gt_mass_dump_12032026.xlsx
```

File naming pattern:

```
gt_mass_dump_DDMMYYYY.xlsx
```

---

# Folder Structure

```
project/
│
├── gt_mass_automation.py
├── requirements.txt
├── README.md
├── DOCUMENTATION.md
│
├── output/
│   └── gt_mass_dump_12032026.xlsx
│
└── .venv/
```

---

# Installation

## 1. Clone Repository

```
git clone <repository-url>
cd project
```

---

## 2. Create Virtual Environment

```
python -m venv .venv
```

Activate environment:

Windows

```
.venv\Scripts\activate
```

---

## 3. Install Dependencies

```
pip install -r requirements.txt
```

Required libraries:

```
pandas
openpyxl
```

---

# Running the Application

Execute the script:

```
python gt_mass_automation.py
```

A desktop window will appear.

---

# Application Workflow

1. Click **Select Excel Files**
2. Choose multiple Sales Order files
3. Click **Generate Dump**
4. Output file is created inside the **output** folder

---

# Excel Template Requirements

Each Sales Order file must contain these columns:

| Column    | Description       |
| --------- | ----------------- |
| BC Code   | Product Item Code |
| Order Qty | Ordered quantity  |

The script automatically detects the header row and extracts the required data.

---

# Data Processing Rules

The automation performs the following cleaning steps:

### BC Code Handling

Excel values such as

```
200453.0
```

are converted to

```
200453
```

---

### Quantity Cleaning

The script converts values like:

```
1,000 → 1000
- → 0
```

Rows with **quantity ≤ 0** are ignored.

---

# Sales Order Format

Excel filenames typically appear as:

```
SOGTM5985.xlsx
```

The script converts them back to the original format:

```
SO/GTM/5985
```

---

# Technologies Used

| Technology | Purpose                   |
| ---------- | ------------------------- |
| Python     | Core language             |
| pandas     | Excel data processing     |
| openpyxl   | Excel reading and writing |
| tkinter    | Desktop UI                |

---

# Logging Example

During execution, logs display processing status:

```
Reading SOGTM5985.xlsx
Reading SOGTM5986.xlsx
17 rows extracted
```

---

# Future Enhancements

Planned improvements include:

* Progress bar during processing
* Drag-and-drop file upload
* Parallel file processing
* Automatic ERP upload
* Integration with **Dynamics 365**
* Packaging as a standalone executable

---

# License

Internal automation tool for operational efficiency.

---

# Author

Developed to automate repetitive Excel workflows and improve operational productivity.
