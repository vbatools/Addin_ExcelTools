**English** | [Русский](README.md)
---

# Addin_ExcelTools

> **Excel VBA Add-in for Professional Excel Automation**
>
> **Author:** VBATools
>
> **Version:** 3.0.24
>
> **License:** [Apache License](LICENSE)

> [!IMPORTANT]
> **VBA Project Password:** `1`
>
> The code is open; the password only protects against accidental changes.

---

## 📑 Table of Contents

* [📋 Description](#-description)
* [🚀 Installation](#-installation)
* [🧰 Tools and Functions List](#-tools-and-functions-list)

---

## 📋 Description

**ExcelTools** is a professional Excel VBA add-in that provides a comprehensive set of tools for automating routine tasks in Excel. It includes 100+ tools: formulas, data manipulation, sheet management, file operations, charts, shapes, and much more.

<p align="center">
  <img src="./img/excel_tools_main_panel.png" alt="Main Control Panel" border="1">
  <br>
  <em>Main Control Panel</em>
</p>

<p align="center">
  <img src="./img/excel_tools_other_panel.png" alt="Other Tools" border="1">
  <br>
  <em>Other Management Tools</em>
</p>

---

## 🚀 Installation

1. Download `Addin_ExcelTools.xlsm` from the repository.
2. Open the file in Excel and allow macro execution.
3. Click the **Install** button.
4. Two tabs will appear on the ribbon: **EXCELToolsMain** and **EXCELToolsOther**.

<p align="center">
  <img src="./img/installation.png" alt="Add-in Installation" border="1">
  <br>
  <em>Add-in Installation</em>
</p>

---

## 🧰 Tools and Functions List

The add-in includes over 100 tools divided into logical categories:

### 📊 Cell Ranges

#### 📥 Data Collection and Loading from Sheets and Workbooks
* Collect data from sheets with the same data structure
* Load data into sheets with the same data structure
* Collect data from workbooks with the same data structure
* Load data into workbooks with the same data structure
* Create workbooks from a template workbook

#### 🔢 Unique Values
* Get a range of unique values

#### 🧩 Merging and Splitting Cells
* Merge cells without data loss
* Merge cells with identical data
* Split cells

#### 🎯 Cross-highlighting Cells
* Highlight the row and column of the active cell

### 🧮 Functions

#### 🔄 Reference Style
* Toggle cell reference style: A1 ↔ R1C1

#### ➗ Calculations
* **Calculation Operations** - performs calculations on a range (round, multiply, add, subtract, divide, etc.)
* **Error Blocking** - wraps a formula in `IFERROR()`
* **Copy Formula Unchanged**
* **Lock Formulas** - cyclically changes reference types in formulas: `$A1 → $A$1 → A$1 → A1`
* **Text Operations** - a set of string transformation tools
* **New Thread** - launches an Excel application instance in a new thread
* **Open Folder** - opens the folder where the active Excel workbook is located
* **Sheet Options**:
  * Show/Hide zeros
  * Show page breaks
  * Data grouping:
    * Totals in rows below data
    * Totals in columns to the right of data
    * Show/Hide outline symbols

#### 📐 Custom Functions
* **Text and String Processing**:
  * `REPLACE_CHARS` - character-by-character replacement
  * `LEFT_TEXT` - returns text to the left of a delimiter
  * `RIGHT_TEXT` - returns text to the right of a delimiter
  * `BETWEEN_TEXT` - returns text between delimiters
  * `FIND_REPLACE` - finds and replaces a substring
  * `SPLIT_STRING` - splits a string by a delimiter character
  * `CONCAT_MULTI` - merges multiple ranges
  * `TEXT_MATCHES_PATTERN` - checks text against a pattern (Like)
  * `TRANSLIT` - transliterates text (4 standards)
  * `REMOVE_CHARS` - removes characters from text
* **Data Extraction from Cells**:
  * `GET_COMMENT` - returns the cell comment text
  * `GET_TEXT` - returns only letters from a cell
  * `GET_NUMBER` - returns only digits from a cell
  * `FORMULA_TEXT` - returns the cell formula as text
* **Fill and Font Color**:
  * `SUM_FILL` - sums by fill color
  * `SUM_FONT` - sums by font color
  * `COUNT_FILL` - counts by fill color
  * `COUNT_FONT` - counts by font color
* **Information Functions**:
  * `BOOK_NAME` - returns the workbook name
  * `SHEET_NAME` - returns the sheet name
  * `USER_NAME` - returns the user name
  * `BOOK_FULL_PATH` - returns the full path to the workbook
* **Number-to-Words**:
  * `NUM_TO_WORDS` - writes a numeric value in words
  * `DATE_TO_WORDS` - writes a date in words
  * `TIME_TO_WORDS` - writes time in words
* **QR Code**:
  * `CREATE_QR` - creates a QR code from a cell value
  * **Create QR Code to File** - creates a QR code as a file from cell data

### 🗂️ Managers

#### 📄 Workbook Sheet Manager
* **Create Table of Contents Sheet** - creates a `Table of Contents` sheet with a list of all sheets and their information
* **Sort Sheets** - sorts sheets alphabetically and by sheet tab color
* **Change Visibility** - changes sheet visibility: hidden, very hidden, and visible
* **Copy Selected Sheets** - copies selected sheets a specified number of times
* **Delete Selected Sheets** - deletes selected sheets
* **Enable and Disable Sheet Protection**
* **Set Zoom** - changes the zoom of selected sheets
* **Export Sheets to Files** - exports sheets to separate files with a specified file type
* **Import Sheets from Workbooks** - inserts sheets from other workbooks into the active workbook
* **Data Grouping** - creates a grouping structure on selected sheets
* **Freeze Panes** - freezes or unfreezes the pane
* **Change Sheet Tab Color**
* **Create New Sheet**
* **Synchronize Sheets** - sets the same zoom and selects the same cell as on the template sheet

#### 📚 Workbook Manager
* **Save Selected** - saves selected workbooks
* **Close Selected Without Saving**

#### 🏷️ Workbook Name Manager
* **Delete Selected Names**
* **Change Name Visibility**

#### 🎨 Workbook Style Manager
* **Delete Selected Styles**

### 🔷 Shapes

#### 🗃️ Shape Dispatcher
* **Copy Selected Shapes** - copies shapes a specified number of times
* **Delete Shapes**
* **Rename Shapes**
* **Change Shape Visibility**
* **Export from Workbook** - exports shapes from the workbook as images in a specified format
* **Import to Workbook** - loads images into the workbook
* **Edit Text in Shape**
* **Append to Text in Shape**
* **Copy Dimensions** - copies the dimensions of selected shapes
* **Paste Dimensions** - applies previously copied dimensions to selected shapes

#### 🔶 Shape Manager
* **Align Shapes** - aligns shapes vertically or horizontally with a specified number of rows or columns and spacing between shapes
* **Copy Range to Shape**
* **Copy to Shape with Range Link**
* **Fit to Cell Size** - adjusts the shape size to match the cell size

#### 📊 Chart Manager
* **Change Data Series Fill** - changes the data series fill of a chart to the fill color of its source data

### 📈 Pivot Tables

* **Create Sheets by Pivot Table Filter**
* **Set Pivot Table Field Format**
* **Clear Pivot Table Cache** of data removed from the pivot table

### 🔌 Add-ins

* **Add-ins** - enables and disables add-ins
* **Add-in Information**

### 🔑 Passwords

* **Remove Passwords from VBA Projects**
* **Remove Passwords from Sheets and Workbook Structure**

### 🧹 Workbook Cleanup

* **File Information** - displays general and custom workbook properties
* **Reset Formats** - removes excessive formats in the workbook
* **Delete Hidden Names in Workbook**
* **Workbook External Links** - creates a sheet with information about all links in the workbook
* **Extract Attached Files from Workbook**

### 📁 Files

* **Create File List** - creates a sheet with a list of files in the selected folder
* **Rename Files** - renames files from a list

### 💬 Comments

* **Comments** - creates a sheet with a list of all comments in the workbook
* **Create Comments** - creates comments from the selected range data
* **Show Comments** - displays all comments on the active workbook sheet
* **Hide Comments** - hides all comments on the active workbook sheet
* **Font Size** - sets the font size for all comments on the active workbook sheet

### 🛠️ Other Tools

* **Create Workbook Backup** - creates backup copies of the workbook, assigning each a version number
* **Group Parameter Solver**
* **Insert Empty Rows/Columns**
* **Export Data Range to JSON or CSV**
