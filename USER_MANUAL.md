# Excel Rule Validation System - User Manual

## Table of Contents
1. [Introduction](#introduction)
2. [System Requirements](#system-requirements)
3. [Installation](#installation)
4. [Using the GUI Application](#using-the-gui-application)
5. [Using the CLI Application](#using-the-cli-application)
6. [Writing Validation Rules](#writing-validation-rules)
7. [Understanding Results](#understanding-results)
8. [Advanced Topics](#advanced-topics)
9. [Troubleshooting](#troubleshooting)
10. [FAQ](#faq)

## Introduction

The Excel Rule Validation System is a powerful tool that allows you to validate Excel data using natural language rules. Instead of writing complex code or formulas, you simply describe what you want to check in plain English.

### Key Features
- ✅ Natural language rule definition
- ✅ Support for Excel .xlsx and .xls files
- ✅ Windows-optimized GUI application
- ✅ Command-line interface for automation
- ✅ Comprehensive validation reports
- ✅ Export capabilities

### Use Cases
- Electrical equipment validation
- Quality control checks
- Data integrity verification
- Business rule enforcement
- Compliance checking

## System Requirements

### Minimum Requirements
- **Operating System**: Windows 7 or higher
- **Python**: Version 3.8 or higher
- **RAM**: 4 GB minimum
- **Disk Space**: 500 MB for installation and dependencies

### Recommended
- **Operating System**: Windows 10/11
- **Python**: Version 3.10 or higher
- **RAM**: 8 GB or more
- **Disk Space**: 1 GB

## Installation

### Step 1: Install Python

1. Download Python from https://www.python.org/downloads/
2. Run the installer
3. ✅ **IMPORTANT**: Check "Add Python to PATH"
4. Click "Install Now"
5. Verify installation:
   ```cmd
   python --version
   ```

### Step 2: Install Excel Rule Validation System

#### Option A: Automatic Setup (Recommended)
1. Extract the downloaded ZIP file
2. Open the folder
3. Double-click `setup.bat`
4. Wait for installation to complete

#### Option B: Manual Setup
1. Open Command Prompt in the project folder
2. Run:
   ```cmd
   pip install -r requirements.txt
   ```

## Using the GUI Application

### Starting the Application

**Method 1: Batch File**
- Double-click `start_gui.bat`

**Method 2: Command Prompt**
```cmd
python gui_app.py
```

### Workflow

#### 1. Select and Load Excel File

1. Click the **"Browse..."** button
2. Navigate to your Excel file
3. Select the file and click **"Open"**
4. Click **"Load Data"** to load the file
5. Your data will appear in the preview table

**Tips:**
- Ensure Excel file is closed before loading
- The first row should contain column headers
- Supported formats: .xlsx, .xls

#### 2. Define Validation Rules

1. Look at the rule input section
2. Read the example rules provided
3. Type your rule in the text box
4. Click **"Add Rule"**
5. Repeat for additional rules

**Example Rules:**
```
If current is greater than 2 and JB_Property is YES, then JB validation is ok
If Ratio is less than 5, then ratio error
If Status is Active, then equipment is operational
```

#### 3. Review Active Rules

- All added rules appear in the "Active Rules" list
- To remove a rule:
  1. Click on the rule in the list
  2. Click **"Remove Selected"**

#### 4. Validate Data

1. Click **"Validate Data"** button
2. Wait for validation to complete
3. Review results in the results panel

#### 5. Review and Export Results

**Reading Results:**
- Total validations performed
- Number passed and failed
- Detailed failure information
- Row-by-row breakdown

**Exporting:**
1. Click **"Export Results"**
2. Choose save location
3. Enter filename
4. Click **"Save"**

### GUI Layout

```
┌─────────────────────────────────────────────────┐
│  Excel Rule Validation System                   │
├─────────────────────────────────────────────────┤
│  1. Select Excel File                           │
│     [File Path............] [Browse] [Load]     │
├─────────────────────────────────────────────────┤
│  2. Define Validation Rules                     │
│     Rule: [........................] [Add Rule] │
│     Active Rules:                               │
│     ┌────────────────────────────────────┐      │
│     │ Rule 1                             │      │
│     │ Rule 2                             │      │
│     └────────────────────────────────────┘      │
├─────────────────────────────────────────────────┤
│  3. Data Preview                                │
│     [Table showing loaded data]                 │
├─────────────────────────────────────────────────┤
│  4. Validation Results                          │
│     [Results display area]                      │
├─────────────────────────────────────────────────┤
│  [Validate] [Clear Results] [Export Results]    │
└─────────────────────────────────────────────────┘
```

## Using the CLI Application

### Basic Commands

**Help:**
```cmd
python cli_app.py --help
```

**Interactive Mode:**
```cmd
python cli_app.py data.xlsx --interactive
```

**Using Rules File:**
```cmd
python cli_app.py data.xlsx --rules rules.txt
```

**Specify Sheet:**
```cmd
python cli_app.py data.xlsx --sheet "Sheet1" --rules rules.txt
```

**Export Results:**
```cmd
python cli_app.py data.xlsx --rules rules.txt --output results.txt
```

### Creating Rules Files

Create a text file (e.g., `rules.txt`) with one rule per line:

```text
# My Validation Rules
# Lines starting with # are comments

If current is greater than 2 and JB_Property is YES, then JB validation is ok
If Ratio is greater than 5, then Ratio exceeds limit
If Starting_Current is less than 10, then starting current too low
```

### Interactive Mode Workflow

1. Run: `python cli_app.py data.xlsx --interactive`
2. System loads your Excel file
3. You'll be prompted to enter rules one by one
4. Type each rule and press Enter
5. Type "done" when finished
6. Validation runs automatically
7. Results are displayed

## Writing Validation Rules

### Rule Syntax

**Basic Structure:**
```
If [condition], then [action/message]
```

**Multiple Conditions:**
```
If [condition1] and [condition2], then [action]
If [condition1] or [condition2], then [action]
```

### Condition Operators

| Operator | Keywords | Example |
|----------|----------|---------|
| Greater than | greater than, more than, above, exceeds | `current is greater than 2` |
| Less than | less than, below, under | `ratio is less than 5` |
| Equal to | equal to, equals, is, has | `status is Active` |
| Not equal | not equal to, not equals | `type not equals Manual` |
| Greater/Equal | greater than or equal, at least | `count is at least 10` |
| Less/Equal | less than or equal, at most | `value is at most 100` |

### Logical Operators

- **AND**: Both conditions must be true
  ```
  If current is greater than 2 and JB_Property is YES, then validation ok
  ```

- **OR**: At least one condition must be true
  ```
  If status is Error or status is Failed, then check required
  ```

### Column References

- Column names are **case-insensitive**
- Use the exact column name from your Excel file
- Spaces are preserved

**Example:**
If your Excel has a column named "Starting Current", use:
```
If Starting Current is greater than 10, then...
```

### Value Types

**Numeric Values:**
```
If current is greater than 2.5, then...
If count equals 10, then...
```

**Text Values:**
```
If status is Active, then...
If JB_Property is YES, then...
```

### Best Practices

1. ✅ **Be Specific**: Clearly state what you're checking
   ```
   Good: If current is greater than 2 and JB_Property is YES, then JB validation ok
   Avoid: If current ok and property ok, then ok
   ```

2. ✅ **Use Meaningful Messages**: Make actions/messages descriptive
   ```
   Good: then JB validation passed successfully
   Avoid: then ok
   ```

3. ✅ **One Rule Per Line**: Keep rules simple and focused
   ```
   Good: Two separate rules for two different checks
   Avoid: Complex nested conditions in one rule
   ```

4. ✅ **Test Incrementally**: Add and test rules one at a time

## Understanding Results

### Result Format

```
================================================================================
VALIDATION REPORT
================================================================================
Total validations: 15
Passed: 10
Failed: 5
================================================================================

FAILED VALIDATIONS:
--------------------------------------------------------------------------------

Row 2: Rule 'if_current_is_greater_than' not satisfied
Rule: if_current_is_greater_than
Row data: {'Equipment_ID': 'EQ002', 'Current': 1.8, ...}
```

### Interpreting Results

**Total Validations:**
- Number of rows × Number of rules
- Example: 5 rows with 3 rules = 15 validations

**Passed:**
- Rules that were satisfied for specific rows

**Failed:**
- Rules that were NOT satisfied
- These are the items that need attention

**Row Data:**
- Shows the complete data for failed rows
- Helps understand why validation failed

### Taking Action

1. **Review Failed Validations**: Look at each failed item
2. **Check Row Data**: Understand the values that caused failure
3. **Verify Rules**: Ensure rules are correctly defined
4. **Fix Data or Rules**: Either correct the data or adjust rules
5. **Re-validate**: Run validation again after changes

## Advanced Topics

### Batch Processing

Create a batch script to validate multiple files:

```batch
@echo off
for %%f in (*.xlsx) do (
    echo Processing %%f
    python cli_app.py "%%f" --rules rules.txt --output "%%~nf_results.txt"
)
```

### Scheduled Validation

Use Windows Task Scheduler:
1. Open Task Scheduler
2. Create Basic Task
3. Set trigger (daily, weekly, etc.)
4. Action: Start a program
5. Program: `python`
6. Arguments: `cli_app.py data.xlsx --rules rules.txt --output results.txt`
7. Start in: Path to project folder

### Custom Rule Functions

For advanced users who want to extend functionality, modify `rule_parser.py`:
1. Add new operators to `ConditionType`
2. Add parsing logic in `OPERATOR_KEYWORDS`
3. Implement evaluation in `rule_engine.py`

## Troubleshooting

### Common Issues and Solutions

**Issue: "Python is not installed or not in PATH"**
- **Cause**: Python not installed or PATH not set
- **Solution**: Reinstall Python and check "Add Python to PATH"

**Issue: "Failed to load Excel file"**
- **Cause**: File locked, corrupted, or wrong path
- **Solution**: 
  - Close Excel file
  - Check file path
  - Verify file is not corrupted

**Issue: "Column not found"**
- **Cause**: Column name in rule doesn't match Excel
- **Solution**: Check exact column name (case-insensitive but must match)

**Issue: "Failed to parse rule"**
- **Cause**: Rule syntax error
- **Solution**: Review rule syntax, check examples

**Issue: "Module not found"**
- **Cause**: Dependencies not installed
- **Solution**: Run `setup.bat` or `pip install -r requirements.txt`

### Getting Detailed Errors

Run with Python to see full error messages:
```cmd
python gui_app.py
```
or
```cmd
python cli_app.py data.xlsx --rules rules.txt
```

## FAQ

**Q: Can I use this on Mac or Linux?**
A: The CLI works on Mac/Linux. The GUI is Windows-optimized but may work with modifications.

**Q: What Excel versions are supported?**
A: Both .xlsx (Excel 2007+) and .xls (Excel 97-2003) formats.

**Q: Can I validate multiple sheets?**
A: Yes, use `--sheet "SheetName"` in CLI or load each sheet separately in GUI.

**Q: Are formulas evaluated?**
A: No, only the calculated values are read.

**Q: Can I save rules for reuse?**
A: Yes! Create a text file with rules and use `--rules` in CLI.

**Q: How many rules can I add?**
A: No practical limit, but performance may vary with very large rule sets.

**Q: Can I use this for production?**
A: Yes, but always review results before taking action.

**Q: Is my data sent anywhere?**
A: No, all processing is local on your machine.

---

For additional help, open an issue on GitHub or refer to README.md.
