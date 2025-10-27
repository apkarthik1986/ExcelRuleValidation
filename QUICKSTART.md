# Quick Start Guide

## Excel Rule Validation System

This guide will help you get started with the Excel Rule Validation System in 5 minutes.

## Prerequisites

- Windows OS
- Python 3.8 or higher installed
- Your Excel file ready for validation

## Installation (One-Time Setup)

### Method 1: Automatic Setup (Recommended)

1. Double-click `setup.bat`
2. Wait for dependencies to install
3. Done!

### Method 2: Manual Setup

1. Open Command Prompt in this directory
2. Run: `pip install -r requirements.txt`
3. Done!

## Usage

### Using the GUI (Recommended)

1. **Launch the Application**
   - Double-click `start_gui.bat`, OR
   - Run: `python gui_app.py`

2. **Load Your Excel File**
   - Click "Browse..." button
   - Select your Excel file (.xlsx or .xls)
   - Click "Load Data"
   - Preview your data in the table

3. **Add Validation Rules**
   - Type your rule using expression syntax in the rule input box
   - Example: `(Current>2) AND (JB_Property=YES)`
   - Click "Add Rule"
   - Repeat for each rule

4. **Run Validation**
   - Click "Validate Data"
   - View results in the results panel
   - Export results if needed

### Using the Command Line

1. **Basic Usage**
   ```cmd
   python cli_app.py your_data.xlsx --interactive
   ```

2. **Using a Rules File**
   ```cmd
   python cli_app.py your_data.xlsx --rules your_rules.txt
   ```

3. **Export Results**
   ```cmd
   python cli_app.py your_data.xlsx --rules your_rules.txt --output results.txt
   ```

## Example Walkthrough

### Try with Sample Data

1. **Create Sample Files**
   ```cmd
   python create_samples.py
   ```
   This creates:
   - `sample_electrical_data.xlsx` - Sample data
   - `sample_rules.txt` - Sample rules

2. **Test with GUI**
   - Run `start_gui.bat`
   - Load `sample_electrical_data.xlsx`
   - Add rules from `sample_rules.txt` manually, OR
   - Copy rules from the file and paste them one by one

3. **Test with CLI**
   ```cmd
   python cli_app.py sample_electrical_data.xlsx --rules sample_rules.txt
   ```

## Writing Rules

### Basic Rule Structure
```
column operator value
(column operator value) LOGICAL_OP (column operator value)
```

### Examples

**Simple Comparison:**
```
Current>2
```

**Multiple Conditions:**
```
(Current>2) AND (JB_Property=YES)
```

**Column-to-Column Comparison:**
```
Starting_Current>Rated_Current
```

**String Contains:**
```
voltage contains "cc_r"
```

### Supported Operators

- **Comparison**: `>`, `<`, `>=`, `<=`, `=`, `!=`
- **String**: `contains`, `starts_with`, `ends_with`
- **Logical**: `AND`, `OR`

## Tips for Success

1. **Column Names**: Make sure column names in rules match your Excel file
2. **Case Insensitive**: Column matching is case-insensitive
3. **Test Small**: Start with a small subset of data
4. **Clear Rules**: Write rules as clearly as possible
5. **Review Results**: Always review validation results before acting

## Common Use Cases

### Electrical Equipment Validation
```
(Current>2) AND (JB_Property=YES)
Ratio>5
Starting_Current>Rated_Current
```

### Quality Control
```
(Weight>100) AND (Quality=A)
Temperature<20
```

### Inventory Management
```
(Quantity<10) AND (Status=Active)
Price>1000
```

## Troubleshooting

### Issue: "Python is not installed"
- **Solution**: Install Python from https://www.python.org/
- Make sure to check "Add Python to PATH" during installation

### Issue: "Module not found"
- **Solution**: Run `setup.bat` or `pip install -r requirements.txt`

### Issue: "Failed to load Excel file"
- **Solution**: Make sure file path is correct and file is not open in Excel

### Issue: "Column not found in rule"
- **Solution**: Check that column names in rules match Excel file exactly

## Next Steps

1. Read the full [README.md](README.md) for detailed documentation
2. Customize rules for your specific use case
3. Create a rules file for reusable validation
4. Automate validation with the CLI in batch scripts

## Getting Help

- Check example rules in `sample_rules.txt`
- Run sample data with `create_samples.py`
- Review error messages for guidance
- Open an issue on GitHub for support

## Best Practices

1. ✅ Always backup your original Excel files
2. ✅ Test rules with sample data first
3. ✅ Use descriptive column names in Excel
4. ✅ Write clear, simple rules
5. ✅ Export and save validation results
6. ✅ Review failed validations manually

---

**Ready to validate your data?**

Run `start_gui.bat` to get started now!
