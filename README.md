# Excel Rule Validation System

A powerful Windows application for validating Excel data using expression-based rules. Define complex validation rules using column names and operators to automatically validate your Excel data.

## Features

- ✅ **Expression-Based Rules**: Define validation rules using column names and operators with intuitive syntax
- ✅ **Excel File Support**: Read and validate .xlsx and .xls files with auto-detection of header rows
- ✅ **Dynamic Rule Engine**: Rules are parsed and applied automatically with support for complex conditions
- ✅ **Modern Windows GUI**: User-friendly graphical interface with tabbed layout and Data-first design
- ✅ **Context Menu Support**: Right-click to copy example rules and expressions
- ✅ **Rule Management**: Enable/disable rules, save/load rule sets, and organize validations
- ✅ **Command-Line Interface**: CLI for automation and scripting needs
- ✅ **Comprehensive Reports**: Detailed validation results with export capabilities

## Installation

### System Requirements

#### Minimum Requirements
- **Operating System**: Windows 7 or higher
- **Python**: Version 3.10 or higher
- **RAM**: 4 GB minimum
- **Disk Space**: 500 MB for installation and dependencies

#### Recommended
- **Operating System**: Windows 10/11 
- **Python**: Version 3.11 (recommended for best performance)
- **RAM**: 8 GB or more
- **Disk Space**: 1 GB

### Quick Setup (Windows)

1. Install Python 3.11 from [python.org](https://www.python.org/downloads/)
   - ✅ **IMPORTANT**: Check "Add Python to PATH" during installation

2. Clone this repository:
```bash
git clone https://github.com/apkarthik1986/ExcelRuleValidation.git
cd ExcelRuleValidation
```

3. Run the automatic setup:
```bash
setup.bat
```

This will create a Python virtual environment and install all dependencies.

### Manual Setup

If you prefer manual control, you can set up the environment yourself:

```powershell
# Create a Python 3.11 virtual environment
py -3.11 -m venv .venv311
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
.\.venv311\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

### Dependencies

The application requires the following main packages:
- `pandas>=2.0.0`: For Excel data handling
- `openpyxl>=3.1.2`: For Excel file support
- `pytest>=8.4.2`: For testing

Optional ML/NLP features require:
- `spacy>=3.7.0`: For text analysis
- `transformers>=4.30.0`: For advanced NLP
- `torch>=2.0.0`: For ML operations
- `sentence-transformers>=2.2.0`: For semantic analysis
- `numpy>=1.24.0`: For numerical operations

These are included in `requirements.txt` but can be installed separately in an isolated environment if needed.
```

Note: `torch` can be large and has platform-specific wheels. If you want a
CPU-only wheel, use the PyTorch CPU index or conda as shown above.


## Usage

### GUI Application (Recommended for Windows)

Launch the graphical interface:
```bash
start_gui.bat
# Or manually:
python gui_app.py
```

#### Modern Interface Features
- **Tabbed Layout**: Organized sections for Data, Rules, Validation, and Examples
- **Data-First Design**: Data tab is selected by default for quick data loading
- **Smart Header Detection**: Auto-detect header rows or specify manually
- **Interactive Examples**: Copy examples via right-click menu
- **Rule Management**: Enable/disable rules with double-click
- **Rule Organization**: Save and load rule sets for reuse

#### Workflow
1. **Data Loading**:
   - Click "Browse..." to select Excel file
   - Set header row (default: 1) or use auto-detect
   - Click "Load Data" to preview the Excel data

2. **Rule Creation**:
   - Switch to Rules tab
   - Enter validation rules using expression syntax
   - Double-click examples to use them as templates
   - Click "Add Rule" to add each rule
   - Enable/disable rules with double-click

3. **Validation**:
   - Switch to Validation tab
   - Click "Validate Data" to run checks
   - View detailed results in the output area
   - Export results for documentation

4. **Examples & Help**:
   - Browse example rules in Examples tab
   - Right-click to copy examples
   - Use column selector for quick insertion
   - Reference rule snippets for complex validations

### Command-Line Interface

#### Interactive Mode
```bash
python cli_app.py sample_electrical_data.xlsx --interactive
```

#### Using Rules File
```bash
python cli_app.py sample_electrical_data.xlsx --rules sample_rules.txt
```

#### With Output File
```bash
python cli_app.py sample_electrical_data.xlsx --rules sample_rules.txt --output results.txt
```

## Rule Syntax

Rules are written using expression-based syntax with column names and operators.

### Rule Structure
```
column operator value
(column operator value) LOGICAL_OP (column operator value)
```

### Supported Operators

- **Comparison**: >, <, >=, <=, =, !=
- **String**: contains, starts_with, ends_with
- **Logical**: AND, OR

### Validation Rules Guide

#### Basic Operators
| Category | Operators | Example |
|----------|-----------|---------|
| Comparison | `>`, `<`, `>=`, `<=`, `=`, `!=` | `Current>2`, `Status!=Error` |
| String | `contains`, `startswith`, `endswith` | `Name contains "test"` |
| Logical | `AND`, `OR`, `NOT` | `(A>2) AND (B<5)` |
| Special | `between`, `in`, `regex_match` | `Value between 1 and 10` |

#### Example Rules

1. **Simple Comparisons**:
   ```
   Current>2
   Status=Active
   Temperature<=100
   ```

2. **String Operations**:
   ```
   Name contains "test"
   ID startswith "EQ"
   Code endswith "_VALID"
   ```

3. **Multiple Conditions**:
   ```
   (Current>2) AND (JB_Property=YES)
   (A>10) OR (B<5)
   NOT (Status=Error)
   ```

4. **Column Comparisons**:
   ```
   Starting_Current>Rated_Current
   Temperature<Max_Limit
   Actual_Value>=Expected_Value
   ```

5. **Complex Logic**:
   ```
   ((A>10 AND B<5) OR (C='x' AND (D>0 OR E<0)))
   (Status=Active) AND (NOT(Error_Count>0))
   (Value between 1 and 10) OR (Code in ('A','B','C'))
   ```

6. **Advanced Functions**:
   ```
   regex_match(Email, '^[^@]+@example\.com$')
   trim(Name)='John'
   LOWER(Category)='approved'
   ```

#### Best Practices

1. **Rule Organization**:
   - Group related rules together
   - Use meaningful rule names
   - Enable/disable rules as needed

2. **Validation Strategy**:
   - Start with basic rules
   - Add complexity incrementally
   - Test with sample data
   - Save working rule sets

3. **Rule Maintenance**:
   - Document complex rules
   - Use the Examples tab for reference
   - Keep a backup of proven rule sets
   - Review and update regularly

## Sample Data

Create sample files for testing:
```bash
python create_samples.py
```

This creates:
- `sample_electrical_data.xlsx` - Sample electrical equipment data
- `sample_rules.txt` - Sample validation rules

### Sample Data Structure

| Equipment_ID | Current | JB_Property | Starting_Current | Rated_Current | Ratio |
|-------------|---------|-------------|------------------|---------------|-------|
| EQ001       | 2.5     | YES         | 12.5             | 2.5           | 5.0   |
| EQ002       | 1.8     | NO          | 8.0              | 2.0           | 4.0   |
| EQ003       | 3.2     | YES         | 18.0             | 3.2           | 5.625 |

## Architecture

### Components

1. **excel_reader.py**: Excel file reading and data extraction
2. **rule_parser.py**: Expression-based rule parsing
3. **rule_engine.py**: Rule execution and validation logic
4. **gui_app.py**: Windows GUI application
5. **cli_app.py**: Command-line interface

### Workflow

```
Excel File → Excel Reader → Data Frame
                               ↓
Expression Rules → Rule Parser → Parsed Rules
                               ↓
                         Rule Engine → Validation Results
```

## Best Practices

1. **Column Names**: Use clear, descriptive column names in your Excel files
2. **Rule Clarity**: Write rules using explicit column names and operators
3. **Testing**: Test rules with a small dataset first
4. **Validation**: Review validation results before taking action
5. **Backup**: Always keep a backup of your original Excel files

## Advanced Features

### Column-to-Column Comparison
The system supports comparing values between columns:
```
Current>Threshold
Starting_Current>Rated_Current
```

### String Operations
Check if column values contain specific strings:
```
Status contains "Active"
Equipment_ID starts_with "EQ"
```

### Batch Validation
Process multiple Excel files by running the CLI tool in a loop or batch script.

### Export Results
Results can be exported to text files for reporting and auditing.

## Troubleshooting

### Common Issues

1. **"Column not found" error**: Ensure column names in rules match Excel columns exactly
2. **"Failed to parse rule" error**: Check rule syntax - use expression format (e.g., A>B, not "A is greater than B")
3. **Import errors**: Ensure all dependencies are installed: `pip install -r requirements.txt`

### Getting Help

- Check the examples in `sample_rules.txt`
- Run with sample data: `python create_samples.py`
- Review error messages for specific guidance

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues for bugs and feature requests.

## License

This project is open source and available under the MIT License.

## Examples

### Example 1: Electrical Equipment Validation

**Data**: Electrical equipment with current ratings
**Rules**:
- `(Current>2) AND (JB_Property=YES)`
- `Ratio>5`
- `Starting_Current>Rated_Current`

### Example 2: Quality Control

**Data**: Product measurements
**Rules**:
- `(Weight>100) AND (Quality=A)`
- `(Temperature<20) OR (Humidity>80)`

## Future Enhancements

- [ ] Support for more complex expressions with parentheses
- [ ] Rule templates library
- [ ] Integration with databases
- [ ] Web-based interface
- [ ] Real-time validation
- [ ] Multi-language support

## Contact

For questions or support, please open an issue on GitHub.