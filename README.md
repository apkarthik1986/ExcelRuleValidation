# Excel Rule Validation System

A powerful Windows application for validating Excel data using expression-based rules. Define complex validation rules using column names and operators to automatically validate your Excel data.

## Features

- ✅ **Expression-Based Rules**: Define validation rules using column names and operators
- ✅ **Excel File Support**: Read and validate .xlsx and .xls files
- ✅ **Dynamic Rule Engine**: Rules are parsed and applied automatically
- ✅ **Windows GUI**: User-friendly graphical interface using tkinter
- ✅ **Command-Line Interface**: CLI for automation and scripting
- ✅ **Comprehensive Reports**: Detailed validation results and exports

## Installation

### Prerequisites

- Python 3.8 or higher
- Windows OS (for GUI application)

### Setup

1. Clone this repository:
```bash
git clone https://github.com/apkarthik1986/ExcelRuleValidation.git
cd ExcelRuleValidation
```

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### GUI Application (Recommended for Windows)

Launch the graphical interface:
```bash
python gui_app.py
```

**Steps:**
1. Click "Browse..." to select your Excel file
2. Click "Load Data" to load the Excel data
3. Enter validation rules using expression syntax
4. Click "Add Rule" to add each rule
5. Click "Validate Data" to run validation
6. View results and export if needed

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

### Example Rules

1. **Simple Comparison**:
   ```
   Current>2
   ```

2. **Multiple Conditions with AND**:
   ```
   (Current>2) AND (JB_Property=YES)
   ```

3. **Column-to-Column Comparison**:
   ```
   Starting_Current>Rated_Current
   ```

4. **String Contains Check**:
   ```
   voltage contains "cc_r"
   ```

5. **OR Logic**:
   ```
   (A>10) OR (B<5)
   ```

6. **Complex Expression**:
   ```
   (Ratio>5) AND (Status=Active)
   ```

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