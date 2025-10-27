# Excel Rule Validation System

A powerful Windows application for validating Excel data using natural language rules. Define complex validation rules in plain English and automatically apply them to your Excel data.

## Features

- ✅ **Natural Language Rules**: Define validation rules in plain English
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
3. Enter validation rules in natural language
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

Rules are written in natural language. The system automatically parses them into executable validation logic.

### Rule Structure
```
If [condition] and/or [condition], then [action/message]
```

### Supported Operators

- **Comparison**: greater than, less than, equal to, not equal to
- **Numeric**: >, <, >=, <=, ==, !=
- **Logical**: and, or

### Example Rules

1. **Simple Comparison**:
   ```
   If current is greater than 2, then current validation ok
   ```

2. **Multiple Conditions**:
   ```
   If current is greater than 2 and JB_Property is YES, then JB validation is ok
   ```

3. **Complex Logic**:
   ```
   If Ratio is less than 5 and Status is Active, then validation passed
   ```

4. **Real-world Example**:
   ```
   If Starting_Current is greater than 10 and Rated_Current is less than 3, then current mismatch error
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
2. **rule_parser.py**: Natural language rule parsing
3. **rule_engine.py**: Rule execution and validation logic
4. **gui_app.py**: Windows GUI application
5. **cli_app.py**: Command-line interface

### Workflow

```
Excel File → Excel Reader → Data Frame
                               ↓
Natural Language Rules → Rule Parser → Parsed Rules
                               ↓
                         Rule Engine → Validation Results
```

## Best Practices

1. **Column Names**: Use clear, descriptive column names in your Excel files
2. **Rule Clarity**: Write rules as clearly as possible
3. **Testing**: Test rules with a small dataset first
4. **Validation**: Review validation results before taking action
5. **Backup**: Always keep a backup of your original Excel files

## Advanced Features

### Custom Column Mapping
The system automatically detects column names from your Excel file and matches them case-insensitively in rules.

### Batch Validation
Process multiple Excel files by running the CLI tool in a loop or batch script.

### Export Results
Results can be exported to text files for reporting and auditing.

## Troubleshooting

### Common Issues

1. **"Column not found" error**: Ensure column names in rules match Excel columns
2. **"Failed to parse rule" error**: Check rule syntax and structure
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
- If current > 2A and JB property is YES, then JB validation ok
- If starting current / rated current < 5, then ratio error

### Example 2: Quality Control

**Data**: Product measurements
**Rules**:
- If weight is greater than 100 and quality is A, then quality check passed
- If temperature is less than 20 or humidity is greater than 80, then environment warning

## Future Enhancements

- [ ] Support for more complex rule logic
- [ ] Machine learning for rule suggestions
- [ ] Integration with databases
- [ ] Web-based interface
- [ ] Real-time validation
- [ ] Multi-language support

## Contact

For questions or support, please open an issue on GitHub.