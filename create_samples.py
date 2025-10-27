"""
Create sample Excel files for testing the validation system.
"""
import pandas as pd
import openpyxl
from pathlib import Path


def create_sample_electrical_data():
    """Create sample electrical equipment data."""
    data = {
        'Equipment_ID': ['EQ001', 'EQ002', 'EQ003', 'EQ004', 'EQ005'],
        'Current': [2.5, 1.8, 3.2, 2.1, 1.5],
        'JB_Property': ['YES', 'NO', 'YES', 'YES', 'NO'],
        'Starting_Current': [12.5, 8.0, 18.0, 9.5, 7.0],
        'Rated_Current': [2.5, 2.0, 3.2, 2.1, 1.5],
        'Ratio': [5.0, 4.0, 5.625, 4.52, 4.67],
        'Status': ['Active', 'Active', 'Active', 'Active', 'Active']
    }
    
    df = pd.DataFrame(data)
    output_path = Path(__file__).parent / 'sample_electrical_data.xlsx'
    df.to_excel(output_path, index=False, sheet_name='Equipment')
    print(f"Created sample file: {output_path}")
    return df


def create_sample_rules_file():
    """Create sample rules file."""
    rules = [
        "# Sample validation rules for electrical equipment data",
        "# Rules can be written in natural language",
        "",
        "If current is greater than 2 and JB_Property is YES, then JB validation is ok",
        "If Ratio is greater than 5, then Ratio exceeds limit",
        "If Starting_Current is less than 10 and Rated_Current is greater than 2, then Current mismatch",
    ]
    
    output_path = Path(__file__).parent / 'sample_rules.txt'
    with open(output_path, 'w') as f:
        f.write('\n'.join(rules))
    print(f"Created sample rules file: {output_path}")


if __name__ == "__main__":
    print("Creating sample files for Excel Rule Validation System...")
    df = create_sample_electrical_data()
    print("\nSample data:")
    print(df)
    print()
    create_sample_rules_file()
    print("\nSample files created successfully!")
