import os
import pandas as pd
from excel_reader import ExcelReader
from rule_parser import RuleParser

ROOT = os.path.dirname(os.path.dirname(__file__))
SAMPLE_XLSX = os.path.join(ROOT, 'sample_smoke.xlsx')

# Create a sample Excel where row 0 is blank, row 1 contains headers (including duplicates)
if True:
    # build raw rows: first row blank, second row header, following rows data
    raw = [
        [None, None, None],
        ['A', 'B', 'A'],  # headers on Excel row 2 (index 1)
        [1, 'x', 10],
        [2, 'y', 20],
        [3, 'z', 30]
    ]
    # create a DataFrame representing raw rows and write with header=False so pandas writes exactly
    df_raw = pd.DataFrame(raw)
    df_raw.to_excel(SAMPLE_XLSX, index=False, header=False)
    print(f"Wrote sample Excel to: {SAMPLE_XLSX}")

# Now load it using ExcelReader with header_row=1 (0-based index)
reader = ExcelReader(SAMPLE_XLSX)
try:
    df = reader.load(header_row=1)
    print('\nLoaded DataFrame:')
    print(df.head())
    print('\nColumns (before unique-suffix):', list(df.columns))

    # emulate unique column renaming logic from gui_app
    orig_columns = list(df.columns)
    counts = {}
    new_columns = []
    for col in orig_columns:
        if col not in counts:
            counts[col] = 0
            new_columns.append(col)
        else:
            counts[col] += 1
            new_columns.append(f"{col}{counts[col]}")
    df.columns = new_columns
    print('\nColumns (after unique-suffix):', list(df.columns))

    # Parse a simple rule that references duplicate-named column
    parser = RuleParser()
    columns = df.columns.tolist()
    sample_rule = "A > 1 AND B = 'y'"
    rule = parser.parse_rule(sample_rule, columns, rule_name='SmokeRule1')
    print('\nParsed rule name:', rule.name)
    print('Conditions:')
    for c in rule.conditions:
        print(' -', c.column, c.operator, c.value)
    print('\nSmoke test completed successfully.')
    exit(0)
except Exception as e:
    print('Smoke test failed:', e)
    raise
