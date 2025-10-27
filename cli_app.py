"""
Command-Line Interface for Excel Rule Validation
Provides a CLI for Excel file validation with natural language rules.
"""
import argparse
import sys
from pathlib import Path

from excel_reader import ExcelReader
from rule_parser import RuleParser
from rule_engine import RuleEngine


def main():
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        description='Excel Rule Validation System - Validate Excel data using natural language rules',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  # Validate with interactive mode
  python cli_app.py data.xlsx
  
  # Validate with a rule file
  python cli_app.py data.xlsx --rules rules.txt
  
  # Export results to a file
  python cli_app.py data.xlsx --output results.txt

Rule Examples:
  If current is greater than 2 and has JB property YES, then JB validation is ok
  If starting current to rated current ratio is less than 5, then starting current error
        '''
    )
    
    parser.add_argument(
        'excel_file',
        type=str,
        help='Path to the Excel file to validate'
    )
    
    parser.add_argument(
        '--rules',
        type=str,
        help='Path to a text file containing rules (one per line)'
    )
    
    parser.add_argument(
        '--sheet',
        type=str,
        help='Name of the sheet to load (default: first sheet)'
    )
    
    parser.add_argument(
        '--output',
        type=str,
        help='Path to output file for validation results'
    )
    
    parser.add_argument(
        '--interactive',
        action='store_true',
        help='Run in interactive mode to enter rules manually'
    )
    
    args = parser.parse_args()
    
    # Validate input file exists
    excel_path = Path(args.excel_file)
    if not excel_path.exists():
        print(f"Error: Excel file '{args.excel_file}' not found.")
        sys.exit(1)
    
    # Load Excel data
    print(f"Loading Excel file: {args.excel_file}")
    try:
        reader = ExcelReader(str(excel_path))
        data = reader.load(sheet_name=args.sheet)
        print(f"✓ Loaded {len(data)} rows with {len(data.columns)} columns")
        print(f"Columns: {', '.join(data.columns.tolist())}")
    except Exception as e:
        print(f"Error loading Excel file: {str(e)}")
        sys.exit(1)
    
    # Load or create rules
    rule_parser = RuleParser()
    rules = []
    
    if args.rules:
        # Load rules from file
        rules_path = Path(args.rules)
        if not rules_path.exists():
            print(f"Error: Rules file '{args.rules}' not found.")
            sys.exit(1)
        
        print(f"\nLoading rules from: {args.rules}")
        with open(rules_path, 'r') as f:
            for line_num, line in enumerate(f, 1):
                line = line.strip()
                if line and not line.startswith('#'):
                    try:
                        rule = rule_parser.parse_rule(line, data.columns.tolist())
                        rules.append(rule)
                        print(f"  ✓ Rule {line_num}: {rule.name}")
                    except Exception as e:
                        print(f"  ✗ Rule {line_num} failed: {str(e)}")
    
    elif args.interactive or not args.rules:
        # Interactive mode
        print("\n" + "=" * 80)
        print("INTERACTIVE RULE DEFINITION MODE")
        print("=" * 80)
        print("Enter rules in natural language (or 'done' to finish):")
        print("\nExamples:")
        print("  • If current is greater than 2 and has JB property YES, then JB validation is ok")
        print("  • If ratio is less than 5, then ratio error")
        print()
        
        while True:
            rule_text = input("Enter rule (or 'done'): ").strip()
            if rule_text.lower() == 'done':
                break
            
            if not rule_text:
                continue
            
            try:
                rule = rule_parser.parse_rule(rule_text, data.columns.tolist())
                rules.append(rule)
                print(f"  ✓ Added rule: {rule.name}")
            except Exception as e:
                print(f"  ✗ Failed to parse rule: {str(e)}")
    
    if not rules:
        print("\nNo rules defined. Exiting.")
        sys.exit(0)
    
    # Validate data
    print(f"\n{'=' * 80}")
    print(f"VALIDATING DATA WITH {len(rules)} RULE(S)")
    print("=" * 80)
    
    engine = RuleEngine()
    try:
        results = engine.validate(data, rules)
        
        # Generate report
        report = engine.generate_report()
        print(report)
        
        # Export if requested
        if args.output:
            output_path = Path(args.output)
            with open(output_path, 'w') as f:
                f.write(report)
            print(f"\n✓ Results exported to: {args.output}")
        
        # Summary
        passed = len(engine.get_passed_validations())
        failed = len(engine.get_failed_validations())
        print(f"\nSummary: {passed} passed, {failed} failed")
        
        # Exit with appropriate code
        sys.exit(0 if failed == 0 else 1)
        
    except Exception as e:
        print(f"\nError during validation: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
