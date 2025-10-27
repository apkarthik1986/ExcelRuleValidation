"""
Tests for Excel Rule Validation System (updated for expression-based rules)
"""
import pandas as pd
import sys
from pathlib import Path

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent))

from excel_reader import ExcelReader
from rule_parser import RuleParser, ConditionType
from rule_engine import RuleEngine


def test_rule_parser():
    """Test rule parsing functionality."""
    print("Testing Rule Parser...")
    
    parser = RuleParser()
    columns = ['Current', 'JB_Property', 'Ratio']
    
    # Test simple rule (expression-based)
    rule_text = "Current>2"
    rule = parser.parse_rule(rule_text, columns, rule_name="test1")
    
    assert rule is not None
    assert len(rule.conditions) > 0
    assert rule.conditions[0].column == 'Current'
    print(f"  ✓ Simple rule parsed: {rule.name}")
    
    # Test complex rule (expression-based)
    rule_text2 = "(Current>2) AND (JB_Property=YES)"
    rule2 = parser.parse_rule(rule_text2, columns, rule_name="test2")
    
    assert rule2 is not None
    assert len(rule2.conditions) == 2
    print(f"  ✓ Complex rule parsed: {rule2.name}")
    
    print("✓ Rule Parser tests passed!\n")


def test_rule_engine():
    """Test rule engine functionality."""
    print("Testing Rule Engine...")
    
    # Create sample data
    data = pd.DataFrame({
        'Current': [2.5, 1.8, 3.2],
        'JB_Property': ['YES', 'NO', 'YES'],
        'Ratio': [5.0, 4.0, 5.5]
    })
    
    # Create rule (expression-based)
    parser = RuleParser()
    rule = parser.parse_rule(
        "(Current>2) AND (JB_Property=YES)",
        data.columns.tolist(),
        rule_name="test_rule"
    )
    
    # Validate
    engine = RuleEngine()
    results = engine.validate(data, [rule])
    
    assert len(results) == 3
    print(f"  ✓ Validated {len(results)} rows")
    
    # Check specific results
    passed = engine.get_passed_validations()
    failed = engine.get_failed_validations()
    
    print(f"  ✓ {len(passed)} passed, {len(failed)} failed")
    print("✓ Rule Engine tests passed!\n")


def test_integration():
    """Test complete workflow."""
    print("Testing Complete Workflow...")
    
    # Create test data
    data = pd.DataFrame({
        'Equipment_ID': ['EQ001', 'EQ002', 'EQ003'],
        'Current': [2.5, 1.8, 3.2],
        'JB_Property': ['YES', 'NO', 'YES'],
        'Ratio': [5.0, 4.0, 5.625]
    })
    
    # Save to Excel
    test_file = Path(__file__).parent / 'test_data.xlsx'
    data.to_excel(test_file, index=False)
    print(f"  ✓ Created test file: {test_file}")
    
    # Load with ExcelReader
    reader = ExcelReader(str(test_file))
    loaded_data = reader.load()
    assert len(loaded_data) == 3
    print(f"  ✓ Loaded {len(loaded_data)} rows")
    
    # Parse rules (expression-based)
    parser = RuleParser()
    rule1 = parser.parse_rule(
        "(Current>2) AND (JB_Property=YES)",
        loaded_data.columns.tolist(),
        rule_name="Rule1"
    )
    rule2 = parser.parse_rule(
        "Ratio>5",
        loaded_data.columns.tolist(),
        rule_name="Rule2"
    )
    print(f"  ✓ Parsed 2 rules")
    
    # Validate
    engine = RuleEngine()
    results = engine.validate(loaded_data, [rule1, rule2])
    
    assert len(results) == 6  # 3 rows × 2 rules
    print(f"  ✓ Generated {len(results)} validation results")
    
    # Generate report
    report = engine.generate_report()
    assert 'VALIDATION REPORT' in report
    print("  ✓ Generated validation report")
    
    # Cleanup
    test_file.unlink()
    print(f"  ✓ Cleaned up test file")
    
    print("✓ Integration tests passed!\n")


def main():
    """Run all tests."""
    print("=" * 80)
    print("Excel Rule Validation System - Test Suite")
    print("=" * 80 + "\n")
    
    try:
        test_rule_parser()
        test_rule_engine()
        test_integration()
        
        print("=" * 80)
        print("ALL TESTS PASSED!")
        print("=" * 80)
        return 0
        
    except Exception as e:
        print(f"\n✗ Test failed with error: {str(e)}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
