"""
Tests for expression-based rule validation.
"""
import pandas as pd
import sys
from pathlib import Path

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent))

from excel_reader import ExcelReader
from rule_parser import RuleParser, ConditionType
from rule_engine import RuleEngine


def test_simple_comparison():
    """Test simple comparison: A>B"""
    print("Testing simple comparison (A>B)...")
    
    parser = RuleParser()
    columns = ['A', 'B', 'X', 'G']
    
    # Test A>B
    rule = parser.parse_rule("A>B", columns, rule_name="test1")
    
    assert rule is not None
    assert len(rule.conditions) == 1
    assert rule.conditions[0].column == 'A'
    assert rule.conditions[0].operator == ConditionType.GREATER_THAN
    print(f"  ✓ Parsed: A>B")
    
    # Test with data
    data = pd.DataFrame({
        'A': [5, 2, 10],
        'B': [3, 8, 5],
        'X': [1, 2, 3],
        'G': [1, 1, 1]
    })
    
    engine = RuleEngine()
    results = engine.validate(data, [rule])
    
    # Row 0: 5>3 = True
    # Row 1: 2>8 = False
    # Row 2: 10>5 = True
    assert results[0].passed == True
    assert results[1].passed == False
    assert results[2].passed == True
    print(f"  ✓ Validation correct: 2 passed, 1 failed")
    print("✓ Simple comparison test passed!\n")


def test_combined_expression():
    """Test combined expression: (A>B) AND (X=G)"""
    print("Testing combined expression ((A>B) AND (X=G))...")
    
    parser = RuleParser()
    columns = ['A', 'B', 'X', 'G']
    
    rule = parser.parse_rule("(A>B) AND (X=G)", columns, rule_name="test2")
    
    assert rule is not None
    assert len(rule.conditions) == 2
    assert rule.conditions[0].column == 'A'
    assert rule.conditions[1].column == 'X'
    assert len(rule.logical_ops) == 1
    print(f"  ✓ Parsed: (A>B) AND (X=G)")
    
    # Test with data
    data = pd.DataFrame({
        'A': [5, 2, 10],
        'B': [3, 8, 5],
        'X': [1, 2, 3],
        'G': [1, 1, 1]
    })
    
    engine = RuleEngine()
    results = engine.validate(data, [rule])
    
    # Row 0: (5>3) AND (1=1) = True AND True = True
    # Row 1: (2>8) AND (2=1) = False AND False = False
    # Row 2: (10>5) AND (3=1) = True AND False = False
    assert results[0].passed == True
    assert results[1].passed == False
    assert results[2].passed == False
    print(f"  ✓ Validation correct: 1 passed, 2 failed")
    print("✓ Combined expression test passed!\n")


def test_contains_operator():
    """Test contains operator: A contains "cc_r\""""
    print("Testing contains operator (A contains \"cc_r\")...")
    
    parser = RuleParser()
    columns = ['voltage', 'current']
    
    rule = parser.parse_rule('voltage contains "cc_r"', columns, rule_name="test3")
    
    assert rule is not None
    assert len(rule.conditions) == 1
    assert rule.conditions[0].column == 'voltage'
    assert rule.conditions[0].operator == ConditionType.CONTAINS
    assert rule.conditions[0].value == "cc_r"
    print(f"  ✓ Parsed: voltage contains \"cc_r\"")
    
    # Test with data
    data = pd.DataFrame({
        'voltage': ['cc_r_123', 'abc', 'test_cc_r', 'xyz'],
        'current': [1, 2, 3, 4]
    })
    
    engine = RuleEngine()
    results = engine.validate(data, [rule])
    
    # Row 0: 'cc_r_123' contains 'cc_r' = True
    # Row 1: 'abc' contains 'cc_r' = False
    # Row 2: 'test_cc_r' contains 'cc_r' = True
    # Row 3: 'xyz' contains 'cc_r' = False
    assert results[0].passed == True
    assert results[1].passed == False
    assert results[2].passed == True
    assert results[3].passed == False
    print(f"  ✓ Validation correct: 2 passed, 2 failed")
    print("✓ Contains operator test passed!\n")


def test_multiple_operators():
    """Test multiple comparison operators"""
    print("Testing multiple operators...")
    
    parser = RuleParser()
    columns = ['A', 'B', 'C']
    
    # Test >=
    rule1 = parser.parse_rule("A>=5", columns, rule_name="test_gte")
    assert rule1.conditions[0].operator == ConditionType.GREATER_EQUAL
    print(f"  ✓ Parsed: A>=5")
    
    # Test <=
    rule2 = parser.parse_rule("B<=10", columns, rule_name="test_lte")
    assert rule2.conditions[0].operator == ConditionType.LESS_EQUAL
    print(f"  ✓ Parsed: B<=10")
    
    # Test !=
    rule3 = parser.parse_rule("C!=0", columns, rule_name="test_ne")
    assert rule3.conditions[0].operator == ConditionType.NOT_EQUAL
    print(f"  ✓ Parsed: C!=0")
    
    print("✓ Multiple operators test passed!\n")


def test_or_operator():
    """Test OR operator"""
    print("Testing OR operator...")
    
    parser = RuleParser()
    columns = ['A', 'B']
    
    rule = parser.parse_rule("A>10 OR B<5", columns, rule_name="test_or")
    
    assert rule is not None
    assert len(rule.conditions) == 2
    assert len(rule.logical_ops) == 1
    print(f"  ✓ Parsed: A>10 OR B<5")
    
    # Test with data
    data = pd.DataFrame({
        'A': [15, 5, 12],
        'B': [10, 3, 8]
    })
    
    engine = RuleEngine()
    results = engine.validate(data, [rule])
    
    # Row 0: (15>10) OR (10<5) = True OR False = True
    # Row 1: (5>10) OR (3<5) = False OR True = True
    # Row 2: (12>10) OR (8<5) = True OR False = True
    assert results[0].passed == True
    assert results[1].passed == True
    assert results[2].passed == True
    print(f"  ✓ Validation correct: 3 passed, 0 failed")
    print("✓ OR operator test passed!\n")


def test_column_to_column_comparison():
    """Test column to column comparison"""
    print("Testing column-to-column comparison...")
    
    parser = RuleParser()
    columns = ['Current', 'Voltage', 'Threshold']
    
    rule = parser.parse_rule("Current>Threshold", columns, rule_name="test_col_compare")
    
    assert rule is not None
    assert len(rule.conditions) == 1
    # The value should be the column name 'Threshold'
    assert rule.conditions[0].value == 'Threshold'
    print(f"  ✓ Parsed: Current>Threshold (column comparison)")
    
    # Test with data
    data = pd.DataFrame({
        'Current': [5, 2, 10],
        'Voltage': [120, 110, 125],
        'Threshold': [3, 8, 5]
    })
    
    engine = RuleEngine()
    results = engine.validate(data, [rule])
    
    # Row 0: 5>3 = True
    # Row 1: 2>8 = False
    # Row 2: 10>5 = True
    assert results[0].passed == True
    assert results[1].passed == False
    assert results[2].passed == True
    print(f"  ✓ Validation correct: 2 passed, 1 failed")
    print("✓ Column-to-column comparison test passed!\n")


def test_integration():
    """Test complete workflow with expression-based rules."""
    print("Testing Complete Workflow with Expression Rules...")
    
    # Create test data
    data = pd.DataFrame({
        'Equipment_ID': ['EQ001', 'EQ002', 'EQ003'],
        'Current': [2.5, 1.8, 3.2],
        'Threshold': [2.0, 2.0, 2.0],
        'Status': ['Active', 'Inactive', 'Active']
    })
    
    # Save to Excel
    test_file = Path(__file__).parent / 'test_expression_data.xlsx'
    data.to_excel(test_file, index=False)
    print(f"  ✓ Created test file: {test_file}")
    
    # Load with ExcelReader
    reader = ExcelReader(str(test_file))
    loaded_data = reader.load()
    assert len(loaded_data) == 3
    print(f"  ✓ Loaded {len(loaded_data)} rows")
    
    # Parse expression-based rules
    parser = RuleParser()
    rule1 = parser.parse_rule(
        "Current>Threshold",
        loaded_data.columns.tolist(),
        rule_name="CurrentCheck"
    )
    rule2 = parser.parse_rule(
        'Status="Active"',
        loaded_data.columns.tolist(),
        rule_name="StatusCheck"
    )
    print(f"  ✓ Parsed 2 expression-based rules")
    
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
    
    print("✓ Integration test passed!\n")


def main():
    """Run all tests."""
    print("=" * 80)
    print("Expression-Based Rule Validation - Test Suite")
    print("=" * 80 + "\n")
    
    try:
        test_simple_comparison()
        test_combined_expression()
        test_contains_operator()
        test_multiple_operators()
        test_or_operator()
        test_column_to_column_comparison()
        test_integration()
        
        print("=" * 80)
        print("ALL EXPRESSION-BASED TESTS PASSED!")
        print("=" * 80)
        return 0
        
    except Exception as e:
        print(f"\n✗ Test failed with error: {str(e)}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
