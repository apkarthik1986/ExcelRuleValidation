"""
Test quoted string literals in expressions.
"""
import pandas as pd
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from rule_parser import RuleParser
from rule_engine import RuleEngine


def test_quoted_strings():
    """Test that quoted strings are treated as literals, not column references"""
    print("Testing quoted string literals...")
    
    parser = RuleParser()
    columns = ['Status', 'Active', 'Type']  # Note: 'Active' is a column name
    
    # Test with quoted string - should treat "Active" as literal, not column
    rule1 = parser.parse_rule('Status="Active"', columns, rule_name="test1")
    assert rule1.conditions[0].column == 'Status'
    assert rule1.conditions[0].value == '__LITERAL__Active'  # Marked as string literal
    print(f"  ✓ Status=\"Active\" - parsed as string literal")
    
    # Test with unquoted - should treat Active as column reference
    rule2 = parser.parse_rule('Status=Active', columns, rule_name="test2")
    assert rule2.conditions[0].column == 'Status'
    assert rule2.conditions[0].value == 'Active'  # Column reference
    print(f"  ✓ Status=Active - parsed as column reference")
    
    # Test with data
    data = pd.DataFrame({
        'Status': ['Active', 'Inactive', 'Active'],
        'Active': ['Yes', 'No', 'Yes'],
        'Type': ['A', 'B', 'C']
    })
    
    engine = RuleEngine()
    
    # Test quoted literal
    results1 = engine.validate(data, [rule1])
    # Row 0: 'Active' == 'Active' = True
    # Row 1: 'Inactive' == 'Active' = False
    # Row 2: 'Active' == 'Active' = True
    assert results1[0].passed == True
    assert results1[1].passed == False
    assert results1[2].passed == True
    print(f"  ✓ String literal comparison works correctly")
    
    print("✓ Quoted string literals test passed!\n")


def test_backward_compatibility():
    """Test backward compatibility with unquoted YES/NO values"""
    print("Testing backward compatibility...")
    
    parser = RuleParser()
    columns = ['JB_Property', 'Status']
    
    # Unquoted YES/NO should still work for backward compatibility
    rule = parser.parse_rule('JB_Property=YES', columns, rule_name="test_bc")
    assert rule.conditions[0].value == 'YES'
    print(f"  ✓ Unquoted YES/NO values still work")
    
    # Test with data
    data = pd.DataFrame({
        'JB_Property': ['YES', 'NO', 'YES'],
        'Status': ['Active', 'Inactive', 'Active']
    })
    
    engine = RuleEngine()
    results = engine.validate(data, [rule])
    
    assert results[0].passed == True
    assert results[1].passed == False
    assert results[2].passed == True
    print(f"  ✓ Backward compatibility maintained")
    
    print("✓ Backward compatibility test passed!\n")


def main():
    """Run tests."""
    print("=" * 80)
    print("Quoted String Literals Test Suite")
    print("=" * 80 + "\n")
    
    try:
        test_quoted_strings()
        test_backward_compatibility()
        
        print("=" * 80)
        print("ALL QUOTED STRING TESTS PASSED!")
        print("=" * 80)
        return 0
        
    except Exception as e:
        print(f"\n✗ Test failed with error: {str(e)}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
