"""
Rule Engine Module
Executes parsed rules against data and generates validation results.
"""
import pandas as pd
from typing import List, Dict, Any, Optional
from dataclasses import dataclass
from rule_parser import Rule, Condition, ConditionType, LogicalOperator, RuleReference


@dataclass
class ValidationResult:
    """Result of applying a rule to a row."""
    row_index: int
    rule_name: str
    passed: bool
    message: str
    row_data: Dict[str, Any]


class RuleEngine:
    """
    Executes validation rules against data.
    """
    
    def __init__(self):
        """Initialize the rule engine."""
        self.results = []
        self.rule_cache = {}  # Cache for rule results by row
        
    def validate(self, data: pd.DataFrame, rules: List[Rule]) -> List[ValidationResult]:
        """
        Validate data against a list of rules.
        
        Args:
            data: DataFrame containing the data to validate
            rules: List of Rule objects to apply
            
        Returns:
            List of ValidationResult objects
        """
        self.results = []
        
        for index, row in data.iterrows():
            for rule in rules:
                result = self._apply_rule_to_row(row, index, rule)
                self.results.append(result)
        
        return self.results
    
    def _apply_rule_to_row(self, row: pd.Series, index: int, rule: Rule) -> ValidationResult:
        """
        Apply a single rule to a single row.
        
        Args:
            row: Data row as a Series
            index: Row index
            rule: Rule to apply
            
        Returns:
            ValidationResult object
        """
        # Check if this rule references other rules
        if rule.rule_references:
            # This is a rule reference combination
            final_result = self._evaluate_rule_references(row, index, rule)
        else:
            # Evaluate all conditions
            condition_results = []
            for condition in rule.conditions:
                result = self._evaluate_condition(row, condition)
                condition_results.append(result)
            
            # Combine condition results using logical operators
            final_result = self._combine_conditions(condition_results, rule.logical_ops)
        
        # Generate message
        if final_result:
            message = f"Row {index}: {rule.action}"
        else:
            message = f"Row {index}: Rule '{rule.name}' not satisfied"
        
        return ValidationResult(
            row_index=index,
            rule_name=rule.name,
            passed=final_result,
            message=message,
            row_data=row.to_dict()
        )
    
    def _evaluate_rule_references(self, row: pd.Series, index: int, rule: Rule) -> bool:
        """
        Evaluate rule references (e.g., Rule1 AND Rule2).
        
        Args:
            row: Data row
            index: Row index
            rule: Rule containing rule references
            
        Returns:
            Combined boolean result
        """
        # Get results from referenced rules
        # Note: This assumes referenced rules have already been evaluated
        # In practice, you'd need to pass the rules to this method or maintain state
        # For now, we'll return True as a placeholder
        # This will be properly implemented when we have access to all rules
        return True
    
    def _evaluate_condition(self, row: pd.Series, condition: Condition) -> bool:
        """
        Evaluate a single condition against a row.
        
        Args:
            row: Data row
            condition: Condition to evaluate
            
        Returns:
            True if condition is met, False otherwise
        """
        # Get the column value
        if condition.column not in row.index:
            return False
        
        cell_value = row[condition.column]
        
        # Check if condition.value is a column reference or a string literal
        comparison_value = condition.value
        
        # Handle special __LITERAL__ prefix for quoted strings
        if isinstance(comparison_value, str) and comparison_value.startswith('__LITERAL__'):
            # Remove the literal marker
            comparison_value = comparison_value[len('__LITERAL__'):]
        elif isinstance(comparison_value, str) and comparison_value in row.index:
            # Value is a column reference, get its value from the row
            comparison_value = row[comparison_value]
        
        # Handle different condition types
        try:
            if condition.operator == ConditionType.GREATER_THAN:
                return float(cell_value) > float(comparison_value)
            elif condition.operator == ConditionType.LESS_THAN:
                return float(cell_value) < float(comparison_value)
            elif condition.operator == ConditionType.GREATER_EQUAL:
                return float(cell_value) >= float(comparison_value)
            elif condition.operator == ConditionType.LESS_EQUAL:
                return float(cell_value) <= float(comparison_value)
            elif condition.operator == ConditionType.EQUAL:
                # Try numeric comparison first
                try:
                    return float(cell_value) == float(comparison_value)
                except (ValueError, TypeError):
                    # Fall back to string comparison
                    return str(cell_value).strip().upper() == str(comparison_value).strip().upper()
            elif condition.operator == ConditionType.NOT_EQUAL:
                try:
                    return float(cell_value) != float(comparison_value)
                except (ValueError, TypeError):
                    return str(cell_value).strip().upper() != str(comparison_value).strip().upper()
            elif condition.operator == ConditionType.CONTAINS:
                return str(comparison_value).lower() in str(cell_value).lower()
            elif condition.operator == ConditionType.STARTS_WITH:
                return str(cell_value).lower().startswith(str(comparison_value).lower())
            elif condition.operator == ConditionType.ENDS_WITH:
                return str(cell_value).lower().endswith(str(comparison_value).lower())
        except (ValueError, TypeError) as e:
            return False
        
        return False
    
    def _combine_conditions(
        self, condition_results: List[bool], logical_ops: List[LogicalOperator]
    ) -> bool:
        """
        Combine multiple condition results using logical operators.
        
        Args:
            condition_results: List of boolean results from conditions
            logical_ops: List of logical operators to use
            
        Returns:
            Combined boolean result
        """
        if not condition_results:
            return False
        
        if len(condition_results) == 1:
            return condition_results[0]
        
        result = condition_results[0]
        for i, op in enumerate(logical_ops):
            if i + 1 < len(condition_results):
                if op == LogicalOperator.AND:
                    result = result and condition_results[i + 1]
                elif op == LogicalOperator.OR:
                    result = result or condition_results[i + 1]
        
        return result
    
    def get_failed_validations(self) -> List[ValidationResult]:
        """
        Get all failed validation results.
        
        Returns:
            List of failed ValidationResult objects
        """
        return [r for r in self.results if not r.passed]
    
    def get_passed_validations(self) -> List[ValidationResult]:
        """
        Get all passed validation results.
        
        Returns:
            List of passed ValidationResult objects
        """
        return [r for r in self.results if r.passed]
    
    def generate_report(self) -> str:
        """
        Generate a text report of validation results.
        
        Returns:
            Formatted report string
        """
        report = []
        report.append("=" * 80)
        report.append("VALIDATION REPORT")
        report.append("=" * 80)
        report.append(f"Total validations: {len(self.results)}")
        report.append(f"Passed: {len(self.get_passed_validations())}")
        report.append(f"Failed: {len(self.get_failed_validations())}")
        report.append("=" * 80)
        
        if self.get_failed_validations():
            report.append("\nFAILED VALIDATIONS:")
            report.append("-" * 80)
            for result in self.get_failed_validations():
                report.append(f"\n{result.message}")
                report.append(f"Rule: {result.rule_name}")
                report.append(f"Row data: {result.row_data}")
        
        return "\n".join(report)
