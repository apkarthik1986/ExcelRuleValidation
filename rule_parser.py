"""
Rule Parser Module
Parses expression-based rules into executable rule objects.
"""
import re
from typing import Dict, List, Any, Tuple, Optional
from dataclasses import dataclass
from enum import Enum


class ConditionType(Enum):
    """Types of conditions supported in rules."""
    GREATER_THAN = ">"
    LESS_THAN = "<"
    GREATER_EQUAL = ">="
    LESS_EQUAL = "<="
    EQUAL = "="
    NOT_EQUAL = "!="
    CONTAINS = "contains"
    STARTS_WITH = "starts_with"
    ENDS_WITH = "ends_with"


class LogicalOperator(Enum):
    """Logical operators for combining conditions."""
    AND = "AND"
    OR = "OR"


@dataclass
class Condition:
    """Represents a single condition in a rule."""
    column: str
    operator: ConditionType
    value: Any
    
    def __str__(self):
        return f"{self.column} {self.operator.value} {self.value}"


@dataclass
class RuleReference:
    """Represents a reference to another rule by name."""
    rule_name: str
    
    def __str__(self):
        return f"Rule:{self.rule_name}"


@dataclass
class Rule:
    """Represents a complete validation rule."""
    name: str
    conditions: List[Condition]
    logical_ops: List[LogicalOperator]
    action: str
    description: str
    rule_references: List[RuleReference] = None
    
    def __post_init__(self):
        if self.rule_references is None:
            self.rule_references = []
    
    def __str__(self):
        cond_str = ""
        for i, condition in enumerate(self.conditions):
            cond_str += str(condition)
            if i < len(self.logical_ops):
                cond_str += f" {self.logical_ops[i].value} "
        return f"Rule: {self.name}\nConditions: {cond_str}\nAction: {self.action}"


class RuleParser:
    """
    Parses expression-based rules into structured Rule objects.
    
    Supported formats:
    - Simple expressions: A>B, X=G
    - Combined expressions: (A>B) AND (X=G)
    - String operations: A contains "cc_r"
    - Rule references: Rule1 AND Rule2
    """
    
    def __init__(self):
        """Initialize the rule parser."""
        self.rules = []
        self.rules_by_name = {}
        
    def parse_rule(self, expression: str, available_columns: List[str], rule_name: str = None) -> Rule:
        """
        Parse an expression-based rule into a structured Rule object.
        
        Args:
            expression: The rule expression (e.g., "(A>B) AND (X=G)")
            available_columns: List of column names available in the data
            rule_name: Optional name for the rule. If not provided, auto-generated.
            
        Returns:
            Parsed Rule object
        """
        # Clean input
        expression = expression.strip()
        
        # Generate or use provided rule name
        if not rule_name:
            rule_name = self._generate_rule_name(expression)
        
        # Check if this is a rule reference combination (e.g., "Rule1 AND Rule2")
        if self._is_rule_reference_expression(expression):
            return self._parse_rule_references(expression, rule_name)
        
        # Parse as normal expression
        conditions, logical_ops, action = self._parse_expression(expression, available_columns)
        
        rule = Rule(
            name=rule_name,
            conditions=conditions,
            logical_ops=logical_ops,
            action=action,
            description=expression
        )
        
        self.rules.append(rule)
        self.rules_by_name[rule_name] = rule
        return rule
    
    def _is_rule_reference_expression(self, expression: str) -> bool:
        """Check if expression is a rule reference combination."""
        # Pattern for Rule references like "Rule1 AND Rule2"
        pattern = r'^Rule\d+\s+(AND|OR)\s+Rule\d+'
        return bool(re.match(pattern, expression, re.IGNORECASE))
    
    def _parse_rule_references(self, expression: str, rule_name: str) -> Rule:
        """Parse rule reference combinations like 'Rule1 AND Rule2'."""
        # Split by AND/OR
        logical_ops = []
        rule_refs = []
        
        # Find all AND/OR operators
        parts = re.split(r'\s+(AND|OR)\s+', expression, flags=re.IGNORECASE)
        
        for i, part in enumerate(parts):
            part = part.strip()
            if part.upper() in ['AND', 'OR']:
                logical_ops.append(LogicalOperator.AND if part.upper() == 'AND' else LogicalOperator.OR)
            elif part:
                # Extract rule name
                match = re.match(r'Rule(\d+)', part, re.IGNORECASE)
                if match:
                    rule_refs.append(RuleReference(rule_name=part))
        
        rule = Rule(
            name=rule_name,
            conditions=[],
            logical_ops=logical_ops,
            action=f"{rule_name} validation",
            description=expression,
            rule_references=rule_refs
        )
        
        self.rules.append(rule)
        self.rules_by_name[rule_name] = rule
        return rule
    
    def _generate_rule_name(self, expression: str) -> str:
        """Generate a rule name from the expression."""
        # Use first part of expression, sanitized
        name = re.sub(r'[^\w\s]', '', expression)[:30]
        name = "_".join(name.split())
        return name if name else f"rule_{len(self.rules) + 1}"
    
    def _parse_expression(
        self, expression: str, available_columns: List[str]
    ) -> Tuple[List[Condition], List[LogicalOperator], str]:
        """
        Parse an expression into conditions and logical operators.
        
        Args:
            expression: The expression to parse
            available_columns: Available column names
            
        Returns:
            Tuple of (conditions, logical_operators, action)
        """
        conditions = []
        logical_ops = []
        action = "validation ok"
        
        # Split by AND and OR (case-insensitive)
        # First, identify logical operators and their positions
        and_pattern = re.compile(r'\bAND\b', re.IGNORECASE)
        or_pattern = re.compile(r'\bOR\b', re.IGNORECASE)
        
        and_matches = [(m.start(), LogicalOperator.AND) for m in and_pattern.finditer(expression)]
        or_matches = [(m.start(), LogicalOperator.OR) for m in or_pattern.finditer(expression)]
        
        all_ops = sorted(and_matches + or_matches, key=lambda x: x[0])
        logical_ops = [op[1] for op in all_ops]
        
        # Split by AND/OR to get condition parts
        condition_parts = re.split(r'\s+AND\s+|\s+OR\s+', expression, flags=re.IGNORECASE)
        
        for part in condition_parts:
            part = part.strip()
            # Remove surrounding parentheses
            part = part.strip('()')
            
            if part:
                condition = self._parse_single_condition(part, available_columns)
                if condition:
                    conditions.append(condition)
        
        return conditions, logical_ops, action
    
    def _parse_single_condition(
        self, condition_text: str, available_columns: List[str]
    ) -> Optional[Condition]:
        """
        Parse a single condition expression.
        
        Supported formats:
        - A>B
        - A>=5
        - A=G
        - A contains "cc_r"
        
        Args:
            condition_text: Text of a single condition
            available_columns: Available column names
            
        Returns:
            Condition object or None
        """
        condition_text = condition_text.strip()
        
        # Check for 'contains', 'starts_with', 'ends_with' operators
        if 'contains' in condition_text.lower():
            return self._parse_string_operation(condition_text, available_columns, ConditionType.CONTAINS)
        elif 'starts_with' in condition_text.lower():
            return self._parse_string_operation(condition_text, available_columns, ConditionType.STARTS_WITH)
        elif 'ends_with' in condition_text.lower():
            return self._parse_string_operation(condition_text, available_columns, ConditionType.ENDS_WITH)
        
        # Parse comparison operators (>=, <=, !=, >, <, =)
        operators = ['>=', '<=', '!=', '>', '<', '=']
        
        for op_str in operators:
            if op_str in condition_text:
                parts = condition_text.split(op_str, 1)
                if len(parts) == 2:
                    left = parts[0].strip()
                    right = parts[1].strip()
                    
                    # Determine which is column and which is value
                    column, value = self._identify_column_and_value(left, right, available_columns)
                    
                    if column:
                        # Map operator string to ConditionType
                        op_type = self._map_operator(op_str)
                        return Condition(column=column, operator=op_type, value=value)
        
        return None
    
    def _parse_string_operation(
        self, condition_text: str, available_columns: List[str], op_type: ConditionType
    ) -> Optional[Condition]:
        """Parse string operations like 'A contains "cc_r"'."""
        # Pattern: column_name contains "value"
        pattern = r'(\w+)\s+(contains|starts_with|ends_with)\s+["\'](.+?)["\']'
        match = re.match(pattern, condition_text, re.IGNORECASE)
        
        if match:
            column_name = match.group(1)
            value = match.group(3)
            
            # Find actual column name (case-insensitive)
            column = self._find_column(column_name, available_columns)
            if column:
                return Condition(column=column, operator=op_type, value=value)
        
        return None
    
    def _identify_column_and_value(
        self, left: str, right: str, available_columns: List[str]
    ) -> Tuple[Optional[str], Any]:
        """Identify which side is the column and which is the value."""
        # Try left as column
        left_col = self._find_column(left, available_columns)
        if left_col:
            value = self._parse_value(right, available_columns)
            return left_col, value
        
        # Try right as column
        right_col = self._find_column(right, available_columns)
        if right_col:
            value = self._parse_value(left, available_columns)
            return right_col, value
        
        return None, None
    
    def _find_column(self, name: str, available_columns: List[str]) -> Optional[str]:
        """Find column name (case-insensitive match)."""
        name = name.strip()
        for col in available_columns:
            if col.lower() == name.lower() or col == name:
                return col
        return None
    
    def _parse_value(self, value_str: str, available_columns: List[str]) -> Any:
        """Parse a value, which can be a number, string, or column reference."""
        value_str = value_str.strip().strip('"\'')
        
        # Check if it's a column reference
        col = self._find_column(value_str, available_columns)
        if col:
            return col
        
        # Try to parse as number
        try:
            if '.' in value_str:
                return float(value_str)
            else:
                return int(value_str)
        except ValueError:
            pass
        
        # Return as string
        return value_str
    
    def _map_operator(self, op_str: str) -> ConditionType:
        """Map operator string to ConditionType."""
        mapping = {
            '>': ConditionType.GREATER_THAN,
            '<': ConditionType.LESS_THAN,
            '>=': ConditionType.GREATER_EQUAL,
            '<=': ConditionType.LESS_EQUAL,
            '=': ConditionType.EQUAL,
            '!=': ConditionType.NOT_EQUAL,
        }
        return mapping.get(op_str, ConditionType.EQUAL)
    
    def get_rules(self) -> List[Rule]:
        """
        Get all parsed rules.
        
        Returns:
            List of Rule objects
        """
        return self.rules
    
    def get_rule_by_name(self, name: str) -> Optional[Rule]:
        """Get a rule by its name."""
        return self.rules_by_name.get(name)
