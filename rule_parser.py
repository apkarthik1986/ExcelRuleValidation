"""
Rule Parser Module
Converts natural language rules into executable rule objects.
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
    EQUAL = "=="
    NOT_EQUAL = "!="
    CONTAINS = "contains"
    STARTS_WITH = "starts_with"
    ENDS_WITH = "ends_with"


class LogicalOperator(Enum):
    """Logical operators for combining conditions."""
    AND = "and"
    OR = "or"


@dataclass
class Condition:
    """Represents a single condition in a rule."""
    column: str
    operator: ConditionType
    value: Any
    
    def __str__(self):
        return f"{self.column} {self.operator.value} {self.value}"


@dataclass
class Rule:
    """Represents a complete validation rule."""
    name: str
    conditions: List[Condition]
    logical_ops: List[LogicalOperator]
    action: str
    description: str
    
    def __str__(self):
        cond_str = ""
        for i, condition in enumerate(self.conditions):
            cond_str += str(condition)
            if i < len(self.logical_ops):
                cond_str += f" {self.logical_ops[i].value} "
        return f"Rule: {self.name}\nConditions: {cond_str}\nAction: {self.action}"


class RuleParser:
    """
    Parses natural language rules into structured Rule objects.
    """
    
    # Keywords mapping for condition types
    OPERATOR_KEYWORDS = {
        "greater than": ConditionType.GREATER_THAN,
        "more than": ConditionType.GREATER_THAN,
        "above": ConditionType.GREATER_THAN,
        "exceeds": ConditionType.GREATER_THAN,
        "less than": ConditionType.LESS_THAN,
        "below": ConditionType.LESS_THAN,
        "under": ConditionType.LESS_THAN,
        "equal to": ConditionType.EQUAL,
        "equals": ConditionType.EQUAL,
        "is": ConditionType.EQUAL,
        "has": ConditionType.EQUAL,
        "not equal to": ConditionType.NOT_EQUAL,
        "not equals": ConditionType.NOT_EQUAL,
        "greater than or equal": ConditionType.GREATER_EQUAL,
        "at least": ConditionType.GREATER_EQUAL,
        "less than or equal": ConditionType.LESS_EQUAL,
        "at most": ConditionType.LESS_EQUAL,
    }
    
    def __init__(self):
        """Initialize the rule parser."""
        self.rules = []
        
    def parse_rule(self, natural_language_rule: str, available_columns: List[str]) -> Rule:
        """
        Parse a natural language rule into a structured Rule object.
        
        Args:
            natural_language_rule: The rule in natural language
            available_columns: List of column names available in the data
            
        Returns:
            Parsed Rule object
        """
        # Clean and normalize the input
        rule_text = natural_language_rule.lower().strip()
        
        # Extract rule name (first part before conditions)
        rule_name = self._generate_rule_name(rule_text)
        
        # Parse conditions and actions
        conditions, logical_ops, action = self._extract_conditions_and_action(
            rule_text, available_columns
        )
        
        rule = Rule(
            name=rule_name,
            conditions=conditions,
            logical_ops=logical_ops,
            action=action,
            description=natural_language_rule
        )
        
        self.rules.append(rule)
        return rule
    
    def _generate_rule_name(self, rule_text: str) -> str:
        """Generate a rule name from the rule text."""
        words = rule_text.split()[:5]  # Take first 5 words
        return "_".join(words).replace(",", "").replace(".", "")
    
    def _extract_conditions_and_action(
        self, rule_text: str, available_columns: List[str]
    ) -> Tuple[List[Condition], List[LogicalOperator], str]:
        """
        Extract conditions, logical operators, and action from rule text.
        
        Args:
            rule_text: The rule text
            available_columns: Available column names
            
        Returns:
            Tuple of (conditions, logical_operators, action)
        """
        conditions = []
        logical_ops = []
        action = "validation_ok"
        
        # Split by 'then' to separate conditions from action
        parts = rule_text.split("then")
        condition_text = parts[0].strip()
        
        if len(parts) > 1:
            action = parts[1].strip()
        
        # Handle 'if' keyword
        if condition_text.startswith("if "):
            condition_text = condition_text[3:].strip()
        
        # Split by 'and' and 'or'
        # First, identify logical operators
        and_positions = [m.start() for m in re.finditer(r'\band\b', condition_text)]
        or_positions = [m.start() for m in re.finditer(r'\bor\b', condition_text)]
        
        # Combine and sort positions
        all_ops = []
        for pos in and_positions:
            all_ops.append((pos, LogicalOperator.AND))
        for pos in or_positions:
            all_ops.append((pos, LogicalOperator.OR))
        all_ops.sort(key=lambda x: x[0])
        
        # Split by logical operators
        condition_parts = re.split(r'\band\b|\bor\b', condition_text)
        
        for part in condition_parts:
            part = part.strip()
            if part:
                condition = self._parse_single_condition(part, available_columns)
                if condition:
                    conditions.append(condition)
        
        # Extract logical operators
        logical_ops = [op[1] for op in all_ops]
        
        return conditions, logical_ops, action
    
    def _parse_single_condition(
        self, condition_text: str, available_columns: List[str]
    ) -> Optional[Condition]:
        """
        Parse a single condition.
        
        Args:
            condition_text: Text of a single condition
            available_columns: Available column names
            
        Returns:
            Condition object or None
        """
        # Normalize column names (case-insensitive matching)
        column_map = {col.lower(): col for col in available_columns}
        
        # Try to find column name in the condition
        found_column = None
        for col_lower, col_original in column_map.items():
            if col_lower in condition_text:
                found_column = col_original
                break
        
        if not found_column:
            # Try to extract from pattern like "column_name operator value"
            words = condition_text.split()
            if words and words[0] in column_map:
                found_column = column_map[words[0]]
        
        if not found_column:
            return None
        
        # Find operator
        found_operator = None
        for keyword, op_type in self.OPERATOR_KEYWORDS.items():
            if keyword in condition_text:
                found_operator = op_type
                break
        
        if not found_operator:
            # Default to equality if no operator found
            found_operator = ConditionType.EQUAL
        
        # Extract value
        value = self._extract_value(condition_text, found_column, found_operator)
        
        return Condition(column=found_column, operator=found_operator, value=value)
    
    def _extract_value(self, condition_text: str, column: str, operator: ConditionType) -> Any:
        """
        Extract the value from a condition.
        
        Args:
            condition_text: The condition text
            column: Column name
            operator: Operator type
            
        Returns:
            Extracted value
        """
        # Remove column name and operator keywords
        text = condition_text.lower()
        text = text.replace(column.lower(), "").strip()
        
        for keyword in self.OPERATOR_KEYWORDS.keys():
            text = text.replace(keyword, "").strip()
        
        # Try to extract number
        numbers = re.findall(r'-?\d+\.?\d*', text)
        if numbers:
            value_str = numbers[0]
            if '.' in value_str:
                return float(value_str)
            else:
                return int(value_str)
        
        # Try to extract YES/NO
        if 'yes' in text:
            return 'YES'
        if 'no' in text:
            return 'NO'
        
        # Return the cleaned text
        text = text.strip()
        if text:
            return text
        
        return None
    
    def get_rules(self) -> List[Rule]:
        """
        Get all parsed rules.
        
        Returns:
            List of Rule objects
        """
        return self.rules
