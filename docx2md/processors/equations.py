"""
Equation processing module for docx2md.

This module provides functions for processing and fixing LaTeX equations
in Markdown files.
"""

import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

from docx2md.processors.base import BaseProcessor
from docx2md.utils.file_utils import read_file, write_file
from docx2md.utils.logging_utils import get_logger

logger = get_logger(__name__)


class EquationProcessor(BaseProcessor):
    """
    Processor for LaTeX equations in Markdown.
    """
    
    def __init__(self, config: Optional[Dict[str, Any]] = None):
        """
        Initialize the equation processor.
        
        Args:
            config: Configuration dictionary
        """
        super().__init__(config)
        
        # Get delimiter configuration
        eq_config = self.config.get("equations", {})
        self.inline_delimiters = tuple(eq_config.get("inline_delimiters", ["$", "$"]))
        self.display_delimiters = tuple(eq_config.get("display_delimiters", ["$$", "$$"]))
    
    def process(self, content: str) -> str:
        """
        Process equations in the content.
        
        Args:
            content: Markdown content with equations
            
        Returns:
            Processed content with standardized equation delimiters
        """
        # Fix inline equations
        content = self._fix_inline_equations(content)
        
        # Fix display equations
        content = self._fix_display_equations(content)
        
        return content
    
    def _fix_inline_equations(self, content: str) -> str:
        """
        Fix inline equation delimiters.
        
        Args:
            content: Markdown content
            
        Returns:
            Content with standardized inline equation delimiters
        """
        # Replace \(...\) with inline delimiters
        pattern = r'\\\((.*?)\\\)'
        start, end = self.inline_delimiters
        content = re.sub(pattern, lambda m: f'{start}{m.group(1)}{end}', content, flags=re.DOTALL)
        
        return content
    
    def _fix_display_equations(self, content: str) -> str:
        """
        Fix display equation delimiters.
        
        Args:
            content: Markdown content
            
        Returns:
            Content with standardized display equation delimiters
        """
        # Replace \[...\] with display delimiters
        pattern = r'\\\[(.*?)\\\]'
        start, end = self.display_delimiters
        content = re.sub(pattern, lambda m: f'{start}{m.group(1)}{end}', content, flags=re.DOTALL)
        
        return content


def fix_delimiters(
    input_file: Union[str, Path],
    output_file: Union[str, Path],
    inline_delimiters: Tuple[str, str] = ("$", "$"),
    display_delimiters: Tuple[str, str] = ("$$", "$$"),
) -> Dict[str, Any]:
    """
    Fix LaTeX equation delimiters in a Markdown file.
    
    Args:
        input_file: Path to the input Markdown file
        output_file: Path to the output Markdown file
        inline_delimiters: Tuple of (start, end) delimiters for inline equations
        display_delimiters: Tuple of (start, end) delimiters for display equations
        
    Returns:
        Dictionary with statistics about the fixes
    """
    logger.info(f"Fixing equation delimiters in {input_file}")
    
    # Read the input file
    content = read_file(input_file)
    if content is None:
        raise ValueError(f"Failed to read input file: {input_file}")
    
    # Count original occurrences
    inline_original_count = len(re.findall(r'\\\((.*?)\\\)', content, re.DOTALL))
    display_original_count = len(re.findall(r'\\\[(.*?)\\\]', content, re.DOTALL))
    
    # Fix inline equations
    pattern_inline = r'\\\((.*?)\\\)'
    content = re.sub(pattern_inline, lambda m: f'{inline_delimiters[0]}{m.group(1)}{inline_delimiters[1]}', content, flags=re.DOTALL)
    
    # Fix display equations
    pattern_display = r'\\\[(.*?)\\\]'
    content = re.sub(pattern_display, lambda m: f'{display_delimiters[0]}{m.group(1)}{display_delimiters[1]}', content, flags=re.DOTALL)
    
    # Count fixed occurrences
    # Display blocks first and removed before counting inline ones, because
    # the default "$$" delimiters would otherwise also match as empty "$ $" pairs.
    display_re = re.escape(display_delimiters[0]) + r'(.*?)' + re.escape(display_delimiters[1])
    inline_re = re.escape(inline_delimiters[0]) + r'(.*?)' + re.escape(inline_delimiters[1])
    display_fixed_count = len(re.findall(display_re, content, re.DOTALL))
    inline_fixed_count = len(re.findall(inline_re, re.sub(display_re, '', content, flags=re.DOTALL), re.DOTALL))
    
    # Write the output file
    if not write_file(output_file, content):
        raise ValueError(f"Failed to write output file: {output_file}")
    
    logger.info(f"Fixed {inline_original_count} inline and {display_original_count} display equations")
    
    return {
        "inline_original": inline_original_count,
        "display_original": display_original_count,
        "inline_fixed": inline_fixed_count,
        "display_fixed": display_fixed_count,
    }


def process_equations(
    input_file: Union[str, Path],
    output_file: Union[str, Path],
    config: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """
    Process equations in a Markdown file.
    
    Args:
        input_file: Path to the input Markdown file
        output_file: Path to the output Markdown file
        config: Configuration dictionary
        
    Returns:
        Dictionary with statistics about the processing
    """
    processor = EquationProcessor(config)
    return processor.process_file(input_file, output_file)


def validate_equations(
    input_file: Union[str, Path],
    strict: bool = False,
) -> Dict[str, Any]:
    """
    Validate LaTeX equations in a Markdown file.
    
    Args:
        input_file: Path to the Markdown file to validate
        strict: Whether to use strict validation
        
    Returns:
        Dictionary with validation results
    """
    logger.info(f"Validating equations in {input_file}")
    
    # Read the input file
    content = read_file(input_file)
    if content is None:
        raise ValueError(f"Failed to read input file: {input_file}")
    
    # Find all inline equations
    # Find all display equations first, then remove them so that "$$" is not
    # also counted as inline "$ $" pairs.
    display_pattern = r'\$\$(.*?)\$\$'
    display_equations = re.findall(display_pattern, content, re.DOTALL)
    without_display = re.sub(display_pattern, '', content, flags=re.DOTALL)

    inline_pattern = r'\$(.*?)\$'
    inline_equations = re.findall(inline_pattern, without_display, re.DOTALL)
    
    # Validate equations
    issues = []
    invalid_count = 0
    
    # Check for basic syntax errors in inline equations
    for i, eq in enumerate(inline_equations):
        if not _is_valid_equation(eq, strict):
            invalid_count += 1
            issues.append(f"Invalid inline equation #{i+1}: {eq[:30]}...")
    
    # Check for basic syntax errors in display equations
    for i, eq in enumerate(display_equations):
        if not _is_valid_equation(eq, strict):
            invalid_count += 1
            issues.append(f"Invalid display equation #{i+1}: {eq[:30]}...")
    
    is_valid = invalid_count == 0
    
    logger.info(f"Validation complete: {len(inline_equations)} inline, {len(display_equations)} display, {invalid_count} invalid")
    
    return {
        "is_valid": is_valid,
        "inline_count": len(inline_equations),
        "display_count": len(display_equations),
        "invalid_count": invalid_count,
        "issues": issues,
    }


def _is_valid_equation(equation: str, strict: bool = False) -> bool:
    """
    Check if a LaTeX equation is valid.
    
    Args:
        equation: LaTeX equation string
        strict: Whether to use strict validation
        
    Returns:
        True if the equation is valid, False otherwise
    """
    # Basic validation: check for unbalanced braces
    open_braces = equation.count('{')
    close_braces = equation.count('}')
    
    if open_braces != close_braces:
        return False
    
    # Check for unbalanced environments
    begins = re.findall(r'\\begin\{([^}]+)\}', equation)
    ends = re.findall(r'\\end\{([^}]+)\}', equation)
    
    if len(begins) != len(ends):
        return False
    
    # In strict mode, check that environment names match
    if strict and begins:
        for b, e in zip(begins, ends):
            if b != e:
                return False
    
    # Additional checks could be added here
    
    return True