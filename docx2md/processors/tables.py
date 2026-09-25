"""
Table processing module for docx2md.

This module provides functions for processing tables in Markdown files.
"""

import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

from docx2md.processors.base import BaseProcessor
from docx2md.utils.file_utils import read_file, write_file
from docx2md.utils.logging_utils import get_logger

logger = get_logger(__name__)


class TableProcessor(BaseProcessor):
    """
    Processor for tables in Markdown.
    """
    
    def __init__(self, config: Optional[Dict[str, Any]] = None):
        """
        Initialize the table processor.
        
        Args:
            config: Configuration dictionary
        """
        super().__init__(config)
        
        # Get table configuration
        table_config = self.config.get("tables", {})
        self.table_format = table_config.get("format", "pipe")
        self.header_style = table_config.get("header_style", "bold")
    
    def process(self, content: str) -> str:
        """
        Process tables in the content.
        
        Args:
            content: Markdown content with tables
            
        Returns:
            Processed content with formatted tables
        """
        # Find all tables
        if self.table_format == "pipe":
            # Process pipe tables
            content = self._process_pipe_tables(content)
        elif self.table_format == "grid":
            # Process grid tables
            content = self._process_grid_tables(content)
        
        return content
    
    def _process_pipe_tables(self, content: str) -> str:
        """
        Process pipe-style Markdown tables.
        
        Args:
            content: Markdown content
            
        Returns:
            Content with processed pipe tables
        """
        # Regular expression to find pipe tables
        # A pipe table has a header row, a separator row, and zero or more data rows
        table_pattern = r'(\|[^\n]+\|\n\|[-:| ]+\|\n(?:\|[^\n]+\|\n)*)'
        
        def format_table(match):
            table = match.group(1)
            lines = table.strip().split('\n')
            
            # Format header row if needed
            if self.header_style == "bold" and len(lines) > 0:
                header = lines[0]
                cells = header.split('|')
                bold_cells = []
                
                for cell in cells:
                    if cell.strip():
                        bold_cells.append(f" **{cell.strip()}** ")
                    else:
                        bold_cells.append(cell)
                
                lines[0] = '|'.join(bold_cells)
            
            return '\n'.join(lines) + '\n'
        
        # Replace tables
        return re.sub(table_pattern, format_table, content)
    
    def _process_grid_tables(self, content: str) -> str:
        """
        Process grid-style Markdown tables.
        
        Args:
            content: Markdown content
            
        Returns:
            Content with processed grid tables
        """
        # Grid tables are more complex and not as common in Markdown
        # This is a placeholder for future implementation
        return content


def process_tables(
    input_file: Union[str, Path],
    output_file: Union[str, Path],
    table_format: str = "pipe",
    header_style: str = "bold",
) -> Dict[str, Any]:
    """
    Process tables in a Markdown file.
    
    Args:
        input_file: Path to the input Markdown file
        output_file: Path to the output Markdown file
        table_format: Table format ('pipe', 'grid', or 'simple')
        header_style: Header style ('bold', 'none')
        
    Returns:
        Dictionary with statistics about the processing
    """
    logger.info(f"Processing tables in {input_file}")
    
    # Read the input file
    content = read_file(input_file)
    if content is None:
        raise ValueError(f"Failed to read input file: {input_file}")
    
    # Count tables before processing
    table_count = 0
    
    # Count pipe tables
    pipe_tables = re.findall(r'\|[^\n]+\|\n\|[-:| ]+\|\n(?:\|[^\n]+\|\n)*', content)
    table_count += len(pipe_tables)
    
    # Process based on format
    if table_format == "pipe":
        # Process pipe tables
        processed_content = _process_pipe_tables(content, header_style)
    elif table_format == "grid":
        # Process grid tables (placeholder)
        processed_content = content
    else:
        # Simple format (no processing)
        processed_content = content
    
    # Write the output file
    if not write_file(output_file, processed_content):
        raise ValueError(f"Failed to write output file: {output_file}")
    
    logger.info(f"Processed {table_count} tables")
    
    return {
        "tables_processed": table_count,
        "table_format": table_format,
    }


def _process_pipe_tables(content: str, header_style: str) -> str:
    """
    Process pipe-style Markdown tables.
    
    Args:
        content: Markdown content
        header_style: Header style ('bold', 'none')
        
    Returns:
        Content with processed pipe tables
    """
    # Regular expression to find pipe tables. The separator must be dashes
    # only in each cell (|---|:--|); a grid table's inner lines (a row above
    # a rule of dash groups, "|  ---- ----  |") must not match, or bolding
    # its "header" would break the grid's alignment.
    table_pattern = (
        r'(?:^|(?<=\n))'
        r'(\|[^\n]+\|\n\|(?:[ ]*:?-+:?[ ]*\|)+\n(?:\|[^\n]+\|\n)*)'
    )
    
    def format_table(match):
        table = match.group(1)
        lines = table.strip().split('\n')
        
        # Format header row if needed
        if header_style == "bold" and len(lines) > 0:
            header = lines[0]
            cells = header.split('|')
            bold_cells = []
            
            for cell in cells:
                if cell.strip():
                    bold_cells.append(f" **{cell.strip()}** ")
                else:
                    bold_cells.append(cell)
            
            lines[0] = '|'.join(bold_cells)
        
        return '\n'.join(lines) + '\n'
    
    # Replace tables
    return re.sub(table_pattern, format_table, content)


def convert_html_tables_to_markdown(content: str) -> str:
    """
    Convert HTML tables to Markdown format.
    
    Args:
        content: Content with HTML tables
        
    Returns:
        Content with HTML tables converted to Markdown
    """
    # This is a simplified implementation that handles basic HTML tables
    # A more robust implementation would use a proper HTML parser
    
    # Pattern to match HTML tables
    table_pattern = r'<table[^>]*>(.*?)</table>'
    
    def convert_table(match):
        html_table = match.group(1)
        
        # Extract rows
        rows = re.findall(r'<tr[^>]*>(.*?)</tr>', html_table, re.DOTALL)
        
        if not rows:
            return match.group(0)
        
        markdown_rows = []
        separator_row = None
        
        for i, row in enumerate(rows):
            # Extract cells
            header_cells = re.findall(r'<th[^>]*>(.*?)</th>', row, re.DOTALL)
            data_cells = re.findall(r'<td[^>]*>(.*?)</td>', row, re.DOTALL)
            
            cells = header_cells or data_cells
            
            if not cells:
                continue
            
            # Clean cell content (remove HTML tags, normalize whitespace)
            clean_cells = []
            for cell in cells:
                # Remove HTML tags
                cell = re.sub(r'<[^>]+>', '', cell)
                # Normalize whitespace
                cell = re.sub(r'\s+', ' ', cell).strip()
                clean_cells.append(cell)
            
            # Create Markdown row
            markdown_row = '| ' + ' | '.join(clean_cells) + ' |'
            markdown_rows.append(markdown_row)
            
            # Create separator row after the first row
            if i == 0:
                separator_row = '| ' + ' | '.join(['---'] * len(cells)) + ' |'
        
        # Combine rows with separator
        if separator_row:
            markdown_rows.insert(1, separator_row)
        
        return '\n'.join(markdown_rows)
    
    # Replace HTML tables with Markdown tables
    return re.sub(table_pattern, convert_table, content, flags=re.DOTALL)