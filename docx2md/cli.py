"""
Command-line interface for docx2md.

This module provides a robust CLI with git-like subcommands for converting
Microsoft Word documents to Markdown+LaTeX.
"""

import os
import sys
from pathlib import Path
from typing import List, Optional

import click
from rich.console import Console
from rich.panel import Panel

from docx2md import __version__
from docx2md.config import load_config
from docx2md.converter import convert_docx_to_markdown, batch_convert
from docx2md.processors.equations import fix_delimiters
from docx2md.utils.logging_utils import setup_logger

# Initialize console for rich output
console = Console()
logger = setup_logger()


@click.group(context_settings={"help_option_names": ["-h", "--help"]})
@click.version_option(__version__, "-v", "--version", prog_name="docx2md")
def main() -> None:
    """
    docx2md: Convert Microsoft Word documents with complex math to Markdown+LaTeX.
    
    This tool provides commands for converting Word documents to Markdown,
    fixing LaTeX delimiters, validating output, and more.
    """
    pass


@main.command()
@click.argument("input_file", type=click.Path(exists=True, dir_okay=False))
@click.argument("output_file", type=click.Path(), required=False)
@click.option(
    "--config", "-c", type=click.Path(exists=True), help="Path to configuration file"
)
@click.option(
    "--log-level",
    type=click.Choice(["debug", "info", "warning", "error"], case_sensitive=False),
    default="info",
    help="Set the logging level",
)
def convert(
    input_file: str, output_file: Optional[str], config: Optional[str], log_level: str
) -> None:
    """
    Convert a Word document to Markdown+LaTeX.
    
    INPUT_FILE: Path to the Word document (.docx) to convert
    OUTPUT_FILE: Path for the output Markdown file (optional, defaults to same name with .md extension)
    """
    logger.setLevel(log_level.upper())
    
    # Load configuration if provided
    config_data = load_config(config) if config else {}
    
    # Determine output file path if not provided
    if not output_file:
        output_file = os.path.splitext(input_file)[0] + ".md"
    
    try:
        console.print(f"Converting [bold]{input_file}[/bold] to [bold]{output_file}[/bold]...")
        result = convert_docx_to_markdown(input_file, output_file, config_data)
        console.print(Panel.fit(
            f"✅ Conversion successful!\n\n"
            f"Output file: [bold]{output_file}[/bold]\n"
            f"Images extracted: {result.get('images_count', 0)}\n"
            f"Equations converted: {result.get('equations_count', 0)}\n"
            f"Tables converted: {result.get('tables_count', 0)}",
            title="Conversion Complete",
            border_style="green",
        ))
    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {str(e)}")
        logger.error(f"Conversion failed: {str(e)}", exc_info=True)
        sys.exit(1)


@main.command()
@click.argument("input_dir", type=click.Path(exists=True, file_okay=False))
@click.argument("output_dir", type=click.Path(file_okay=False), required=False)
@click.option(
    "--config", "-c", type=click.Path(exists=True), help="Path to configuration file"
)
@click.option(
    "--recursive/--no-recursive",
    default=False,
    help="Recursively process subdirectories",
)
@click.option(
    "--parallel/--no-parallel",
    default=True,
    help="Process files in parallel",
)
def batch(
    input_dir: str,
    output_dir: Optional[str],
    config: Optional[str],
    recursive: bool,
    parallel: bool,
) -> None:
    """
    Convert multiple Word documents to Markdown+LaTeX.
    
    INPUT_DIR: Directory containing Word documents (.docx) to convert
    OUTPUT_DIR: Directory for output Markdown files (optional, defaults to same as input)
    """
    # Load configuration if provided
    config_data = load_config(config) if config else {}
    
    # Determine output directory if not provided
    if not output_dir:
        output_dir = input_dir
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    try:
        console.print(f"Batch converting documents from [bold]{input_dir}[/bold] to [bold]{output_dir}[/bold]...")
        result = batch_convert(
            input_dir, output_dir, config_data, recursive=recursive, parallel=parallel
        )
        console.print(Panel.fit(
            f"✅ Batch conversion successful!\n\n"
            f"Files processed: {result.get('files_processed', 0)}\n"
            f"Files succeeded: {result.get('files_succeeded', 0)}\n"
            f"Files failed: {result.get('files_failed', 0)}",
            title="Batch Conversion Complete",
            border_style="green",
        ))
    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {str(e)}")
        logger.error(f"Batch conversion failed: {str(e)}", exc_info=True)
        sys.exit(1)


@main.command("fix-delimiters")
@click.argument("input_file", type=click.Path(exists=True, dir_okay=False))
@click.argument("output_file", type=click.Path(), required=False)
@click.option(
    "--inline-delimiters",
    type=str,
    default="$,$",
    help="Inline math delimiters (start,end)",
)
@click.option(
    "--display-delimiters",
    type=str,
    default="$$,$$",
    help="Display math delimiters (start,end)",
)
def fix_delimiters_cmd(
    input_file: str,
    output_file: Optional[str],
    inline_delimiters: str,
    display_delimiters: str,
) -> None:
    """
    Fix LaTeX delimiters in a Markdown file.
    
    INPUT_FILE: Path to the Markdown file to process
    OUTPUT_FILE: Path for the output file (optional, defaults to overwriting input)
    """
    # Parse delimiter options
    inline_start, inline_end = inline_delimiters.split(",")
    display_start, display_end = display_delimiters.split(",")
    
    # Determine output file path if not provided
    if not output_file:
        output_file = input_file
    
    try:
        console.print(f"Fixing LaTeX delimiters in [bold]{input_file}[/bold]...")
        result = fix_delimiters(
            input_file,
            output_file,
            inline_delimiters=(inline_start, inline_end),
            display_delimiters=(display_start, display_end),
        )
        console.print(Panel.fit(
            f"✅ Delimiter fixing successful!\n\n"
            f"Output file: [bold]{output_file}[/bold]\n"
            f"Inline equations fixed: {result.get('inline_fixed', 0)}\n"
            f"Display equations fixed: {result.get('display_fixed', 0)}",
            title="Delimiter Fixing Complete",
            border_style="green",
        ))
    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {str(e)}")
        logger.error(f"Delimiter fixing failed: {str(e)}", exc_info=True)
        sys.exit(1)


@main.command()
@click.argument("input_file", type=click.Path(exists=True, dir_okay=False))
@click.option(
    "--strict/--no-strict",
    default=False,
    help="Strict validation mode",
)
def validate(input_file: str, strict: bool) -> None:
    """
    Validate a Markdown file for LaTeX equation correctness.
    
    INPUT_FILE: Path to the Markdown file to validate
    """
    from docx2md.processors.equations import validate_equations
    
    try:
        console.print(f"Validating LaTeX equations in [bold]{input_file}[/bold]...")
        result = validate_equations(input_file, strict=strict)
        
        if result.get("is_valid", False):
            console.print(Panel.fit(
                f"✅ Validation successful!\n\n"
                f"Inline equations: {result.get('inline_count', 0)}\n"
                f"Display equations: {result.get('display_count', 0)}\n"
                f"All equations are valid.",
                title="Validation Complete",
                border_style="green",
            ))
        else:
            console.print(Panel.fit(
                f"❌ Validation failed!\n\n"
                f"Inline equations: {result.get('inline_count', 0)}\n"
                f"Display equations: {result.get('display_count', 0)}\n"
                f"Invalid equations: {result.get('invalid_count', 0)}\n\n"
                f"First few issues:\n" + "\n".join(result.get("issues", [])[:5]),
                title="Validation Failed",
                border_style="red",
            ))
            sys.exit(1)
    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {str(e)}")
        logger.error(f"Validation failed: {str(e)}", exc_info=True)
        sys.exit(1)


@main.command()
@click.argument("input_docx", type=click.Path(exists=True, dir_okay=False))
@click.argument("output_md", type=click.Path(exists=True, dir_okay=False))
def verify(input_docx: str, output_md: str) -> None:
    """
    Check that a Markdown file holds everything its Word document does.

    Compares equations, equation numbers and the punctuation before them,
    images, figure captions, tables, headings, index entries, footnotes and
    running text, and reports anything lost or changed in conversion.

    INPUT_DOCX: The Word document that was converted

    OUTPUT_MD: The Markdown file made from it
    """
    from rich.table import Table
    from docx2md.verify import verify as verify_conversion

    try:
        checks = verify_conversion(input_docx, output_md)
    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {str(e)}")
        logger.error(f"Verification failed: {str(e)}", exc_info=True)
        sys.exit(1)

    table = Table(title=f"{Path(input_docx).name} → {Path(output_md).name}")
    table.add_column("")
    table.add_column("Check")
    table.add_column(".docx", justify="right")
    table.add_column(".md", justify="right")
    for check in checks:
        table.add_row("✅" if check.ok else "❌", check.name, check.docx, check.md)
    console.print(table)
    for check in checks:
        for detail in check.details:
            style = "red" if not check.ok else "dim"
            console.print(f"[{style}]{check.name}:[/{style}] {detail}")

    if all(check.ok for check in checks):
        console.print("[bold green]The Markdown holds the whole document.[/bold green]")
    else:
        failed = [check.name for check in checks if not check.ok]
        console.print(f"[bold red]Lost or changed in conversion:[/bold red] {', '.join(failed)}")
        sys.exit(1)


if __name__ == "__main__":
    main()