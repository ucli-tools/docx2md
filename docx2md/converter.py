"""
Core conversion functionality for docx2md.

This module provides the main functions for converting Word documents to
Markdown+LaTeX, including the processing pipeline and batch conversion.

Pipeline order:
  1. Extract DOCX metadata (python-docx)
  2. pandoc: docx → raw markdown  +  --extract-media=./img
  3. WordCleanupProcessor   — remove TOC, heading markup, heading IDs, fix image paths
  3b. FrontMatterStructureProcessor — detect body front matter, wrap with headings
  4. UnicodeFixProcessor    — replace Unicode math chars with LaTeX
  5. FigureProcessor        — fix double figure captions
  6. EquationFixProcessor   — fix garbled OMML equation patterns
  7. fix_delimiters         — \\(...\\) → $...$ and \\[...\\] → $$...$$
  8. process_tables         — normalize pipe/grid tables
  9. extract_and_process_images — optimize/rename images, fix remaining path issues
 10. generate_yaml_frontmatter  — prepend mdtexpdf YAML frontmatter
"""

import os
import tempfile
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

import pypandoc

from docx2md.processors.cleanup import WordCleanupProcessor
from docx2md.processors.docx import extract_docx_content
from docx2md.processors.equation_fix import EquationFixProcessor
from docx2md.processors.equations import fix_delimiters
from docx2md.processors.figures import FigureProcessor
from docx2md.processors.frontmatter import generate_yaml_frontmatter
from docx2md.processors.images import extract_and_process_images
from docx2md.processors.tables import process_tables
from docx2md.processors.unicode_fix import UnicodeFixProcessor
from docx2md.utils.logging_utils import get_logger

logger = get_logger(__name__)


def convert_docx_to_markdown(
    input_file: Union[str, Path],
    output_file: Union[str, Path],
    config: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """
    Convert a Word document to Markdown+LaTeX.

    Args:
        input_file: Path to the Word document (.docx) to convert
        output_file: Path for the output Markdown file
        config: Configuration dictionary (optional)

    Returns:
        Dict containing statistics about the conversion
    """
    if config is None:
        config = {}

    input_path = Path(input_file)
    output_path = Path(output_file)

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    if input_path.suffix.lower() != ".docx":
        raise ValueError(f"Input file must be a .docx file: {input_path}")

    output_path.parent.mkdir(parents=True, exist_ok=True)

    logger.info(f"Converting {input_path} to {output_path}")

    # Step 1: Extract metadata from the Word document (title, author, etc.)
    doc_content = extract_docx_content(input_path)

    # Step 2: pandoc docx → markdown with image extraction
    media_dir = output_path.parent / config.get("images", {}).get("extract_path", "./img")
    media_dir.mkdir(parents=True, exist_ok=True)

    extra_args = list(config.get("pandoc", {}).get("extra_args", ["--wrap=none"]))
    extra_args.append(f"--extract-media={media_dir}")

    # Never use --mathml: we want LaTeX $...$ output, not MathML
    # (use_pandoc_mathml is kept in config for backwards compat but defaults False)

    stats: Dict[str, Any] = {}

    processing = config.get("processing", {})
    math_extraction_enabled = processing.get("math_extraction", True)
    skip_equation_fix = False
    skip_fix_delimiters = False

    if math_extraction_enabled:
        try:
            from docx2md.processors.math_extraction import MathExtractor

            extractor = MathExtractor(config)
            markdown_content, math_stats = extractor.extract_and_convert(
                input_path, media_dir, extra_args
            )
            stats.update(math_stats)
            skip_equation_fix = True
            skip_fix_delimiters = True
            logger.info(
                "Math extraction: %d equations processed",
                math_stats.get("math_equations_extracted", 0),
            )
        except Exception as e:
            logger.warning(f"Math extraction failed, falling back to pandoc: {e}")
            markdown_content = pypandoc.convert_file(
                str(input_path),
                "markdown",
                format="docx",
                extra_args=extra_args,
            )
    else:
        markdown_content = pypandoc.convert_file(
            str(input_path),
            "markdown",
            format="docx",
            extra_args=extra_args,
        )

    # Step 3: Word structural cleanup (TOC, heading markup, image paths)
    if processing.get("cleanup", True):
        proc = WordCleanupProcessor(config, output_dir=output_path.parent)
        markdown_content = proc.process(markdown_content)

    # Step 3b: Detect and structure body front matter (dedication, copyright, title repeats)
    if processing.get("structure_front_matter", True):
        from docx2md.processors.front_matter_structure import FrontMatterStructureProcessor
        proc = FrontMatterStructureProcessor(config)
        markdown_content = proc.process(
            markdown_content,
            doc_properties=doc_content.get("properties", {}),
        )

    # Step 4: Unicode → LaTeX replacement
    if processing.get("fix_unicode", True):
        proc = UnicodeFixProcessor(config)
        markdown_content = proc.process(markdown_content)

    # Step 5: Figure caption fixing (replaces AI alt-text with real captions)
    if processing.get("fix_figures", True):
        proc = FigureProcessor(config)
        markdown_content = proc.process(markdown_content)

    # Step 6: Equation fix (garbled OMML patterns)
    # Skipped when math extraction succeeded — equations already clean
    if processing.get("fix_equations", True) and not skip_equation_fix:
        proc = EquationFixProcessor(config)
        markdown_content = proc.process(markdown_content)

    # Steps 7–9: file-based processors (write to temp file, process, read back)

    # Create temp file in the output directory so that the images processor can
    # resolve relative image paths (e.g. media/media/img.png) correctly.
    with tempfile.NamedTemporaryFile(
        suffix=".md", mode="w+", encoding="utf-8", delete=False,
        dir=output_path.parent,
    ) as tmp:
        tmp_path = tmp.name
        tmp.write(markdown_content)

    try:
        # Step 7: fix \(...\) → $...$ and \[...\] → $$...$$
        # Skipped when math extraction succeeded — delimiters already correct
        if processing.get("fix_delimiters", True) and not skip_fix_delimiters:
            inline_delimiters = tuple(
                config.get("equations", {}).get("inline_delimiters", ["$", "$"])
            )
            display_delimiters = tuple(
                config.get("equations", {}).get("display_delimiters", ["$$", "$$"])
            )
            eq_stats = fix_delimiters(
                tmp_path, tmp_path,
                inline_delimiters=inline_delimiters,
                display_delimiters=display_delimiters,
            )
            stats.update(eq_stats)

        # Step 8: normalize pipe tables
        if processing.get("process_tables", True):
            table_stats = process_tables(
                tmp_path, tmp_path,
                table_format=config.get("tables", {}).get("format", "pipe"),
                header_style=config.get("tables", {}).get("header_style", "bold"),
            )
            stats.update(table_stats)

        # Step 9: image path cleanup / optimization
        if processing.get("extract_images", True):
            img_cfg = config.get("images", {})
            image_stats = extract_and_process_images(
                tmp_path, tmp_path,
                images_dir=str(media_dir),
                optimize=img_cfg.get("optimize", False),
                max_width=img_cfg.get("max_width", 1200),
                max_height=img_cfg.get("max_height", 900),
            )
            stats.update(image_stats)

        # Read back processed content
        with open(tmp_path, "r", encoding="utf-8") as f:
            markdown_content = f.read()

    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass

    # Step 10: final LaTeX sanitization (runs AFTER fix_delimiters may have
    # introduced new $$ or changed delimiter forms)
    if processing.get("cleanup", True):
        from docx2md.processors.cleanup import final_sanitize
        markdown_content = final_sanitize(markdown_content)

    # Step 11: prepend YAML frontmatter for mdtexpdf
    if processing.get("generate_frontmatter", True):
        frontmatter, markdown_content = generate_yaml_frontmatter(
            doc_content.get("properties", {}),
            config,
            input_path,
            markdown_content,
        )
        markdown_content = frontmatter + markdown_content

        # Create metadata.yaml on first run (never overwrite user edits).
        # mdtexpdf auto-detects this file and uses it instead of embedded YAML.
        metadata_yaml_path = output_path.parent / "metadata.yaml"
        if not metadata_yaml_path.exists():
            yaml_lines = frontmatter.strip().split('\n')
            # Strip --- delimiters for standalone YAML file
            if yaml_lines and yaml_lines[0].strip() == '---':
                yaml_lines = yaml_lines[1:]
            if yaml_lines and yaml_lines[-1].strip() == '---':
                yaml_lines = yaml_lines[:-1]
            with open(metadata_yaml_path, "w", encoding="utf-8") as f:
                f.write('\n'.join(yaml_lines) + '\n')
            logger.info(f"Created metadata.yaml at {metadata_yaml_path}")

    # Write final output
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(markdown_content)

    logger.info(f"Conversion completed: {input_path} -> {output_path}")

    return {
        "input_file": str(input_path),
        "output_file": str(output_path),
        "equations_count": (
            stats.get("math_equations_extracted", 0)
            or stats.get("inline_fixed", 0) + stats.get("display_fixed", 0)
        ),
        "images_count": stats.get("images_processed", 0),
        "tables_count": stats.get("tables_processed", 0),
        **stats,
    }


def batch_convert(
    input_dir: Union[str, Path],
    output_dir: Union[str, Path],
    config: Optional[Dict[str, Any]] = None,
    recursive: bool = False,
    parallel: bool = True,
) -> Dict[str, Any]:
    """
    Convert multiple Word documents to Markdown+LaTeX.

    Args:
        input_dir: Directory containing Word documents (.docx) to convert
        output_dir: Directory for output Markdown files
        config: Configuration dictionary (optional)
        recursive: Whether to recursively process subdirectories
        parallel: Whether to process files in parallel

    Returns:
        Dict containing statistics about the batch conversion
    """
    if config is None:
        config = {}

    input_path = Path(input_dir)
    output_path = Path(output_dir)

    if not input_path.exists() or not input_path.is_dir():
        raise NotADirectoryError(f"Input directory not found: {input_path}")

    output_path.mkdir(parents=True, exist_ok=True)

    pattern = "**/*.docx" if recursive else "*.docx"
    docx_files = list(input_path.glob(pattern))

    if not docx_files:
        logger.warning(f"No .docx files found in {input_path}")
        return {"files_processed": 0, "files_succeeded": 0, "files_failed": 0, "results": []}

    logger.info(f"Found {len(docx_files)} .docx files to process")

    conversion_tasks = []
    for docx_file in docx_files:
        rel_path = docx_file.relative_to(input_path)
        out_file = output_path / rel_path.with_suffix(".md")
        out_file.parent.mkdir(parents=True, exist_ok=True)
        conversion_tasks.append((docx_file, out_file))

    results = []
    failed_count = 0

    if parallel and len(conversion_tasks) > 1:
        with ProcessPoolExecutor() as executor:
            futures = {
                executor.submit(convert_docx_to_markdown, inp, out, config): (inp, out)
                for inp, out in conversion_tasks
            }
            for future in as_completed(futures):
                inp, out = futures[future]
                try:
                    results.append(future.result())
                except Exception as e:
                    failed_count += 1
                    logger.error(f"Failed to convert {inp}: {e}")
                    results.append({"input_file": str(inp), "output_file": str(out),
                                    "error": str(e), "success": False})
    else:
        for inp, out in conversion_tasks:
            try:
                results.append(convert_docx_to_markdown(inp, out, config))
            except Exception as e:
                failed_count += 1
                logger.error(f"Failed to convert {inp}: {e}")
                results.append({"input_file": str(inp), "output_file": str(out),
                                 "error": str(e), "success": False})

    return {
        "files_processed": len(conversion_tasks),
        "files_succeeded": len(conversion_tasks) - failed_count,
        "files_failed": failed_count,
        "results": results,
    }
