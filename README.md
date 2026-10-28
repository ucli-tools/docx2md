# docx2md: MS Word to Markdown+LaTeX Converter

[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)

A powerful tool for converting Microsoft Word documents with complex mathematical content to Markdown with LaTeX equations.

## Table of Contents

- [docx2md: MS Word to Markdown+LaTeX Converter](#docx2md-ms-word-to-markdownlatex-converter)
  - [Table of Contents](#table-of-contents)
  - [Overview](#overview)
  - [Features](#features)
  - [Prerequisites](#prerequisites)
  - [Get Started](#get-started)
    - [Using the Makefile](#using-the-makefile)
  - [Verifying a Conversion](#verifying-a-conversion)
  - [Troubleshooting](#troubleshooting)
    - [Common Issues](#common-issues)
    - [Logging](#logging)
  - [Contributing](#contributing)
  - [License](#license)

## Overview

docx2md is designed to convert Microsoft Word documents containing complex mathematical equations, tables, and figures into clean Markdown with LaTeX. It's particularly useful for technical books, academic papers, and scientific documentation.

This tool bridges the gap between Word documents and the Markdown+LaTeX workflow, enabling you to leverage tools like [mdtexpdf](https://github.com/ucli-tools/mdtexpdf) for generating beautiful PDFs.

## Features

- **Equation Conversion**: Accurately converts Word equations to LaTeX; Word equation numbers `#(1.2.3)` become `\tag{1.2.3}`, and font runs such as `\mathbb{\in R}` are trimmed to the letters they style (`\in \mathbb{R}`)
- **Image Extraction**: Properly extracts and references images into one flat image directory
- **Table Handling**: Converts Word tables to Markdown tables
- **Batch Processing**: Process multiple files at once
- **LaTeX Delimiter Fixing**: Standardizes equation delimiters
- **Cross-Reference Handling**: Maintains document references
- **Front Matter**: The document's own dedication, copyright page and title page become `# Dedication`, `# Copyright Page` and `# Title Page`; the title page keeps its logo and the type sizes Word set it in (`[*Title*]{size="28pt"}`)
- **Verification**: `docx2md verify` compares the Markdown with the Word document and names anything lost or changed
- **Modular Architecture**: Extensible for custom processing needs
- **Robust CLI**: Git-like command structure for ease of use

## Prerequisites

- Python 3.8 or higher
- Pandoc (for Word document conversion)
- Microsoft Word documents with equations created using the built-in equation editor

## Get Started

You can use docx2md directly from the cloned repository.

This approach is useful for quick testing.

```bash
# Clone the repository
git clone https://github.com/ucli-tools/docx2md.git
cd docx2md

# Create input and output directories
mkdir -p files/input files/output

# Copy your Word documents to the input directory
cp /path/to/your/chapter1.docx files/input/
cp /path/to/your/chapter2.docx files/input/
# ... and so on for other chapters

# Set up a virtual environment and install the package
uv venv
uv pip install -e .  # Install the package in development mode with all dependencies

# Activate the virtual environment
# For bash/zsh:
#source .venv/bin/activate
# For fish shell:
source .venv/bin/activate.fish

# Run the batch conversion
python3 -m docx2md batch ./files/input ./files/output
```

> **Note**: If you encounter an error about missing modules, make sure to install the package with `-e .` instead of `-r requirements.txt` to ensure all dependencies are correctly installed.

### Using the Makefile

You can also use the provided Makefile to simplify the process:

```bash
# For bash/zsh shell
make local-run

# For fish shell
make local-run-fish
```

This will:
1. Set up a virtual environment
2. Install the package with all dependencies
3. Run the batch conversion on files in ./files/input and output to ./files/output

## Correcting a Word Document, and Proving the Correction

When a book's original must change, two commands make the change exact and its proof readable.

`docx2md edit` makes a small correction (a value, a word) in one place, and copies every other part
of the document byte for byte. It refuses unless the text is there exactly once, in the table cell
or the paragraph named:

```bash
docx2md edit book.docx book-corrected.docx --replace "10iπ" "20iπ" --table 1 --row 14 --cell 12
docx2md edit book.docx book-corrected.docx --replace "teh" "the" --paragraph "the words around teh typo"
```

If several paragraphs hold the `--paragraph` words (identical paragraphs), `edit` refuses; add
`--paragraph-nth K` (counted from 1, in document order) to use the K-th one. `--nth` still counts
the occurrences of the text inside that paragraph:

```bash
docx2md edit book.docx book-corrected.docx --replace "teh" "the" --paragraph "the words around teh typo" --paragraph-nth 2
```

`docx2md diff` compares two Word documents passage by passage (a paragraph, or a table row with its
cells) and lists every difference: where it is (the headings above it, the table and its caption,
the row, the cell), the whole passage before and after, and, given the document's PDF, its page. A
change of colour or style counts too, and so do the images and every other part of the file; what
Word rewrites on every save does not. `--expect N` fails unless exactly N passages changed and no
other part did; `-o` writes the report as Markdown, for a change log:

```bash
docx2md diff book.docx book-corrected.docx --pdf book.pdf --expect 1 -o change-0001.md
```

## Verifying a Conversion

`docx2md verify` reads the Word document's XML directly and compares it with the Markdown made from it, so a defect in the converter cannot hide itself:

```bash
docx2md verify book.docx book.md
```

| Check | What must match |
|---|---|
| equations | every Word equation is in the Markdown |
| equation numbers | the set of `#(1.2.3)` numbers and the `\tag{1.2.3}` numbers |
| punctuation before equation numbers | the comma or period that ends each numbered equation |
| images, figure captions | every picture, and the caption under it (side-by-side pictures included) |
| tables | every table, as pandoc reads the Markdown (pipe, grid, simple or multiline) |
| headings | every non-empty Word heading |
| index entries | every Word `XE` field and its `[index:...]` marker |
| footnotes | every note mark, each with its text |
| text | the running text, word by word (at most 0.5% of words may differ) |
| leftovers | no converter placeholders, raw `#(n)` numbers, stray math letters, control characters, or words split by an index marker |

It prints a table and each difference, and exits with status 1 when a check fails, so a build can stop there:

```make
md:
	docx2md convert book.docx book.md
	docx2md verify book.docx book.md
```

A failed check is a converter defect to fix in docx2md, not in the Markdown: fix it, convert again, and verify again.

## Troubleshooting

### Common Issues

- **Equations not converting properly**: Ensure equations in Word are created with the built-in equation editor
- **Images missing**: Check if the extraction path exists and has write permissions
- **Pandoc errors**: Verify Pandoc is installed and in your PATH

### Logging

Enable detailed logging for troubleshooting:

```bash
python3 -m docx2md convert --log-level debug document.docx
```

## Contributing

Contributions are welcome! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

## License

This project is licensed under the Apache License 2.0 - see the [LICENSE](LICENSE) file for details.
