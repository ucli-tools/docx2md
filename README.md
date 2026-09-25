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
