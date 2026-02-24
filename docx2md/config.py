"""
Configuration handling for docx2md.

This module provides functions for loading and validating configuration
from YAML files, environment variables, and command-line arguments.
"""

import os
from pathlib import Path
from typing import Any, Dict, Optional, Union

import yaml

from docx2md.utils.logging_utils import get_logger

logger = get_logger(__name__)

# Default configuration
DEFAULT_CONFIG: Dict[str, Any] = {
    "equations": {
        "inline_delimiters": ["$", "$"],
        "display_delimiters": ["$$", "$$"],
        "use_pandoc_mathml": False,  # False = keep LaTeX $...$ (correct for mdtexpdf)
    },
    "images": {
        "extract_path": "./img",     # image extraction directory
        "optimize": False,           # don't resize/compress by default
        "max_width": 1200,
        "max_height": 900,
    },
    "tables": {
        "format": "pipe",  # Options: pipe, grid, simple
        "header_style": "bold",
    },
    "pandoc": {
        "extra_args": ["--wrap=none"],
    },
    "processing": {
        "math_extraction": True,
        "fix_delimiters": True,
        "extract_images": True,
        "process_tables": True,
        # New processors
        "cleanup": True,
        "fix_figures": True,
        "fix_unicode": True,
        "fix_equations": True,
        "generate_frontmatter": True,
        "structure_front_matter": True,
    },
    "cleanup": {
        "strip_triple_dollar": True,
        "remove_toc": True,
        "strip_heading_markup": True,
        "strip_heading_ids": True,
        "remove_image_attrs": True,
        "fix_image_paths": True,
    },
    "figures": {
        "enabled": True,
    },
    "unicode_fix": {
        "enabled": True,
        "custom_replacements": [],
        # Each entry: {"char": "X", "always": "Y"}  or
        #             {"char": "X", "text": "Y", "math": "Z"}
    },
    "equation_fix": {
        "enabled": True,
    },
    "section_numbers": {
        "style": "none",    # Options: none | arabic | roman | alpha
        "heading_level": 1,
    },
    "yaml_frontmatter": {
        "enabled": True,
        "extract_from_body": True,      # scan body for **Title**, *subtitle*, author
        "strip_body_title_block": True, # remove title/author lines from body after extraction
        "default_author": "",
        "mdtexpdf": {
            "format": "book",
            "toc": True,
            "no_numbers": True,
            "header_footer_policy": "all",
            "pageof": True,
            "date_footer": True,
            "copyright_page": True,
            "chapters_on_recto": True,
            "drop_caps": True,
        },
    },
}


def load_config(config_path: Optional[Union[str, Path]] = None) -> Dict[str, Any]:
    """
    Load configuration from a YAML file and merge with defaults.
    
    Args:
        config_path: Path to the configuration file (optional)
        
    Returns:
        Dict containing the merged configuration
    """
    config = DEFAULT_CONFIG.copy()
    
    # Check for environment variables
    env_config = _load_from_env()
    if env_config:
        _deep_merge(config, env_config)
    
    # Load from file if provided
    if config_path:
        file_config = _load_from_file(config_path)
        if file_config:
            _deep_merge(config, file_config)
    
    # Validate the configuration
    _validate_config(config)
    
    return config


def _load_from_file(config_path: Union[str, Path]) -> Dict[str, Any]:
    """
    Load configuration from a YAML file.
    
    Args:
        config_path: Path to the configuration file
        
    Returns:
        Dict containing the configuration from the file
    """
    try:
        with open(config_path, "r") as f:
            config = yaml.safe_load(f)
        
        if not isinstance(config, dict):
            logger.warning(f"Configuration file {config_path} does not contain a dictionary")
            return {}
        
        return config
    except Exception as e:
        logger.warning(f"Failed to load configuration from {config_path}: {str(e)}")
        return {}


def _load_from_env() -> Dict[str, Any]:
    """
    Load configuration from environment variables.
    
    Environment variables should be prefixed with DOCX2MD_ and use double underscores
    to indicate nesting, e.g., DOCX2MD_EQUATIONS__INLINE_DELIMITERS="$,$"
    
    Returns:
        Dict containing the configuration from environment variables
    """
    config: Dict[str, Any] = {}
    
    for key, value in os.environ.items():
        if key.startswith("DOCX2MD_"):
            parts = key[8:].lower().split("__")
            
            # Build nested dictionary
            current = config
            for part in parts[:-1]:
                if part not in current:
                    current[part] = {}
                current = current[part]
            
            # Set the value
            current[parts[-1]] = _parse_env_value(value)
    
    return config


def _parse_env_value(value: str) -> Any:
    """
    Parse an environment variable value into the appropriate type.
    
    Args:
        value: String value from environment variable
        
    Returns:
        Parsed value (bool, int, float, list, or string)
    """
    # Check for boolean
    if value.lower() in ("true", "yes", "1"):
        return True
    if value.lower() in ("false", "no", "0"):
        return False
    
    # Check for integer
    try:
        return int(value)
    except ValueError:
        pass
    
    # Check for float
    try:
        return float(value)
    except ValueError:
        pass
    
    # Check for list (comma-separated)
    if "," in value:
        return [item.strip() for item in value.split(",")]
    
    # Default to string
    return value


def _deep_merge(target: Dict[str, Any], source: Dict[str, Any]) -> None:
    """
    Deep merge two dictionaries, modifying the target in place.
    
    Args:
        target: Target dictionary to merge into
        source: Source dictionary to merge from
    """
    for key, value in source.items():
        if key in target and isinstance(target[key], dict) and isinstance(value, dict):
            _deep_merge(target[key], value)
        else:
            target[key] = value


def _validate_config(config: Dict[str, Any]) -> None:
    """
    Validate the configuration and set default values for missing fields.
    
    Args:
        config: Configuration dictionary to validate
    """
    # Ensure required sections exist
    for section in DEFAULT_CONFIG:
        if section not in config:
            config[section] = DEFAULT_CONFIG[section].copy()
    
    # Validate equation delimiters
    eq_config = config.get("equations", {})
    if not isinstance(eq_config.get("inline_delimiters"), list) or len(eq_config.get("inline_delimiters", [])) != 2:
        logger.warning("Invalid inline_delimiters configuration, using defaults")
        eq_config["inline_delimiters"] = DEFAULT_CONFIG["equations"]["inline_delimiters"]
    
    if not isinstance(eq_config.get("display_delimiters"), list) or len(eq_config.get("display_delimiters", [])) != 2:
        logger.warning("Invalid display_delimiters configuration, using defaults")
        eq_config["display_delimiters"] = DEFAULT_CONFIG["equations"]["display_delimiters"]