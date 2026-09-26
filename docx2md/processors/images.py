"""
Image processing module for docx2md.

This module provides functions for extracting and processing images
from Word documents and handling them in Markdown files.
"""

import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

from PIL import Image

from docx2md.processors.base import BaseProcessor
from docx2md.utils.file_utils import ensure_directory, read_file, write_file
from docx2md.utils.logging_utils import get_logger

logger = get_logger(__name__)


class ImageProcessor(BaseProcessor):
    """
    Processor for images in Markdown.
    """
    
    def __init__(self, config: Optional[Dict[str, Any]] = None):
        """
        Initialize the image processor.
        
        Args:
            config: Configuration dictionary
        """
        super().__init__(config)
        
        # Get image configuration
        img_config = self.config.get("images", {})
        self.extract_path = img_config.get("extract_path", "./images")
        self.optimize = img_config.get("optimize", True)
        self.max_width = img_config.get("max_width", 800)
        self.max_height = img_config.get("max_height", 600)
    
    def process(self, content: str) -> str:
        """
        Process images in the content.
        
        Args:
            content: Markdown content with image references
            
        Returns:
            Processed content with updated image references
        """
        # Find all image references
        image_pattern = r'!\[(.*?)\]\((.*?)\)'
        
        def replace_image(match):
            alt_text = match.group(1)
            image_path = match.group(2)
            
            # Process the image if it exists
            if os.path.exists(image_path):
                new_path = self._process_image(image_path)
                return f'![{alt_text}]({new_path})'
            
            return match.group(0)
        
        # Replace image references
        processed_content = re.sub(image_pattern, replace_image, content)
        
        return processed_content
    
    def _process_image(self, image_path: str) -> str:
        """
        Process an image file.
        
        Args:
            image_path: Path to the image file
            
        Returns:
            New path to the processed image
        """
        try:
            # Ensure the extract directory exists
            extract_dir = ensure_directory(self.extract_path)
            
            # Get the image filename
            image_filename = os.path.basename(image_path)
            
            # Construct the new path
            new_path = os.path.join(self.extract_path, image_filename)
            
            # If optimization is disabled, just copy the file
            if not self.optimize:
                import shutil
                shutil.copy2(image_path, new_path)
                return new_path
            
            # Open and optimize the image
            with Image.open(image_path) as img:
                # Resize if needed
                width, height = img.size
                if width > self.max_width or height > self.max_height:
                    # Calculate new dimensions while preserving aspect ratio
                    ratio = min(self.max_width / width, self.max_height / height)
                    new_width = int(width * ratio)
                    new_height = int(height * ratio)
                    img = img.resize((new_width, new_height), Image.LANCZOS)
                
                # Save the optimized image
                img.save(new_path, optimize=True, quality=85)
            
            return new_path
        
        except Exception as e:
            logger.error(f"Failed to process image {image_path}: {str(e)}")
            return image_path


def extract_and_process_images(
    input_file: Union[str, Path],
    output_file: Union[str, Path],
    images_dir: str = "./images",
    optimize: bool = True,
    max_width: int = 800,
    max_height: int = 600,
) -> Dict[str, Any]:
    """
    Extract and process images referenced in a Markdown file.
    
    Args:
        input_file: Path to the input Markdown file
        output_file: Path to the output Markdown file
        images_dir: Directory to store processed images
        optimize: Whether to optimize images
        max_width: Maximum image width
        max_height: Maximum image height
        
    Returns:
        Dictionary with statistics about the processing
    """
    logger.info(f"Processing images in {input_file}")
    
    # Read the input file
    content = read_file(input_file)
    if content is None:
        raise ValueError(f"Failed to read input file: {input_file}")
    
    # Ensure the images directory exists
    os.makedirs(images_dir, exist_ok=True)
    
    # Find all image references
    image_pattern = r'!\[(.*?)\]\((.*?)\)'
    image_matches = re.findall(image_pattern, content)
    
    processed_images = 0
    failed_images = 0
    images_root = os.path.abspath(images_dir)
    # Sources copied up from a subdirectory of images_dir (pandoc's media/)
    extracted_sources = set()

    # Process each image
    for alt_text, image_path in image_matches:
        # Normalize the image path
        norm_path = image_path.strip()
        
        # Skip URLs
        if norm_path.startswith(("http://", "https://")):
            continue
        
        # Construct the full path
        full_path = norm_path
        if not os.path.isabs(norm_path):
            # If the path is relative, make it relative to the input file directory
            input_dir = os.path.dirname(os.path.abspath(input_file))
            full_path = os.path.join(input_dir, norm_path)
        
        # Check if the image exists
        if not os.path.exists(full_path):
            logger.warning(f"Image not found: {full_path}")
            failed_images += 1
            continue
        
        try:
            # Generate a new filename
            image_filename = os.path.basename(full_path)
            new_path = os.path.join(images_dir, image_filename)
            
            # Process the image if optimization is enabled
            if optimize:
                with Image.open(full_path) as img:
                    # Resize if needed
                    width, height = img.size
                    if width > max_width or height > max_height:
                        # Calculate new dimensions while preserving aspect ratio
                        ratio = min(max_width / width, max_height / height)
                        new_width = int(width * ratio)
                        new_height = int(height * ratio)
                        img = img.resize((new_width, new_height), Image.LANCZOS)
                    
                    # Save the optimized image
                    img.save(new_path, optimize=True, quality=85)
            elif os.path.abspath(full_path) != os.path.abspath(new_path):
                # Just copy the file
                import shutil
                shutil.copy2(full_path, new_path)

            src = os.path.abspath(full_path)
            if src != os.path.abspath(new_path) and src.startswith(images_root + os.sep):
                extracted_sources.add(src)

            # Calculate the relative path from the output file to the image
            output_dir = os.path.dirname(os.path.abspath(output_file))
            
            # Get the path relative to the output directory
            rel_new_path = os.path.relpath(new_path, output_dir)
            
            # Remove 'files/output/' prefix if present
            if rel_new_path.startswith('files/output/'):
                rel_new_path = rel_new_path[len('files/output/'):]
            
            # Ensure the path starts with './'
            if not rel_new_path.startswith('./'):
                rel_new_path = './' + rel_new_path
            
            # Use this as the relative path
            rel_path = rel_new_path
            
            # Replace the image reference in the content
            old_ref = f'![{alt_text}]({image_path})'
            new_ref = f'![{alt_text}]({rel_path})'
            content = content.replace(old_ref, new_ref)
            
            processed_images += 1
            
        except Exception as e:
            logger.error(f"Failed to process image {full_path}: {str(e)}")
            failed_images += 1

    # pandoc extracts into images_dir/media/. Every image referenced from there
    # now has its copy in images_dir, so the extraction copies are duplicates:
    # remove them, then any subdirectories they leave empty.
    for src in extracted_sources:
        if os.path.exists(src):
            os.remove(src)
    for dirpath, _, _ in os.walk(images_root, topdown=False):
        if dirpath != images_root and not os.listdir(dirpath):
            os.rmdir(dirpath)

    # Final fix: directly replace any remaining incorrect image paths
    # This is a fallback in case the earlier replacements didn't catch everything
    
    # Fix paths with 'files/output/' prefix
    content = re.sub(
        r'!\[(.*?)\]\(files/output/(\.\/images/.*?)\)',
        r'![\1](\2)',
        content
    )
    
    # Fix paths with just 'files/output' prefix (no trailing slash)
    content = re.sub(
        r'!\[(.*?)\]\(files/output(\.\/images/.*?)\)',
        r'![\1](\2)',
        content
    )
    
    # Fix any other paths that might have 'files/output' anywhere in them
    content = re.sub(
        r'!\[(.*?)\]\((.*?)files/output/(.*?images/.*?)\)',
        r'![\1](./\3)',
        content
    )
    
    # Ensure all image paths start with ./images/
    content = re.sub(
        r'!\[(.*?)\]\((?!\.\/images)(.*?)images/(.*?)\)',
        r'![\1](./images/\3)',
        content
    )
    
    # Remove width and height attributes from image tags (a title page's
    # logo keeps its own)
    from docx2md.processors.cleanup import strip_image_sizes
    content = strip_image_sizes(content)
    
    # Write the output file
    if not write_file(output_file, content):
        raise ValueError(f"Failed to write output file: {output_file}")
    
    logger.info(f"Processed {processed_images} images, {failed_images} failed")
    
    return {
        "images_processed": processed_images,
        "images_failed": failed_images,
        "total_images": len(image_matches),
    }