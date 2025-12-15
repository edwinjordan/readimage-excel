"""
Main Script for Image to Excel Feature Extraction
Command-line interface for extracting features from images and exporting to Excel.
"""

import argparse
import os
import sys
from pathlib import Path
from image_processor import ImageProcessor
from excel_exporter import ExcelExporter


def validate_image_file(file_path: str) -> bool:
    """
    Validate if the file is an image.
    
    Args:
        file_path: Path to the file
        
    Returns:
        True if file is a valid image, False otherwise
    """
    valid_extensions = ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.gif', '.webp']
    return os.path.splitext(file_path)[1].lower() in valid_extensions


def process_single_image(image_path: str, output_path: str):
    """
    Process a single image and export to Excel.
    
    Args:
        image_path: Path to the image file
        output_path: Path for the output Excel file
    """
    print(f"Processing image: {image_path}")
    
    processor = ImageProcessor()
    exporter = ExcelExporter()
    
    try:
        features = processor.extract_all_features(image_path)
        exporter.export_single_image(features, output_path)
        print("Processing completed successfully!")
    except Exception as e:
        print(f"Error processing image: {str(e)}")
        sys.exit(1)


def process_multiple_images(image_paths: list, output_path: str, with_summary: bool = False):
    """
    Process multiple images and export to Excel.
    
    Args:
        image_paths: List of paths to image files
        output_path: Path for the output Excel file
        with_summary: Whether to include a summary sheet
    """
    print(f"Processing {len(image_paths)} images...")
    
    processor = ImageProcessor()
    exporter = ExcelExporter()
    
    features_list = []
    
    for i, image_path in enumerate(image_paths, 1):
        print(f"  [{i}/{len(image_paths)}] Processing: {image_path}")
        try:
            features = processor.extract_all_features(image_path)
            features_list.append(features)
        except Exception as e:
            print(f"  Warning: Error processing {image_path}: {str(e)}")
            continue
    
    if not features_list:
        print("Error: No images were successfully processed")
        sys.exit(1)
    
    try:
        if with_summary:
            exporter.export_with_summary(features_list, output_path)
        else:
            exporter.export_multiple_images(features_list, output_path)
        print(f"Processing completed! Processed {len(features_list)} images successfully.")
    except Exception as e:
        print(f"Error exporting to Excel: {str(e)}")
        sys.exit(1)


def process_directory(directory_path: str, output_path: str, recursive: bool = False, with_summary: bool = False):
    """
    Process all images in a directory.
    
    Args:
        directory_path: Path to the directory containing images
        output_path: Path for the output Excel file
        recursive: Whether to process subdirectories recursively
        with_summary: Whether to include a summary sheet
    """
    image_files = []
    
    if recursive:
        for root, _, files in os.walk(directory_path):
            for file in files:
                file_path = os.path.join(root, file)
                if validate_image_file(file_path):
                    image_files.append(file_path)
    else:
        for file in os.listdir(directory_path):
            file_path = os.path.join(directory_path, file)
            if os.path.isfile(file_path) and validate_image_file(file_path):
                image_files.append(file_path)
    
    if not image_files:
        print(f"No image files found in {directory_path}")
        sys.exit(1)
    
    process_multiple_images(image_files, output_path, with_summary)


def main():
    """Main entry point for the CLI."""
    parser = argparse.ArgumentParser(
        description="Extract features from images and export to Excel",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Process a single image
  python main.py -i image.jpg -o output.xlsx
  
  # Process multiple images
  python main.py -i image1.jpg image2.png image3.jpg -o output.xlsx
  
  # Process all images in a directory
  python main.py -d ./images -o output.xlsx
  
  # Process directory recursively with summary
  python main.py -d ./images -o output.xlsx --recursive --summary
        """
    )
    
    input_group = parser.add_mutually_exclusive_group(required=True)
    input_group.add_argument('-i', '--image', nargs='+', help='Path to image file(s)')
    input_group.add_argument('-d', '--directory', help='Path to directory containing images')
    
    parser.add_argument('-o', '--output', required=True, help='Output Excel file path')
    parser.add_argument('-r', '--recursive', action='store_true', 
                        help='Process subdirectories recursively (only with -d)')
    parser.add_argument('-s', '--summary', action='store_true', 
                        help='Include summary sheet (only for multiple images)')
    
    args = parser.parse_args()
    
    # Validate output path
    if not args.output.endswith(('.xlsx', '.xls')):
        args.output += '.xlsx'
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Process based on input type
    if args.image:
        # Validate input files
        for image_path in args.image:
            if not os.path.exists(image_path):
                print(f"Error: File not found: {image_path}")
                sys.exit(1)
            if not validate_image_file(image_path):
                print(f"Error: Not a valid image file: {image_path}")
                sys.exit(1)
        
        if len(args.image) == 1:
            process_single_image(args.image[0], args.output)
        else:
            process_multiple_images(args.image, args.output, args.summary)
    
    elif args.directory:
        if not os.path.exists(args.directory):
            print(f"Error: Directory not found: {args.directory}")
            sys.exit(1)
        if not os.path.isdir(args.directory):
            print(f"Error: Not a directory: {args.directory}")
            sys.exit(1)
        
        process_directory(args.directory, args.output, args.recursive, args.summary)


if __name__ == "__main__":
    main()
