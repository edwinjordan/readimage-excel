"""
Example usage of the image to Excel feature extraction tool.
This script demonstrates how to use the modules programmatically.
"""

from image_processor import ImageProcessor
from excel_exporter import ExcelExporter


def example_single_image():
    """Example: Process a single image."""
    print("Example 1: Processing a single image")
    print("-" * 50)
    
    processor = ImageProcessor()
    exporter = ExcelExporter()
    
    # Replace with your image path
    image_path = "sample_image.jpg"
    
    try:
        # Extract features
        features = processor.extract_all_features(image_path)
        
        # Print features
        print("Extracted features:")
        for key, value in features.items():
            print(f"  {key}: {value}")
        
        # Export to Excel
        exporter.export_single_image(features, "single_image_output.xlsx")
        print("\nExcel file created: single_image_output.xlsx")
    except Exception as e:
        print(f"Error: {e}")


def example_multiple_images():
    """Example: Process multiple images."""
    print("\n\nExample 2: Processing multiple images")
    print("-" * 50)
    
    processor = ImageProcessor()
    exporter = ExcelExporter()
    
    # Replace with your image paths
    image_paths = [
        "image1.jpg",
        "image2.png",
        "image3.jpg"
    ]
    
    features_list = []
    
    for image_path in image_paths:
        try:
            print(f"Processing: {image_path}")
            features = processor.extract_all_features(image_path)
            features_list.append(features)
        except Exception as e:
            print(f"  Error processing {image_path}: {e}")
    
    if features_list:
        # Export to Excel with summary
        exporter.export_with_summary(features_list, "multiple_images_output.xlsx")
        print(f"\nProcessed {len(features_list)} images")
        print("Excel file created: multiple_images_output.xlsx")


def example_text_extraction():
    """Example: Extract text only from an image."""
    print("\n\nExample 3: Text extraction only")
    print("-" * 50)
    
    processor = ImageProcessor()
    
    # Replace with your image path
    image_path = "text_image.jpg"
    
    try:
        text = processor.extract_text(image_path)
        print(f"Extracted text from {image_path}:")
        print(text)
    except Exception as e:
        print(f"Error: {e}")


def example_color_analysis():
    """Example: Analyze dominant colors in an image."""
    print("\n\nExample 4: Color analysis")
    print("-" * 50)
    
    processor = ImageProcessor()
    
    # Replace with your image path
    image_path = "colorful_image.jpg"
    
    try:
        colors = processor.detect_dominant_colors(image_path, n_colors=5)
        print(f"Top 5 dominant colors in {image_path}:")
        for i, color in enumerate(colors, 1):
            print(f"  {i}. RGB{color}")
    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    print("Image to Excel Feature Extraction - Examples")
    print("=" * 50)
    print("\nNote: Update the image paths in this script to use your own images.\n")
    
    # Uncomment the examples you want to run:
    # example_single_image()
    # example_multiple_images()
    # example_text_extraction()
    # example_color_analysis()
    
    print("\nTo run these examples, uncomment the function calls at the bottom of this file.")
