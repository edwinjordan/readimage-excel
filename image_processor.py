"""
Image Processor Module
Extracts features and text from images using OCR and computer vision techniques.
"""

import cv2
import numpy as np
from PIL import Image
import pytesseract
from typing import Dict, List, Tuple, Any


class ImageProcessor:
    """Process images and extract features for Excel export."""
    
    def __init__(self):
        """Initialize the image processor."""
        pass
    
    def load_image(self, image_path: str) -> np.ndarray:
        """
        Load an image from the given path.
        
        Args:
            image_path: Path to the image file
            
        Returns:
            Loaded image as numpy array
        """
        image = cv2.imread(image_path)
        if image is None:
            raise ValueError(f"Could not load image from {image_path}")
        return image
    
    def extract_text(self, image_path: str) -> str:
        """
        Extract text from image using OCR (Optical Character Recognition).
        
        Args:
            image_path: Path to the image file
            
        Returns:
            Extracted text as string
        """
        try:
            image = Image.open(image_path)
            text = pytesseract.image_to_string(image)
            return text.strip()
        except Exception as e:
            return f"Error extracting text: {str(e)}"
    
    def get_image_properties(self, image_path: str) -> Dict[str, Any]:
        """
        Get basic properties of an image.
        
        Args:
            image_path: Path to the image file
            
        Returns:
            Dictionary containing image properties
        """
        image = self.load_image(image_path)
        height, width = image.shape[:2]
        channels = image.shape[2] if len(image.shape) > 2 else 1
        
        # Calculate file size
        import os
        file_size = os.path.getsize(image_path)
        
        return {
            'width': width,
            'height': height,
            'channels': channels,
            'file_size_kb': round(file_size / 1024, 2),
            'aspect_ratio': round(width / height, 2),
            'total_pixels': width * height
        }
    
    def detect_dominant_colors(self, image_path: str, n_colors: int = 5) -> List[Tuple[int, int, int]]:
        """
        Detect dominant colors in the image using K-means clustering.
        
        Args:
            image_path: Path to the image file
            n_colors: Number of dominant colors to detect
            
        Returns:
            List of RGB tuples representing dominant colors
        """
        image = self.load_image(image_path)
        image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
        
        # Reshape image to be a list of pixels
        pixels = image.reshape(-1, 3)
        pixels = np.float32(pixels)
        
        # Apply K-means clustering
        criteria = (cv2.TERM_CRITERIA_EPS + cv2.TERM_CRITERIA_MAX_ITER, 100, 0.2)
        _, labels, centers = cv2.kmeans(pixels, n_colors, None, criteria, 10, cv2.KMEANS_RANDOM_CENTERS)
        
        # Convert centers to integers
        centers = np.uint8(centers)
        
        return [tuple(color) for color in centers]
    
    def calculate_brightness(self, image_path: str) -> float:
        """
        Calculate average brightness of the image.
        
        Args:
            image_path: Path to the image file
            
        Returns:
            Average brightness value (0-255)
        """
        image = self.load_image(image_path)
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        return round(float(np.mean(gray)), 2)
    
    def detect_edges(self, image_path: str) -> int:
        """
        Detect number of edges in the image using Canny edge detection.
        
        Args:
            image_path: Path to the image file
            
        Returns:
            Number of edge pixels detected
        """
        image = self.load_image(image_path)
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        edges = cv2.Canny(gray, 100, 200)
        return int(np.sum(edges > 0))
    
    def extract_all_features(self, image_path: str) -> Dict[str, Any]:
        """
        Extract all features from an image.
        
        Args:
            image_path: Path to the image file
            
        Returns:
            Dictionary containing all extracted features
        """
        features = {}
        
        # Add file path
        features['image_path'] = image_path
        
        # Extract text
        features['extracted_text'] = self.extract_text(image_path)
        
        # Get image properties
        properties = self.get_image_properties(image_path)
        features.update(properties)
        
        # Calculate brightness
        features['avg_brightness'] = self.calculate_brightness(image_path)
        
        # Detect edges
        features['edge_count'] = self.detect_edges(image_path)
        
        # Get dominant colors
        dominant_colors = self.detect_dominant_colors(image_path, n_colors=3)
        for i, color in enumerate(dominant_colors):
            features[f'dominant_color_{i+1}'] = f"RGB{color}"
        
        return features
