# readimage-excel

A Python tool to extract features from images and export them to Excel files. This tool uses computer vision and OCR (Optical Character Recognition) to analyze images and extract valuable information.

## Features

- **Text Extraction**: Extract text from images using OCR (pytesseract)
- **Image Properties**: Get width, height, channels, file size, aspect ratio, and pixel count
- **Color Analysis**: Detect dominant colors in images using K-means clustering
- **Brightness Analysis**: Calculate average brightness levels
- **Edge Detection**: Count edges using Canny edge detection
- **Excel Export**: Export all extracted features to formatted Excel files
- **Batch Processing**: Process multiple images or entire directories at once
- **Summary Reports**: Generate summary statistics for multiple images

## Installation

1. Clone this repository:
```bash
git clone https://github.com/edwinjordan/readimage-excel.git
cd readimage-excel
```

2. Install Python dependencies:
```bash
pip install -r requirements.txt
```

3. Install Tesseract OCR:
   - **Ubuntu/Debian**: `sudo apt-get install tesseract-ocr`
   - **macOS**: `brew install tesseract`
   - **Windows**: Download installer from [GitHub](https://github.com/UB-Mannheim/tesseract/wiki)

## Usage

### Command Line Interface

#### Process a single image:
```bash
python main.py -i image.jpg -o output.xlsx
```

#### Process multiple images:
```bash
python main.py -i image1.jpg image2.png image3.jpg -o output.xlsx
```

#### Process all images in a directory:
```bash
python main.py -d ./images -o output.xlsx
```

#### Process directory recursively with summary:
```bash
python main.py -d ./images -o output.xlsx --recursive --summary
```

### Command Line Arguments

- `-i, --image`: Path to one or more image file(s)
- `-d, --directory`: Path to directory containing images
- `-o, --output`: Output Excel file path (required)
- `-r, --recursive`: Process subdirectories recursively (only with `-d`)
- `-s, --summary`: Include summary sheet for multiple images

### Python API

You can also use the modules directly in your Python code:

```python
from image_processor import ImageProcessor
from excel_exporter import ExcelExporter

# Initialize
processor = ImageProcessor()
exporter = ExcelExporter()

# Process single image
features = processor.extract_all_features('image.jpg')
exporter.export_single_image(features, 'output.xlsx')

# Process multiple images
features_list = []
for image_path in ['img1.jpg', 'img2.jpg', 'img3.jpg']:
    features = processor.extract_all_features(image_path)
    features_list.append(features)

exporter.export_with_summary(features_list, 'output.xlsx')
```

## Extracted Features

The tool extracts the following features from each image:

| Feature | Description |
|---------|-------------|
| `image_path` | Full path to the image file |
| `extracted_text` | Text extracted using OCR |
| `width` | Image width in pixels |
| `height` | Image height in pixels |
| `channels` | Number of color channels (1 for grayscale, 3 for RGB) |
| `file_size_kb` | File size in kilobytes |
| `aspect_ratio` | Width to height ratio |
| `total_pixels` | Total number of pixels |
| `avg_brightness` | Average brightness (0-255) |
| `edge_count` | Number of detected edges |
| `dominant_color_1/2/3` | Top 3 dominant colors in RGB format |

## Dependencies

- `openpyxl`: Excel file creation and manipulation
- `Pillow`: Image loading and basic processing
- `opencv-python`: Computer vision operations
- `pytesseract`: OCR text extraction
- `numpy`: Numerical operations

## Requirements

- Python 3.7 or higher
- Tesseract OCR installed on your system

## Example Output

The tool generates Excel files with:
- **Single Image**: Two-column format (Feature | Value)
- **Multiple Images**: Rows for each image, columns for each feature
- **With Summary**: Additional summary sheet with statistics

## Supported Image Formats

- JPEG (.jpg, .jpeg)
- PNG (.png)
- BMP (.bmp)
- TIFF (.tiff)
- GIF (.gif)
- WebP (.webp)

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is open source and available under the MIT License.

## Troubleshooting

### Tesseract not found error
Make sure Tesseract OCR is installed and accessible from your PATH. On Windows, you might need to set the path manually:
```python
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
```

### OpenCV import error
Try reinstalling opencv-python:
```bash
pip uninstall opencv-python
pip install opencv-python
```