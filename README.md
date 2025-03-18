# PowerPoint Utility Scripts

This repository contains two Python scripts that automate the processing of PowerPoint (.pptx) files:

1. **Convert PowerPoint to PDF and Images**: Converts a PowerPoint presentation into a PDF file and extracts each slide as a high-resolution image.
2. **Search for Text in PowerPoint Slides**: Searches for specific text within a PowerPoint file and displays matching content.

## Features

- **Convert PowerPoint (.pptx) to PDF**
- **Extract slides as images (JPEG format)**
- **Search for specific text inside a PowerPoint presentation**

## Requirements

Before running the scripts, ensure you have the following dependencies installed:

```sh
pip install pdf2image comtypes python-pptx
```

Additionally, you need to install **poppler** for PDF processing:

- Windows: Download and install from [Poppler for Windows](https://github.com/oschwartz10612/poppler-windows/releases)
- macOS: Install via Homebrew:
  ```sh
  brew install poppler
  ```
- Linux: Install via package manager:
  ```sh
  sudo apt install poppler-utils
  ```

## Usage

### 1. Convert PowerPoint to PDF and Images

```python
import os
import comtypes.client
from pdf2image import convert_from_path

def ppt_to_pdf_png(input_ppt, output_pdf):
    foldername = input_ppt.split('.')[0]
    if not os.path.exists(foldername):
        os.makedirs(foldername)
    
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 2  # Set visibility mode
    
    presentation = powerpoint.Presentations.Open(os.path.abspath(input_ppt))
    pdf_save_path = os.path.join(foldername, output_pdf)
    presentation.SaveAs(os.path.abspath(pdf_save_path), 32)  # Save as PDF (format 32)
    presentation.Close()
    powerpoint.Quit()
    
    pages = convert_from_path(pdf_save_path, 500)
    for count, page in enumerate(pages):
        page.save(os.path.join(foldername, f'out{count}.jpg'), 'JPEG')

ppt_to_pdf_png('example.pptx', 'example.pdf')
```

### 2. Search for Text in a PowerPoint File

```python
from pptx import Presentation
import os

def ppt_search_text(file_path, text):
    n = 0
    prs = Presentation(os.path.abspath(file_path))
    for slide in prs.slides:
        n += 1
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                shape.text = shape.text.lower()
                if str(text) in shape.text:
                    print(shape.text)
                    print('=================================')

ppt_search_text('example.pptx', 'artificial intelligence')
```

## Notes

- Ensure that Microsoft PowerPoint is installed on your system for the conversion process.
- The script extracts slides as images at 500 DPI for high resolution.
- The search function converts text to lowercase for case-insensitive matching.

## License

This project is open-source and available under the MIT License.

