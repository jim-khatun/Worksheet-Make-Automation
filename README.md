# Bulk Image to PowerPoint Automation

## Description
This Python script automates creating PowerPoint presentations from a folder of images. It generates **Letter-sized slides (8.5″ × 11″)** with a **black rectangle border (7.5″ × 10″)** and inserts **one image per slide**, centered inside the border. The script can also optionally export the presentation as a **PDF** and **JPG images** for each slide.

## Features
- Automatically create slides from all images in a folder  
- Optional PDF export  
- Optional JPG export of each slide  
- Works with any folder on your computer  
- Fully dynamic: select input folder and output locations via dialogs  
- Repeatable workflow for bulk image processing  

## Requirements
- Windows computer  
- Python 3.x (only for running `.py` file, not needed if using EXE)  
- Libraries: `python-pptx`, `Pillow`, `pywin32` (only for `.py`)  
- Microsoft PowerPoint installed (required for PDF and JPG export)  

## Usage
### Using Python script
1. Run the script:  
   ```bash
   python create_ppt_complete.py
