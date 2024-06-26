#!/bin/bash

# Create and activate a virtual environment
python -m venv venv
source venv/bin/activate

# Install necessary packages
pip install PyMuPDF python-docx reportlab python-pptx googletrans==4.0.0-rc1

# Run the handle.py script
python handle.py

# Deactivate the virtual environment
deactivate