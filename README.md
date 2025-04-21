# PPT Automator

A PowerPoint automation tool built with FastMCP that allows you to programmatically update PowerPoint slides with text in rectangles. This tool is perfect for automating slide updates, batch processing presentations, and integrating PowerPoint modifications into your workflow.

## Features

- Update text in existing rectangle shapes on PowerPoint slides
- Automatically create new rectangles if none exist (with customizable dimensions)
- Simple API interface through FastMCP
- Batch processing capabilities
- Maintains original presentation formatting
- Non-destructive updates to existing slides
- Error handling and validation

## Prerequisites

- Python 3.x (3.7 or higher recommended)
- Windows, macOS, or Linux operating system
- Required packages:
  - `python-pptx`: For PowerPoint file manipulation
  - `fastmcp`: For API interface and communication
  - `typing`: For type hints and better code documentation

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/ppt-automator.git
   cd ppt-automator
   ```

2. Create and activate a virtual environment (recommended):
   ```bash
   python -m venv venv
   # On Windows
   .\venv\Scripts\activate
   # On macOS/Linux
   source venv/bin/activate
   ```

3. Install dependencies:
   ```bash
   pip install python-pptx fastmcp
   ```

## Detailed Usage

### Basic Example
