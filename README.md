# Excel Paste Handler

A simple web application that captures and displays data pasted from Excel.

## How to Use

1. Open `index.html` in a web browser
2. Copy data from Excel (select cells and press Ctrl+C)
3. Click on the designated paste area in the web app
4. Press Ctrl+V to paste the data
5. The app will automatically format and display the pasted data in a table

## Features

- Captures both HTML and plain text data from Excel
- Preserves formatting when pasting HTML content
- Supports merged cells (rowspan and colspan)
- Maintains cell styling from Excel
- Handles images pasted from Excel
- Automatically formats the data into a table
- Responsive design that works on all screen sizes
- First row is treated as headers

## Files

- `index.html` - The main HTML file
- `styles.css` - CSS for styling the application
- `script.js` - JavaScript code to handle the paste event and process the data

## Technical Details

The application works by capturing the paste event and processing the clipboard data. It first checks for images in the clipboard, then tries to get HTML data, and if HTML data is not available, it falls back to plain text (tab-separated values).

For HTML data, the application preserves all structure including:
- Merged cells (using rowspan and colspan attributes)
- Cell styling (colors, borders, alignment)
- Text formatting (bold, italic, etc.)
- Images (both embedded base64 images and references)

### Image Support

The application handles images in two ways:
1. Direct image paste - When an image is copied and pasted directly, it's processed as a blob and displayed in the table
2. Images in HTML - When HTML containing images is pasted, the application attempts to display them:
   - Base64 encoded images are displayed normally
   - File references (file://) are replaced with placeholders due to security restrictions

This allows the application to handle both formatted and plain text data from Excel while maintaining the original layout and appearance. 