# Image to Excel Color Art
thanks to jadi❤️(https://github.com/jadijadi)

This Python script converts any image into a piece of **"pixel art"** within an Excel spreadsheet.  
It works by analyzing the image in 10x10 pixel blocks, calculating the average color for each block, and filling a corresponding cell in an Excel sheet with that color.

The result is a down-scaled, pixelated representation of your original image, viewable directly in Microsoft Excel or any compatible spreadsheet program.

---

## How It Works

The script performs the following steps:

1. **Opens an Image**  
   Uses the Pillow (PIL) library to open a source image file (`sample_image.png`).  
   Ensures the image is in RGB format, converting it if necessary.

2. **Divides the Image into Blocks**  
   The script divides the image into a grid of 10x10 pixel blocks.

3. **Calculates Average Color**  
   For each block, it iterates through all the pixels, sums their **Red, Green, and Blue** values,  
   then calculates the average RGB value for that block.

4. **Creates an Excel Sheet**  
   Using the `openpyxl` library, it creates a new Excel workbook.

5. **Fills Cells with Color**  
   - Maps each image block to a cell in the worksheet.  
   - Converts the calculated average RGB color to a hexadecimal code.  
   - Sets the cell's background fill and writes the hex code as the cell’s value.

6. **Formats the Sheet**  
   Adjusts row height and column width to make cells appear as small squares, creating a pixel grid.

7. **Saves the File**  
   The final spreadsheet is saved as **`average_colors.xlsx`**.

---

## Requirements

You need **Python 3** and the following libraries:

- **Pillow**: For image processing  
- **openpyxl**: For writing to `.xlsx` Excel files

