"""
Inspired by Matt Parker's YouTube video where he showed how image pixels can be expressed as Excel cells.
YouTube video: https://www.youtube.com/watch?v=UBX2QQHlQ_I

Description: This is a simple program that takes a user defined image file, optionally shrinks it, and then saves it
             in an Excel file at the same location of the original image's location.

Feature: User supplied image (.jpg or .png files) is converted from RGB to HEX pixels.
         Individual cells of Excel sheet are treated as pixels of the image, and openpyxl.styles.PatternFill is used to
         color these cells according to the values of pixel colors.
         Optinally, user can resize (shrink or expand) the image before exporting it to excel file as well.
         Resulting Excel file is saved at the same location as original file.

Purpose: This program is provided as an example of usage of these functions.
"""

# ===================================================================================================|
# Imports: Functions imports the following libraries. So ensure they are installed before using them.|
# import os                                                                                          |
# import PIL                                                                                         |
# import openpyxl                                                                                    |
# ===================================================================================================|

from img2xl import img2xl

# Manual Inputs: Only the name of the image path
img_src = r'assets/dog.jpg'

# Calling functions from img2xl
img2xl(img_src, True, 0.2)

# User can find the resulting excel file at the same location as the original image file's location