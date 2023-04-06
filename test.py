import pytesseract
from PIL import Image

# Open the image file
image = Image.open('2023-04-05_14-21-23/CR_12345_BR_1_US_2_TC_2.png')

# Perform OCR using pytesseract
text = pytesseract.image_to_string(image)

# Print the recognized text
print(text)