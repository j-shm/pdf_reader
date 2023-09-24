from PIL import Image
import numpy as np
from pdf2image import convert_from_path
import os
import glob
import pytesseract

def GetPdf():
    pdfs = []
    for file in glob.glob("*.pdf"):
        pdfs.append(file)
    return pdfs

def ProcessPdf(pdf):
    pdf_image = GetImageFromPdf(pdf)
    pdf_image.show()
    print("pdf: " + pdf)
    print(ExtractEmail(pdf_image))
    print(ExtractName(pdf_image))
    print("-=-")


def ExtractName(pdf_image):

    black_crop = CropFromTop(pdf_image, [0,0,0], offset=-20)

    blue_crop = CropFromTop(black_crop, [67,103,132])

    white_crop = CropFromTop(blue_crop, [255,255,255],offset=3)

    temp_crop_width, temp_crop_height = white_crop.size
    temp_crop = white_crop.crop((0, 30, temp_crop_width,temp_crop_height))

    top_x,top_y = GetFirstOccurrenceOfColour(temp_crop, [67,103,132])
    pdf_image_width, pdf_image_height = white_crop.size
    pdf_image_vertically_cropped = white_crop.crop((0, 0, pdf_image_width,top_y))

    pdf_image_cropped_right = CropFromRightLast(pdf_image_vertically_cropped, [67,103,132],offset=-40)

    pdf_image_cropped_left = CropFromLeft(pdf_image_cropped_right, [67,103,132],offset=2)
    pdf_image_cropped_left.show()
    name = pytesseract.image_to_string(pdf_image_cropped_left)

    if "\n" in name:
        return name.split("\n")[0].strip()
    return name.strip()

def CropFromTop(pdf_image, colour, offset=0):
    top_x,top_y = GetFirstOccurrenceOfColour(pdf_image, colour)
    pdf_image_width, pdf_image_height = pdf_image.size
    pdf_image_cropped = pdf_image.crop((0, top_y+offset, pdf_image_width,pdf_image_height))
    return pdf_image_cropped

def CropFromBottom(pdf_image, colour,offset=0):
    bottom_x,bottom_y = GetFirstOccurrenceOfColour(pdf_image, colour)
    pdf_image_width, pdf_image_height = pdf_image.size
    pdf_image_cropped = pdf_image.crop((0, 0, pdf_image_width,bottom_y))
    return pdf_image_cropped


def CropFromLeft(pdf_image, colour,offset=0):
    bottom_x,bottom_y = GetLastOccurrenceOfColour(pdf_image, colour)
    pdf_image_width, pdf_image_height = pdf_image.size
    pdf_image_cropped = pdf_image.crop((bottom_x+offset, 0, pdf_image_width,pdf_image_height))
    return pdf_image_cropped

def CropFromRightLast(pdf_image, colour,offset=0):
    bottom_x,bottom_y = GetLastOccurrenceOfColour(pdf_image, colour)
    pdf_image_width, pdf_image_height = pdf_image.size
    pdf_image_cropped = pdf_image.crop((0, 0, bottom_x+offset,pdf_image_height))
    return pdf_image_cropped

def ExtractEmail(pdf_image):
    #first we need to crop the image to the top of the email address box
    pdf_image_half_cropped = CropFromTop(pdf_image, [223,225,232])
    pdf_image_cropped = CropFromBottom(pdf_image_half_cropped, [67,103,132])
    pdf_image_cropped.show()
    email = pytesseract.image_to_string(pdf_image_cropped)

    return email.split(":")[1].split("VAT")[0].strip()
    
def GetFirstOccurrenceOfColour(pdf_image, colour):
    pdf_image_array = np.array(pdf_image)
    y, x = np.where(np.all(pdf_image_array==colour,axis=2))
    return (x[0], y[0])

def GetLastOccurrenceOfColour(pdf_image, colour):
    pdf_image_array = np.array(pdf_image)
    y, x = np.where(np.all(pdf_image_array==colour,axis=2))
    return (x[-1], y[-1])

def GetImageFromPdf(pdf):
    return convert_from_path(pdf, poppler_path=os.getcwd() + "/poppler/Library/bin")[0]

if __name__ == '__main__':
    print("pdf reader: colour method!")
    for pdf in GetPdf():
        print("-=-")
        ProcessPdf(pdf)