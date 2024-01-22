import sys
import os
import comtypes.client

def convert_pptx_to_pdf(input_file_path, output_file_path):
    # Convert file paths to Windows format
    input_file_path = os.path.abspath(input_file_path)
    output_file_path = os.path.abspath(output_file_path)

    # Create PowerPoint application object
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")

    # Set visibility to minimize
    powerpoint.Visible = 1

    # Open the PowerPoint slides
    slides = powerpoint.Presentations.Open(input_file_path)

    # Save as PPTX (formatType = 32)
    slides.SaveAs(output_file_path, 32)

    # Close the slide deck
    slides.Close()

    powerpoint.Quit()