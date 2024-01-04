import sys
import os
import comtypes.client

def convert_ppt_to_pdf(input_file_path, output_file_path):
    # Convert file paths to Windows format
    input_file_path = os.path.abspath(input_file_path)
    output_file_path = os.path.abspath(output_file_path)

    # Create PowerPoint application object
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")

    # Set visibility to minimize
    powerpoint.Visible = 1

    try:
        # Open the PowerPoint slides
        slides = powerpoint.Presentations.Open(input_file_path)

        # Save as PDF (formatType = 32)
        slides.SaveAs(output_file_path, 32)
    finally:
        # Close the slide deck
        if 'slides' in locals():
            slides.Close()

if __name__ == "__main__":
    # Check if the correct number of command-line arguments is provided
    if len(sys.argv) != 3:
        print("Usage: python Convert.py input-file output-file")
        sys.exit(1)

    # Get console arguments
    input_file_path = sys.argv[1]
    output_file_path = sys.argv[2]

    # Convert PowerPoint to PDF
    convert_ppt_to_pdf(input_file_path, output_file_path)
