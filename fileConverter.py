#!ONLY WORKS WITH WIN32
# import win32com.client 
# import os
# type_in_file = input("What kind of file would you like to convert? (ppt, docx, pdf, jpg, etc.)\n")
# type_out_file = input("What kind of file would you like to output? (ppt, docx, pdf, jpg, etc.)\n")
# in_file = input("Enter the path of the file:\n")
# out_file = os.path.splitext(in_file)

# if type_in_file=="ppt" and type_out_file=="pdf":
#     powerpoint=win32com.client.Dispatch("Powerpoint.Application")
#     pdf = powerpoint.Presentations.open(in_file)
#     pdf.SaveAs(out_file, 32)
#     pdf.Close()
#     powerpoint.Quit()

# else:
#     print("Error occurred, or feature is not supported yet.")
#---------------------------


#%% Convert a Folder of PowerPoint PPTs to PDFs

# Purpose: Converts all PowerPoint PPTs in a folder to Adobe PDF

# Author:  Matthew Renze

# Usage:   python.exe ConvertAll.py input-folder output-folder
#   - input-folder = the folder containing the PowerPoint files to be converted
#   - output-folder = the folder where the Adobe PDFs will be created

# Example: python.exe ConvertAll.py C:\InputFolder C:\OutputFolder

# Note: Also works with PPTX file format

#%% Import libraries
import sys
import os
from comtypes.client import CreateObject
import os
import time

#%% Get console arguments
input_folder_path = sys.argv[1]
output_folder_path = sys.argv[2]

#%% Convert folder paths to Windows format
input_folder_path = os.path.abspath(input_folder_path)
output_folder_path = os.path.abspath(output_folder_path)

#%% Get files in input folder
input_file_paths = os.listdir(input_folder_path)

#%% Convert each file
for input_file_name in input_file_paths:

    # Skip if file does not contain a power point extension
    if not input_file_name.lower().endswith((".ppt", ".pptx")):
        continue
    
    # Create input file path
    input_file_path = os.path.join(input_folder_path, input_file_name)
        
    # Create powerpoint application object
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    
    # Set visibility to minimize
    powerpoint.Visible = 1
    
    # Open the powerpoint slides
    slides = powerpoint.Presentations.Open(input_file_path)
    
    # Get base file name
    file_name = os.path.splitext(input_file_name)[0]
    
    # Create output file path
    output_file_path = os.path.join(output_folder_path, file_name + ".pdf")
    
    # Save as PDF (formatType = 32)
    slides.SaveAs(output_file_path, 32)
    
    # Close the slide deck
    slides.Close()
    # powerpoint.QueryInterface(quit())    

    #! THIS WILL CLOSE OUT ANY ITERATION OF POWERPOINT OPEN - DO NOT USE IF YOU HAVE AN UNSAVED PPT OPEN
    try:
        os.system('TASKKILL /F /IM POWERPNT.exe')
    except:
        print("Error")
 
