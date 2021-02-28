from PyPDF2 import PdfFileMerger
import os
# pdf_folder = ['./pdfs/channels-converted.pdf', './pdfs/effectiveC-converted.pdf']
pdf_folder = os.listdir("./pdfs/")

merger = PdfFileMerger()

for pdf in pdf_folder:
    input_file_path = os.path.join("./pdfs/", pdf)
    merger.append(input_file_path)

merger.write("result.pdf")
merger.close()



#! TRYING TO GET THIS TO WORK IF GIVEN CURRENT DIRECTORY INSTEAD OF MANUALLY TYPING EACH FILE NAME TO MERGE
# import os
# from PyPDF2 import PdfFileMerger

# pdf_folder = os.listdir("./pdfs/")

# merger = PdfFileMerger()

# for pdf in pdf_folder:
#     # Skip if file does not contain a pdf extension
#     if not pdf.lower().endswith(".pdf"):
#         continue

#     # Create input file path
#     input_file_path = os.path.join("./pdfs/", pdf)

#     # Get base file name
#     file_name = os.path.splitext(pdf)[0]

#     merger.append(file_name)

# # Create output file path
# output_file_path = os.path.join(pdf_folder, "result.pdf")

# merger.write(output_file_path)
# merger.close()