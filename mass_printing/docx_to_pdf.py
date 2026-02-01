import argparse
import os
import sys
from typing import List
import comtypes.client
import docx
from PyPDF2 import PdfMerger

def get_options(argv: List[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(usage="docx_to_pdf.py [-h] [-i <INPUT_FOLDER>] [-o <OUTPUT_FOLDER>]")

    parser.add_argument(
        "-i", "--input-folder",
        dest="inputFolder",
        action="store",
        required=False, 
        default=os.path.join(os.getcwd(), "."),
        help="Set the input folder. If not supplied, the <current folder> will be used as an input folder."
    )

    parser.add_argument(
        "-o", "--output-folder",
        dest="outputFolder",
        action="store",
        required=False, 
        default=os.path.join(os.getcwd(), "out"),
        help="Set the output folder. If not supplied, the <current folder>/out will be used as an output folder."
    )

    parser.add_argument(
        "-c", "--combine-output-file",
        dest="combineOutputFile",
        action="store",
        required=False, 
        default=os.path.join(os.getcwd(), "out.pdf"),
        help="Set the output combined PDF file. If not supplied, the <current folder>/out.pdf will be used as an output file."
    )


    options = parser.parse_args(argv)
    return options

def convert_doc_to_pdf( word_COM, word_path, pdf_path):
    print(f"\nInput file: '{word_path}'")
    docx_path = os.path.abspath(word_path)
    pdf_path = os.path.abspath(pdf_path)

    pdf_format = 17  # PDF file format code
    word_COM.Visible = False
    in_file = word_COM.Documents.Open(docx_path)
    in_file.SaveAs(pdf_path, FileFormat=pdf_format)
    in_file.Close()
    print(f"Output file: '{pdf_path}'\n")

def combinePdfFiles(file_in, combineOutputFile):
    merger = PdfMerger()
    if os.path.exists(combineOutputFile) and os.path.getsize(combineOutputFile) > 0:
        merger.append(combineOutputFile)
    merger.append(file_in)

    temp_output = combineOutputFile + ".tmp"
    with open(temp_output, "wb") as f_out:
        merger.write(f_out)
    merger.close()
    os.replace(temp_output, combineOutputFile)


def main(argv: List[str] = sys.argv[1:]) -> None:
    options = get_options(argv)

    outputFolder = os.path.abspath(options.outputFolder) 
    inputFolder =  os.path.abspath(options.inputFolder) 
    combineOutputFile = os.path.abspath(options.combineOutputFile)

    if not os.path.exists(outputFolder):
        os.mkdir(outputFolder)


    # List all files in the folder
    files = os.listdir(options.inputFolder)
    # Filter out only the .docx files
    docx_files = [file for file in files if file.endswith(".docx")]
    if docx_files:
        # Save the Word document as a PDF using Microsoft Word
        word = comtypes.client.CreateObject("Word.Application")
        for f in docx_files:
            file_in = os.path.join(inputFolder, f)
            file_out= os.path.join(outputFolder, f + ".pdf")
            convert_doc_to_pdf(word, file_in, file_out) 
            if(combineOutputFile):
                combinePdfFiles( file_out, combineOutputFile)

        # Quit Microsoft Word
        word.Quit()

if __name__ == "__main__":
    main()
