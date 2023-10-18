import re
import sys
from docx import Document
import pandas as pd

from subprocess import  Popen
LIBRE_OFFICE = r"soffice"

def convert_to_pdf(input_docx, out_folder):
    p = Popen([LIBRE_OFFICE, '--headless', '--convert-to', 'pdf', '--outdir',
               out_folder, input_docx])
    print([LIBRE_OFFICE, '--convert-to', 'pdf', input_docx])
    p.communicate()



# Check if Command Line Arguments are passed.
if len(sys.argv) < 3:
    print('Not Enough arguments where supplied')
    print(f'Use: {sys.argv[0]} <file_name.docx> <output_dir>')

    sys.exit()


sample_doc = sys.argv[1]
out_folder = sys.argv[2]


convert_to_pdf(sample_doc, out_folder)


print(f"spreadsheet = {spreadsheet}, template_name = {template_name}")

excel_data = pd.read_excel(spreadsheet)
cols = excel_data.columns
print(cols)
for index, row in excel_data.iterrows():
    file_obj = Document(template_name)
    #if index != 29: continue
    for c in cols:
        if c is not None:
            print(f"c='{c}' : {row[c]}")
            regex1 = re.compile(f"{c}")
            docx_replace_regex(file_obj, regex1, row[c])

    file_name= f"{index}_{row[cols[0]]}_{row[cols[1]]}.docx"
    print(file_name)
    file_obj.save(file_name)
    #exit()
