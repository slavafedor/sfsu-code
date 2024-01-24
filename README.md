# sfsu-code

Set of Python scripts to help with mass mailing and other work


## docx_replace.py
Used for replacing {TAGS} in Template (.docx) files and creating resulting file(s)

### usage: 
	docx_replace.py [-h] -t TEMPLATEFILE -d DATAFILE [-b BEGININDEX] [-e ENDINDEX] [-o OUTPUTFOLDER] [-f]

	options:
	-h, --help            show this help message and exit
	-t TEMPLATEFILE, --template-file TEMPLATEFILE
							Set the input data file. This is a required argument.
	-d DATAFILE, --data-file DATAFILE
							Set the input data file. This is a required argument.
	-b BEGININDEX, --begin-index BEGININDEX
							Set the input data file starting data row index (starts with 0).
	-e ENDINDEX, --end-index ENDINDEX
							Set the input data file ending data row index (starts with 0).
	-o OUTPUTFOLDER, --output-folder OUTPUTFOLDER
							Set the output folder. If not supplied, the <current folder>/out will be used as an output folder.
	-f, --first-only      If set - the only first match will be replaced with value, and only one file will be created.

### Examples:

Create Certificates based on MS Word file `cert_landscape.docx` using data from the `rd-wo-addr.xlsx` Excel data file, upt to 4th row (-e 3) row index starts with 0:

	python docx_replace.py -t cert_landscape.docx -d rd-wo-addr.xlsx -e 3

Create SFSU 'from' labels. Here we are using the `SFSU_yellow_addr_labels_5160.docx` as a template and data file `first-last-addresses.xlsx` starting row index 49 (`-b 49`) and w/o `-f` flag which means ALL the template tags will be replaced with the values of one data row:

	python docx_replace.py -t ./templates/SFSU_yellow_addr_labels_5160.docx -d ./data/first-last-addresses.xlsx -o ./out/ -b 49

Create 'to' labels. Here we are using the `addr_labels_5160.docx` as a template and data file `first-last-addresses.xlsx` ending row index 29 (`-e 29`) and w/o `-f` flag which means only first tag in the template will be replaced with the values of one data row:

	python docx_replace.py -t ./templates/addr_labels_5160.docx -d ./data/first-last-addresses.xlsx -o ./out/ -e 29 -f
