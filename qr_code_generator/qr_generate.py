import argparse
import io
import os
import sys
from pathlib import Path
from typing import List
import csv

_HTML_FILE_TOP_ = """
<html>
<head>
<link rel="stylesheet" href="qr_generate.css"/>
</head>
<body>
<table>
"""
_HTML_FILE_BOTTOM_ = """
</table>
</body>
</html>
"""

def get_options(argv: List[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    group = parser.add_mutually_exclusive_group()

    group.add_argument(
        "-c", "--coma-separated",
        dest="delimeter_coma",
        action="store_true", 
        help="Treat file content as a coma ',' separated text"
    )

    group.add_argument(
        "-s", "--space-separated",
        dest="delimeter_space",
        action="store_true", 
        help="Treat file content as a space ' ' separated text"
    )

    group.add_argument(
        "-t", "--tab-separated",
        dest="delimeter_tab",
        action="store_true", 
        help="Treat file content as a 'tab' separated text"
    )

    group.add_argument(
        "-p", "--pipe-separated",
        dest="delimeter_pipe",
        action="store_true", 
        help="Treat file content as a pipe '|' separated text"
    )

    parser.add_argument(
        "-i", "--input-file",
        dest="inputFile",
        action="store", 
        required=True, 
        help="Set the input file. This is a required argument." # If not supplied, the current folder will be used as a starting point."
    )

    parser.add_argument(
        "-o", "--output-folder",
        dest="outputFolder",
        action="store",
        required=False, 
        default=Path.joinpath(Path.cwd(), "out"),
        help="Set the output folder. If not supplied, the <current folder>/out will be used as an output folder."
    )

    options = parser.parse_args(argv)
    return options


def main(argv: List[str] = sys.argv) -> None:
    options = get_options(argv)
    inputFile = Path(options.inputFile)
    outputFolder = Path(options.outputFolder)

    if not Path.exists(outputFolder):
        Path.mkdir(outputFolder)

    if not Path.exists(inputFile):
        print(f"Error - input file '{inputFile}' must exist")
        return 

    separator = " " if options.delimeter_space else \
        "\t" if options.delimeter_tab else \
        "|" if options.delimeter_pipe else ","

    reader = csv.reader(open(inputFile), delimiter=separator)
    header = reader.__next__()
    cols = dict(zip(header, range(len(header))))

    os.system(f"cp qr_generate.css {outputFolder}")
    outputHtmlFile = open(Path.joinpath(outputFolder, f"{inputFile}.html"), "wt")
    outputHtmlFile.write(_HTML_FILE_TOP_)

    for row in reader:
        
        fileName = f"{row[cols['Name']]}.png"
        for ch in [" ", "\t", ",", ":", ";", "'", "#", "\""]:
            fileName = fileName.replace(ch, "_")

        outputFileName = Path.joinpath(outputFolder, fileName)
        s = f"qrencode -o {outputFileName} {row[cols['Url']]}"
        
        htmlStr = f""" 
    <tr>
        <td>
            <h2>{row[cols['Name']]} {row[cols["Code"]]}</h2>
            <a href='{row[cols['Url']]}'>{row[cols['Url']]}</a></td>"""
        if "Description" in cols:
            htmlStr += f"""
        <td>{row[cols["Description"]]}</td>"""
        htmlStr += f"""
        <td> <img src='{fileName}' style='height:200px;' /></td>
    </tr>"""
        outputHtmlFile.write(htmlStr)
        print(s)
        os.system(s)

    outputHtmlFile.write(_HTML_FILE_BOTTOM_)
    outputHtmlFile.close()


if __name__ == "__main__":
    main(["-i", "./for_qr.csv", "-c", "-o", "./art_items_out"])
    #main(["-i", "./donations.csv", "-c", "-o", "./donations_out"])
