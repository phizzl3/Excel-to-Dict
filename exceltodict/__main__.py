#!/usr/bin/env python3

"""
Gets an Excel ("*.xlsx") file via drag-and-drop and creates a .py file 
containing a dictionary "DATA" generated from the info contained in 
the xlsx and saves it to the User's Downloads folder as "data.py".
"""

from datetime import datetime
from pathlib import Path
from pprint import pformat

import dropfile
from xlclass import Xlsx

DLS = Path().home() / "Downloads"
TODAY = datetime.today()


def get_header_row() -> int:
    """
    Gets header row number via user input.

    Returns:
        int: Row number containing headers.
    """
    try:
        while True:
            hr = input("\n Which Row Number contains the headers?: ")
            if hr.isdigit() and hr != '0':
                return int(hr)
    except Exception as e:
        input(f"\n Error on get_header_row: {e}")
        exit()


def get_key_col() -> str:
    """
    Gets column letter containing data to be used as keys in generated dictionary.

    Returns:
        str: Column letter containing data for dictionary keys.
    """
    try:
        while True:
            cl = input("\n What Column Letter contains data to use as Keys: ")
            if cl.isalpha() and not cl.isdigit():
                return cl.upper()
    except Exception as e:
        input(f"\n Error on get_key_col: {e}")
        exit()


def get_data_cols() -> list:
    """
    Gets a COMMA-SEPARATED list (str) of column letters that contain the 
    data that will be added as the dictionary values in the data structure.

    Returns:
        list: List of uppercase column letters containing needed data.
    """
    try:
        print("\n Enter a COMMA-SEPARATED list of Column Letters to use as values in the dictionary.")
        datacols = input(" : ")
        return [x.upper().strip() for x in datacols.split(",")]
    except Exception as e:
        input("\n Error on get_data_cols: {e}")
        exit()


def generate_py(gendict, xl) -> None:
    """
    Generates "data.py" file containing a dictionary "DATA" of all 
    information read from the source file. Also adds the date generated 
    and the source filename as comments up top.

    Args:
        gendict (dict): Dictionary of information read from source file.
        xl (Xlsx): Xlsx object generated from source Excel file.
    """
    try:
        with open(f"{DLS}/data.py", "w") as f:
            f.write(f"# Generated: {TODAY.month}/{TODAY.day}/{TODAY.year}\n")
            f.write(f"# Source File: {xl.path.name}\n\n")
            f.write(f"DATA = {pformat(gendict)}")
            print('\n "data.py" added to User\'s Downloads folder.')
    except Exception as e:
        input("\n Error on generate_py: {e}")
        exit()


if __name__ == "__main__":
    print('\n Drag and drop your source Excel ("*.xlsx") file below.')
    xl = Xlsx(dropfile.get())
    headerrow = get_header_row()
    keycol = get_key_col()
    dcols = get_data_cols()
    gendict = xl.generate_dictionary(
        keycol, dcols, hdrrow=headerrow, datastartrow=headerrow+1)
    generate_py(gendict, xl)
