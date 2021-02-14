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


class PyGen(Xlsx):

    def get_header_row(self) -> object:
        """
        Gets header row number via user input.
        Adds row number to .headerrow attribute.

        Returns:
            self: PyGen Object.
        """
        try:
            while True:
                hr = input("\n Which Row Number contains the headers?: ")
                if hr.isdigit() and hr != '0':
                    self.headerrow = int(hr)
                    return self

        except Exception as e:
            input(f"\n Error on get_header_row: {e}")
            exit()


    def get_key_col(self) -> object:
        """
        Gets column letter containing data to be used as keys in generated dictionary.
        Adds column letter to .keycol attribute.

        Returns:
            self: PyGen Object.
        """
        try:
            while True:
                print("\n What Column Letter contains data to use as Keys? ")
                cl = input(" Enter column letter (Enter 0 to use row numbers): ")
                if cl == '0':
                    self.keycol = None
                    return self
                if cl.isalpha() and not cl.isdigit():
                    self.keycol = cl.upper()
                    return self

        except Exception as e:
            input(f"\n Error on get_key_col: {e}")
            exit()


    def get_data_cols(self) -> object:
        """
        Gets a COMMA-SEPARATED list (str) of column letters that contain the 
        data that will be added as the dictionary values in the data structure.
        Adds the Upper-Case list to the .dcols attribute.

        Returns:
            self: PyGen Object.
        """
        try:
            print("\n Enter a COMMA-SEPARATED list of Column Letters to use as values in the dictionary.")
            datacols = input(" : ")
            self.dcols = [x.upper().strip() for x in datacols.split(",")]
            return self

        except Exception as e:
            input("\n Error on get_data_cols: {e}")
            exit()

    def get_dict(self) -> object:
        """
        Uses Xlsx generate_dictionary method to generate a dictionary. 
        Adds the dictionary to .gendict attribute.

        Returns:
            self: PyGen object.
        """
        self.gendict = self.generate_dictionary(
            self.dcols, keycol=self.keycol, hdrrow=self.headerrow)
        return self


    def generate_py(self) -> None:
        """
        Generates "data.py" file containing a dictionary "DATA" of all 
        information read from the source file. Also adds the date generated 
        and the source filename as comments up top.
        """
        try:
            with open(f"{DLS}/data.py", "w") as f:
                f.write(f"# Generated: {TODAY.month}/{TODAY.day}/{TODAY.year}\n")
                f.write(f"# Source File: {self.path.name}\n\n")
                f.write(f"DATA = {pformat(self.gendict)}")
                print('\n "data.py" added to User\'s Downloads folder.')

        except Exception as e:
            input("\n Error on generate_py: {e}")
            exit()


if __name__ == "__main__":
    print('\n Drag and drop your source Excel ("*.xlsx") file below.')
    xl = PyGen(dropfile.get())
    xl.get_header_row().get_key_col().get_data_cols()
    xl.get_dict().generate_py()
