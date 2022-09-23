import sys
import glob
import docx2txt
import PyPDF2
import re
import pandas as pd
import os
from typing import List


class PatternMatching:

    """Contains functions either related to the parsing of data for patterns relating to PII or the sanitation of
    such data """

    @staticmethod
    def __luhn(n: int) -> bool:

        """An algorithm to decipher whether a credit card number is legitimate or random.
        Returns True if legitimate, false if random"""

        r = [int(ch) for ch in str(n)][::-1]

        return (sum(r[0::2]) + sum(sum(divmod(d * 2, 10)) for d in r[1::2])) % 10 == 0

    @staticmethod
    def __censor(pii_list: List) -> List:

        """Censors each instance of PII with asterisks. This excludes dashes, spaces, and the last 4 characters"""

        pii_return = []

        for pii in pii_list:

            to_censor = pii[:-4]
            uncensored = pii[-4:]
            censored = []

            for char in to_censor:
                if char != "-":
                    char = "*"
                    censored.append(char)
                else:
                    censored.append(char)

            censored_word = "".join(censored) + uncensored
            pii_return.append(censored_word)

        return pii_return

    def ssn(self, data: str) -> List:

        """Uses regex to find patterns that match social security numbers. Examples include:
        111-11-1111
        111 11 1111
        111-111111"""

        pii_list = re.findall(r"\b\d{3}-\d{2}-\d{4}\b|\b\d{3}\s\d{2}\s\d{4}\b|\b\d{3}-\d{6}\b", data)

        return self.__censor(pii_list)

    def cc(self, data: str) -> List:

        """Uses regex to find patterns that match credit card numbers. Examples include:
        1111 1111 1111 1111
        1111111111111111
        1111-1111-1111-1111"""

        pii_list = re.findall(r"\b\d{4}\s\d{4}\s\d{4}\s\d{4}\b|\b\d{16}\b|\b\d{4}-\d{4}-\d{4}-\d{4}\b", data)
        pii_clean = []
        for pii in pii_list:
            pii = pii.replace("-", "")
            pii = pii.replace(" ", "")
            if self.__luhn(pii):
                pii_clean.append(pii)

        return self.__censor(pii_clean)


class DataExtract:

    """Extracts data within files from the following file types:
    Word
    CSV
    Excel
    Text
    PDF"""

    @staticmethod
    def word(directory: str) -> str:

        """Extracts the data within a given Word document"""

        try:
            data = docx2txt.process(directory)
            return data
        except Exception as e:
            print(e)

    @staticmethod
    def csv(directory: str) -> str:

        """Extracts the data within a given CSV document"""

        try:
            data = str(pd.read_csv(directory, encoding="latin1"))
            return data
        except Exception as e:
            print(e)

    @staticmethod
    def excel(directory: str) -> List[str]:

        """Extracts the data within a given Excel document"""

        try:
            spreadsheet = pd.ExcelFile(directory, engine="openpyxl")
            data_list = []
            for sheet in spreadsheet.sheet_names:
                data = str(spreadsheet.parse(sheet))
                data_list.append(data)
            return data_list
        except Exception as e:
            print(e)

    @staticmethod
    def text(directory: str) -> List:

        """Extracts the data within a given Text document"""

        try:
            data_list = []
            with open(directory, mode='r', encoding="latin-1") as f:
                for line in f:
                    data_list.append(line)
                return data_list
        except Exception as e:
            print(e)

    @staticmethod
    def pdf(directory: str) -> List:

        """Extracts the data within a given PDF document"""

        try:
            with open(directory, mode='rb') as file:
                reader = PyPDF2.PdfFileReader(file)
                data_list = []
                for page in reader.pages:
                    pdf_text = page.extractText()
                    pdf_text = pdf_text.replace('\n', '')
                    data_list.append(pdf_text)
                return data_list
        except Exception as e:
            print(e)


class Scanner:

    """Iterates recursively over a given folder's files
    Extracts data from the files differently based on the files extension or skips the file all together
    Scans the extracted data for patterns matching Personal identifiable Information
    Outputs the metadata into the given output CSV file"""

    def __init__(self, scan_path, output_path):
        self.scan_path = scan_path
        self.output_path = output_path

        self.PatternMatching = PatternMatching()

    def __output(self, pii_type: str, location: str, pii: str):

        """Appends metadata on PII to the given CSV file"""

        pii_list = []
        pii_type_list = []
        location_list = []
        pii_type_list.append(pii_type)
        location_list.append(location)
        pii_list.append(pii)
        df = pd.DataFrame(data={"PII Type": pii_type_list, "File Path": location_list, "PII": pii_list})

        if os.path.isfile(self.output_path):
            df.to_csv(self.output_path, sep=',', index=False, mode='a', header=False)
        else:
            df.to_csv(self.output_path, sep=',', index=False)

    def __ssn_process(self, data: str, filename: str):

        """Processes Social Security Numbers from the data and prepares the data for output"""

        ssn_pii = self.PatternMatching.ssn(str(data))

        for pii in ssn_pii:
            pii_type = "Social Security Number"
            self.__output(pii_type, filename, pii)

    def __cc_process(self, data: str, filename: str):

        """Processes Credit Cards from the data and prepares the data for output"""

        cc_pii = self.PatternMatching.cc(str(data))

        for pii in cc_pii:
            pii_type = "Credit Card Number"
            self.__output(pii_type, filename, pii)

    @staticmethod
    def __extract_by_extension(filepath: str) -> str:

        """Checks what the file's extension is and extracts the data from the file based on the file type"""

        data = ""

        if filepath.endswith(".docx") or filepath.endswith(".doc"):
            data = DataExtract.word(filepath)

        elif filepath.endswith(".xlsx") or filepath.endswith(".xlx"):
            data = DataExtract.excel(filepath)

        elif filepath.endswith(".pdf"):
            data = DataExtract.pdf(filepath)

        elif filepath.endswith(".txt"):
            data = DataExtract.text(filepath)

        elif filepath.endswith(".csv"):
            data = DataExtract.csv(filepath)

        return data

    def run(self):

        """Recursively iterates over the given folder's files and then sorts the files by extension"""

        for filename in glob.iglob(f'{self.scan_path}\\**', recursive=True):
            print(filename)
            data = self.__extract_by_extension(filename)

            if data:
                self.__ssn_process(data, filename)
                self.__cc_process(data, filename)

        print("*** Scan is Complete ***")


if __name__ == '__main__':

    """When the script is run 2 arguments must be given. The folder that you wish to scan,
     and the path to the CSV file you wish to output the results to."""

    scan = sys.argv[1]
    output = sys.argv[2]

    scanner = Scanner(scan, output)
    scanner.run()


