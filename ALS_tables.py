import pandas as pd
import os
from glob import glob
import easygui

input_folder = easygui.diropenbox()
input_files_path = glob(input_folder + "/*.xlsx")


class ALS_file:
    def __init__(self, path_to_file: str, skip_blank_rows: bool = True) -> None:
        self.path_to_file = path_to_file
        self.skip_rows = 9
        if not skip_blank_rows:
            self.skip_rows = 0
        self.file_name = path_to_file.split("/")[-1].replace(".xlsx", "")
        self.worksheets = {
            "Cliente": "",
            "Coleta": "",
            "Metodos": "",
            "Resultados": "",
            "QAQC": "",
            "Dil": "",
        }
        for key in self.worksheets.keys():
            #FIXME raise error if corrupt, then call function to convert corrupt xls to xlsx
            self.worksheets[key] = pd.read_excel(
                self.path_to_file, sheet_name=key, skiprows=self.skip_rows
            )


    def check_qaqc(self):
        qaqc_ws = self.worksheets["QAQC"]
        return qaqc_ws



xlsx_files = [ALS_file(xlsx) for xlsx in input_files_path]
xlsx_files[0].worksheets["Resultados"]

arquivo.get_worksheets()
arquivo.check_qaqc()

# ANCHOR - def check_qaqc
# ANCHOR - def to RESUMO
