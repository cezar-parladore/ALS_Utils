import pandas as pd
import os
from glob import glob
import easygui

input_folder = easygui.diropenbox()
input_files_path = glob(input_folder+"/*.xlsx")
class ALS_file:
    # TODO path must be mandatory
    def __init__(self, path_to_file: str) -> None:
        self.path_to_file = path_to_file
        self.file_name = path_to_file.split("/")[-1].replace(".xlsx", "")
        self.worksheets = {
            "Cliente":'',
            "Coleta":'',
            "Metodos":'',
            "Resultados":'',
            "QAQC":'',
            "Dil":''
        }

    def get_worksheets(self, tables_only:bool=True):
        for key in self.worksheets.keys():
            if tables_only:
                self.worksheets[key] = pd.read_excel(
                    self.path_to_file,
                    sheet_name=key,
                    skiprows=9
                )
            else:
                self.worksheets[key] = pd.read_excel(
                    self.path_to_file,
                    sheet_name=key
                )

            #TODO implemment check if is corrupt; then call fix corrupt xls
    def check_qaqc(self):
        qaqc_ws = self.worksheets['QAQC']
        return qaqc_ws

xlsx_files = [ALS_file(xlsx).get_worksheets() for xlsx in input_files_path]

arquivo.get_worksheets()
arquivo.check_qaqc()

    #ANCHOR - def check_qaqc
    #ANCHOR - def to RESUMO

