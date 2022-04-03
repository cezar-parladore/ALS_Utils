import pandas as pd
import os
from glob import glob
import easygui

input_folder = easygui.diropenbox()
input_files = glob(input_folder+"/*.xlsx")

class ALS_file:
    # TODO path must be mandatory
    def __init__(self, path_to_file: str) -> None:
        self.path_to_files = path_to_file
        self.file_name = path_to_file.split("/")[-1]
        self.worksheets = {
            "Cliente":'',
            "Coleta":'',
            "Metodos":'',
            "Resultados":'',
            "QAQC":'',
            "Dil":''
        }
        pass


    def __str__(self) -> str:
        return self.file_name


    def get_worksheets(self):
        for key in self.worksheets.keys():
            print(key)
            self.worksheets[key] = pd.read_excel(
                self.path_to_file,
                sheet_name=key
            )
            #TODO implemment check if is corrupt; then call fix corrupt xls



files_dict ={}

for file in files

    
glob(input_files)



worksheets = {
            "Cliente":'',
            "Coleta":'',
            "Metodos":'',
            "Resultados":'',
            "QAQC":'',
            "Dil":''
        }
