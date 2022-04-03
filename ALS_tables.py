from calendar import c
import pandas as pd
import os
from glob import glob


input_files = '/home/cparladore/Documentos/01_pyscripts/ALS_to_Tabular/01_inputs/xlsx/'

class ALS:
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



files = [ALS(input_files+arquivo) for arquivo in glob(f'{input_files}*')]#
print(files[0])
for arquivo in os.listdir(input_files):

    
glob(input_files)



worksheets = {
            "Cliente":'',
            "Coleta":'',
            "Metodos":'',
            "Resultados":'',
            "QAQC":'',
            "Dil":''
        }
