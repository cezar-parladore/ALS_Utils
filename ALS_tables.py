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
                self.path_to_file, sheet_name=key, skiprows=self.skip_rows, header=None
            )
        self.als_grupo = self.worksheets["Cliente"].iat[1,1]


    def check_qaqc(self):
        _ = self.worksheets["QAQC"]
        _ = _.dropna(axis=0,how='all')
        type_rows = _.loc[:,1:].isna().all(1)
        types = _[[0]].loc[type_rows][0].tolist()
        return types

    def fColeta(self):
        _ = self.worksheets["Coleta"]
        _ = _.rename(columns=_.iloc[0]).drop(_.index[0]).reset_index(drop=True)
        _ = _.drop(columns=['Tipo de Amostra'])
        
        return _    
    # def fix_worksheets(self):

    
