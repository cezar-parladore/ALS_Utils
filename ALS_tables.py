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

    def fMetodos(self):
        _ = self.worksheets["Metodos"]
        _.loc[0][0] = 'Ref'
        _ = _.rename(columns=_.iloc[0]).drop(_.index[0])
        _.dropna(axis=1, how='all', inplace=True)
        
        return _

    def fResultados(self):
        _ = self.worksheets["Resultados"]
        grupo = self.als_grupo
        # grupo = _.loc[0][0].split(': ')[1]
        _ = (_.drop([0], axis=0)
        .reset_index(drop=True)
        .transpose()
        .fillna("")
        )

        _[6] = _[0]+';'+_[1]+';'+_[2]+';'+_[3]+';'+_[4]+';'+_[5]

        _ = (_.transpose()
        .drop([0, 1, 2, 3, 4, 5])
        .reset_index(drop=True)
        )

        _.loc[0,1:3] = _.loc[0,1:3].str.replace(';','')
        _ = _.drop(labels=4,axis=1)
        _.loc[0][0] = 'Método de Análise'

        _ = _.rename(columns=_.iloc[0]).drop(_.index[0]).reset_index(drop=True)

        _ = _[_['Parâmetro']!='']
        _ = _.melt(id_vars=['Método de Análise','Parâmetro','Unidade','CAS'],
            var_name='melt',
            value_name='Resultado'
            )

        new_cols = _['melt'].str.split(';',expand=True).drop(labels=5,axis=1)
        new_cols.columns = ['lab_sample_id','Data de Coleta','Hora de Coleta','Matriz','Amostra']

        _ = pd.concat([_[['Método de Análise','Parâmetro','Unidade','CAS','Resultado']], new_cols], axis=1)
        
        _['Boletim'] = _['lab_sample_id'].str.split('-').str[0]
        _['Grupo de Amostras'] = grupo
        
        return _
    
    def fDil(self):
        _ = self.worksheets["Dil"]

        _ = _.rename(columns=_.iloc[0]).drop(_.index[0]).reset_index(drop=True)

        new_cols = _['Amostras'].str.split(';', expand=True).replace({None:''})
        _ = pd.concat([_, new_cols], axis=1).drop(columns='Amostras').rename(columns={'Parâmeto':'Parâmetro'})
    
        _ = pd.melt(_,
            id_vars=['Método de Análise','Parâmetro','Diluição'],
            var_name='melt',
            value_name='Boletim')

        _ = _[_['Boletim'] != ''].drop(columns=['melt'])
        _['Boletim'] = _['Boletim'].astype('str').str.replace('-','/')
        _['Parâmetro'] = _['Parâmetro'].str.split('@').str[1]

        return _

    # def fQaqc(self):






files = [ALS_file(path) for path in input_files_path]
files[1].fResultados()
files[1].worksheets["Resultados"]