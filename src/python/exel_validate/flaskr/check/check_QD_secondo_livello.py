import os
import pandas as pd
import numpy as np
import requests
import xlrd
from xlsxwriter.utility import xl_rowcol_to_cell
import yaml
import logging
import re
import openpyxl 
import sys
from pathlib import Path

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
f_handler = logging.FileHandler('generator.log', 'a+', 'utf-8')
c_handler = logging.StreamHandler()
formatter = logging.Formatter(
    '%(asctime)s - %(name)s - %(levelname)s - %(message)s')
f_handler.setFormatter(formatter)
c_handler.setFormatter(formatter)
logger.addHandler(f_handler)
logger.addHandler(c_handler)

ROOT_DIR = os.path.realpath(os.path.join(os.path.dirname(__file__), '..'))

class Check_QD():

    file_name = ""
    output_message = ""
    error_list = {}
    file_data = {}

    work_sheet = "" #sheet di lavoro di df_mapping
    work_codice_prestazione_siss = ""
    work_descrizione_prestazione_siss = ""
    work_codice_agenda_siss = ""
    work_casi_1_n = ""
    work_abilitazione_esposizione_siss = ""
    work_codici_disciplina_catalogo = ""
    work_descrizione_disciplina_catalogo = ""
    work_codice_QD = ""
    work_descrizione_QD = ""
    work_operatore_logico_QD = ""
    work_codice_metodica = ""
    work_descrizione_metodica = ""
    work_codice_distretto = ""
    work_descrizione_distretto = ""
    work_operatore_logico_distretto = ""
    work_priorita_U = ""
    work_priorita_primo_accesso_D = ""
    work_priorita_primo_accesso_P = ""
    work_priorita_primo_accesso_B = ""
    work_accesso_programmabile_ZP = ""

    work_index_codice_QD = 0
    work_index_codice_SISS_agenda = 0
    work_index_abilitazione_esposizione_SISS = 0
    work_index_codice_prestazione_SISS = 0
    work_index_operatore_logico_distretto = 0
    work_index_codici_disciplina_catalogo = 0
    work_index_codici_descri_disciplina_catalogo = 0

    '''def __init__(self, data, excel_file):
        #self.output_message = ""
        #with open("./flaskr/config_validator_PSM.yml", "rt", encoding='utf8') as yamlfile:
        #    data = yaml.load(yamlfile, Loader=yaml.FullLoader)
        #logger.debug(data)
        self.work_sheet = data[0]["work_column"]["work_sheet"] 
        self.work_codice_prestazione_siss = data[0]["work_column"]["work_codice_prestazione_siss"]
        self.work_descrizione_prestazione_siss = data[0]["work_column"]["work_descrizione_prestazione_siss"]
        self.work_codice_agenda_siss = data[0]["work_column"]["work_codice_agenda_siss"]
        self.work_casi_1_n = data[0]["work_column"]["work_casi_1_n"]
        self.work_abilitazione_esposizione_siss = data[0]["work_column"]["work_abilitazione_esposizione_siss"]
        self.work_codici_disciplina_catalogo = data[0]["work_column"]["work_codici_disciplina_catalogo"]
        self.work_descrizione_disciplina_catalogo = data[0]["work_column"]["work_descrizione_disciplina_catalogo"]
        self.work_codice_QD = data[0]["work_column"]["work_codice_QD"]
        self.work_descrizione_QD = data[0]["work_column"]["work_descrizione_QD"]
        self.work_operatore_logico_QD = data[0]["work_column"]["work_operatore_logico_QD"]
        self.work_codice_metodica = data[0]["work_column"]["work_codice_metodica"]
        self.work_descrizione_metodica = data[0]["work_column"]["work_descrizione_metodica"]
        self.work_codice_distretto = data[0]["work_column"]["work_codice_distretto"]
        self.work_descrizione_distretto = data[0]["work_column"]["work_descrizione_distretto"]
        self.work_operatore_logico_distretto = data[0]["work_column"]["work_operatore_logico_distretto"]
        self.work_priorita_U = data[0]["work_column"]["work_priorita_U"]
        self.work_priorita_primo_accesso_D = data[0]["work_column"]["work_priorita_primo_accesso_D"]
        self.work_priorita_primo_accesso_P = data[0]["work_column"]["work_priorita_primo_accesso_P"]
        self.work_priorita_primo_accesso_B = data[0]["work_column"]["work_priorita_primo_accesso_B"]
        self.work_accesso_programmabile_ZP = data[0]["work_column"]["work_accesso_programmabile_ZP"]

        self.work_index_sheet = data[1]["work_index"]["work_index_sheet"]
        self.work_index_codice_QD = data[1]["work_index"]["work_index_codice_QD"]
        self.work_index_op_logic_distretto = data[1]["work_index"]["work_index_op_logic_distretto"]
        self.work_index_codice_SISS_agenda = data[1]["work_index"]["work_index_codice_SISS_agenda"]
        self.work_index_abilitazione_esposizione_SISS = data[1]["work_index"]["work_index_abilitazione_esposizione_SISS"]
        self.work_index_codice_prestazione_SISS = data[1]["work_index"]["work_index_codice_prestazione_SISS"]
        self.work_index_operatore_logico_distretto = data[1]["work_index"]["work_index_operatore_logico_distretto"]
        self.work_index_codici_disciplina_catalogo = data[1]["work_index"]["work_index_codici_disciplina_catalogo"]
        self.work_index_codici_descri_disciplina_catalogo = data[1]["work_index"]["work_index_codici_descri_disciplina_catalogo"]
        '''

    def ck_QD_agenda(self, df_mapping, error_dict):
        print("start checking if foreach agenda there are the same QD")
        
        error_dict.update({'QD_II_livello': []})
        

        #catalogo_dir = "c:\\Users\\aless\\exel_validate\\src\\python\\exel_validate\\CCR-BO-CATGP#01_Codifiche attributi catalogo GP++_202007.xls"
        catalogo_dir = os.path.join(ROOT_DIR, 'CCR-BO-CATGP#01_Codifiche attributi catalogo GP++_110322.xls')
        sheet_QD = pd.read_excel(catalogo_dir, sheet_name='QD', converters={"Cod Disciplina": str})

        xfile = openpyxl.load_workbook(self.file_data) #recupero file excel da file system
        sheet = xfile.get_sheet_by_name(self.work_sheet) #recupero sheet excel
        QD_list = []
        for index, row in df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss] == "S":
                row_replace = row[self.work_codice_QD].replace(" ", "")
                QD_string = row_replace.split(",")
                for QD in QD_string:
                    if QD not in QD_list:
                        if QD != "":
                            print("check QD: " + QD)
                            short_sheet = sheet_QD.loc[sheet_QD["Codice Quesito"] == QD.strip()]
                            di_list = []
                            #QD_II = short_sheet.at[int(index),"Codice Quesito"]
                            #print("QD II: " + QD_II)
                            for ss in short_sheet["Tipo Quesito"].values:#Tipo Quesito
                                print("secondo livello: " + str(ss))
                                if ss == "II Livello" and QD not in QD_list:
                                    #trovato QD secondo livello
                                    #QD_II_level = short_sheet.at[int(index),"Codice Quesito"]
                                    #print("QD secondo livello: " + QD)
                                    error_dict['QD_II_livello'].append(QD)
                                    QD_list.append(QD)
                            if QD not in QD_list:
                                QD_list.append(QD)
        
        

        return error_dict

    def ck_QD_disciplina_agenda(self, df_mapping, sheet_QD, error_dict):
        print("start checking if foreach agenda there is the same Disciplina for all the QD")
        #tutti i QD di un agenda hanno la stessa disciplina
        error_dict.update({
            'error_QD_disciplina_agenda': [],
            'error_QD_descrizione_disciplina_agenda': [],
            'error_disciplina_mancante' : [],
            'error_discipline_agende_diverse': []
        }) 

        
        return error_dict

    def ck_QD_sintassi(self, df_mapping, error_dict):
        print("start checking if there is ',' separator between each QD defined")
        error_dict.update({
            'error_QD_caratteri_non_consentiti': [],
            'error_QD_spazio_bordi': [],
            'error_QD_spazio_internamente': [],
        })
        
        return error_dict

    def ck_QD_descrizione(self, df_mapping, sheet_QD, error_dict):
        print("start checking if there are the relative QD description")
        error_dict.update({
            'error_QD_descrizione': [],
            'error_QD_descrizione_space_bordo': [],
            'error_QD_descrizione_space_interno': []
        })

        
        return error_dict

    def ck_QD_operatori_logici(self, df_mapping, error_dict):
        print("start checking if there are the same logic op. for each agenda")
        error_dict.update({
                        'error_QD_operatori_logici': [],
                        'error_QD_operatori_logici_mancante': []})
        QD_dict_error = {}

        
        return error_dict