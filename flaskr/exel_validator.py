import argparse
import itertools
import json
import logging
import random
import re
import time
from collections import OrderedDict
from functools import reduce
from pathlib import Path
from typing import Dict, List

#import openpyxl 
import pandas as pd
import numpy as np
import requests
import xlrd
from xlsxwriter.utility import xl_rowcol_to_cell
import yaml
#import xlsxwriter
#from openpyxl.utils import get_column_letter

from Vale_validator_check.Vale_validator import Validator
from check.check_QD import Check_QD
from check.check_metodiche import Check_metodiche
from check.check_distretti import Check_distretti
from check.check_priorita import Check_priorita
from check.check_univocita_prestazione import Check_univocita_prestazione

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


class Check_action():

    file_name = ""
    file_data = {}
    catalogo = OrderedDict()
    flag_check_list = []
    error_list = {}
    output_message = ""
    
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
    work_index_op_logic_distretto = 0
    work_index_codice_SISS_agenda = 0
    work_index_abilitazione_esposizione_SISS = 0
    work_index_codice_prestazione_SISS = 0
    work_index_operatore_logico_distretto = 0
    work_index_codici_disciplina_catalogo = 0


    def __init__(self):
        self.output_message = ""
        with open("./flaskr/config_validator.yml", "rt", encoding='utf8') as yamlfile:
            data = yaml.load(yamlfile, Loader=yaml.FullLoader)
        logger.debug(data)
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

    def import_file(self):
        logging.warning("import excel")

        template_file = input("Enter your mapping file.xlsx: ") ##insert mapping file
        print(template_file) 
        template_file = Path(template_file)
        self.file_name = template_file
        if template_file.is_file(): #C:\Users\aless\csi-progetti\FaqBot\faqbot-09112020.xlsx
            self.read_exel_file(template_file)
        else:
            print("Il file non esiste, prova a ricaricare il file con la directory corretta.\n")

    def read_exel_file(self, template_file):
        #pd.set_option("display.max_rows", None, "display.max_columns", None)
        df_mapping = pd.read_excel(template_file, sheet_name=self.work_sheet, converters={self.work_codici_disciplina_catalogo: str, self.work_codice_prestazione_siss: str}).replace(np.nan, '', regex=True)
        #print ("print JSON")
        #print(sh)
        
        catalogo_dir = "c:\\Users\\aless\\exel_validate\\CCR-BO-CATGP#01_Codifiche_attributi_catalogo GP++_201910.xls"

        sheet_QD = pd.read_excel(catalogo_dir, sheet_name='QD' )
        sheet_Metodiche = pd.read_excel(catalogo_dir, sheet_name='METODICHE', converters={"Codice SISS": str, "Codice Metodica": str})
        sheet_Distretti = pd.read_excel(catalogo_dir, sheet_name='DISTRETTI' )
        
        print("sheet_QD caricato\n")
        #print(sheet_QD)
        print("sheet_Metodiche caricato\n")
        #print(sheet_Metodiche)
        print("sheet_Distretti caricato\n")
        #print(sheet_Distretti)

        self.analizer(df_mapping, sheet_QD, sheet_Metodiche, sheet_Distretti)

    def analizer(self, df_mapping, sheet_QD, sheet_Metodiche, sheet_Distretti):

        print("FASE 0: precheck")
        self.check_column_name(df_mapping)

        print('Start analisys:\n', df_mapping)

        print("Fase 1") #FASE 1: CONTROLLO I QUESITI DIAGNOSTICI
        QD_error = self.check_qd(df_mapping, sheet_QD)
        #QD_error = {}
        print("Fase 2") #FASE 2: CONTROLLO LE METODICHE
        metodiche_error = self.check_metodiche(df_mapping, sheet_Metodiche)
        #metodiche_error = {}
        print("Fase 3") #FASE 3: CONTROLLO I DISTRETTI
        distretti_error = self.check_distretti(df_mapping, sheet_Distretti)
        #distretti_error = {}
        print("Fase 4") #FASE 4: CONTROLLO LE PRIORITA'
        priorita_error = self.check_priorita(df_mapping)
        #priorita_error = {}
        print("Fase 5") #FASE 5: CONTROLLO UNIVOCITA' PRESTAZIONI'
        univocita_prestazione_error = self.check_univocita_prestazione(df_mapping)
        #univocita_prestazione_error = {}
        print("Fase Vale Validator")
        catalogo_dir = "c:\\Users\\aless\\exel_validate\\CCR-BO-CATGP#01_Codifiche_attributi_catalogo GP++_201910.xls"
        wb = xlrd.open_workbook(catalogo_dir)
        sheet_QD_OW = wb.sheet_by_index(1)
        sheet_Metodiche_OW = wb.sheet_by_index(2)
        sheet_Distretti_OW = wb.sheet_by_index(3)
        QD_validator_error = {}
        metodiche_validator_error = {}
        distretti_validator_error = {}
        #QD_validator_error = Validator.ck_QD_description(self, df_mapping, sheet_QD_OW)
        #metodiche_validator_error = Validator.ck_metodiche_description(self, df_mapping, sheet_Metodiche_OW)
        #distretti_validator_error = Validator.ck_distretti_description(self, df_mapping, sheet_Distretti_OW)


        error_dict = {
            "QD_error": QD_error,
            "metodiche_error": metodiche_error,
            "distretti_error": distretti_error,
            "priorita_error": priorita_error,    
            "univocita_prestazione_error": univocita_prestazione_error,
            "QD_validator_error": QD_validator_error,
            "metodiche_validator_error": metodiche_validator_error,
            "distretti_validator_error": distretti_validator_error
        }

        self._validation(error_dict)

    def check_column_name(self, df_mapping):
        print("check the used column name of the excel file")

    def check_qd(self, df_mapping, sheet_QD):
        print("start checking QD") #Codice Quesito Diagnostico
        #controllo se per ogni Agenda sono inseriti gli stessi QD
        error_dict = {}
        

        error_QD_sintassi = Check_QD.ck_QD_sintassi(self, df_mapping, error_dict)
        error_QD_agenda = Check_QD.ck_QD_agenda(self, df_mapping, error_QD_sintassi)
        error_QD_disciplina_agenda = Check_QD.ck_QD_disciplina_agenda(self, df_mapping, sheet_QD, error_QD_agenda)
        error_QD_descrizione = Check_QD.ck_QD_descrizione(self, df_mapping, sheet_QD, error_QD_disciplina_agenda)
        error_QD_operatori_logici = Check_QD.ck_QD_operatori_logici(self, df_mapping, error_QD_descrizione)

        error_dict = error_QD_operatori_logici
        '''error_list = {
            "error_QD_agenda": error_QD_agenda,
            "error_QD_disciplina_agenda": error_QD_disciplina_agenda,
            "error_QD_sintassi": error_QD_sintassi,
            "error_QD_descrizione": error_QD_descrizione
        }'''

        return error_dict

    def check_metodiche(self, df_mapping, sheet_Metodiche):
        print("start checking Metodiche")
        error_dict = {}

        self.output_message = self.output_message + "\nErrori presenti nelle metodiche e riportate attraverso gli indici:\n"

        error_metodica_sintassi = Check_metodiche.ck_metodica_sintassi(self, df_mapping, error_dict)
        error_metodica_inprestazione = Check_metodiche.ck_metodica_inprestazione(self, df_mapping, sheet_Metodiche, error_metodica_sintassi)
        error_metodica_descrizione = Check_metodiche.ck_metodica_descrizione(self, df_mapping, sheet_Metodiche, error_metodica_inprestazione)

        error_dict = error_metodica_descrizione
        '''error_dict = {
            "error_metodica_inprestazione": error_metodica_inprestazione,
            "error_metodica_separatore": error_metodica_separatore,
            "error_metodica_descrizione": error_metodica_descrizione
        }'''

        #print("error_dict: %s", error_dict)
        return error_dict

    def check_distretti(self, df_mapping, sheet_Distretti):
        print("start checking Distretti")
        error_dict = {}

        error_distretti_sintassi = Check_distretti.ck_distretti_sintassi(self, df_mapping, error_dict)
        error_distretti_inprestazione = Check_distretti.ck_distretti_inprestazione(self, df_mapping, sheet_Distretti, error_distretti_sintassi)
        error_distretti_descrizione = Check_distretti.ck_distretti_descrizione(self, df_mapping, sheet_Distretti, error_distretti_inprestazione)
        error_distretti_operatori_logici = Check_distretti.ck_distretti_operatori_logici(self, df_mapping, error_distretti_descrizione)

        error_dict = error_distretti_operatori_logici
        '''error_dict = {
            "error_distretti_inprestazione": error_distretti_inprestazione,
            "error_distretti_separatore": error_distretti_separatore,
            "error_distretti_descrizione": error_distretti_descrizione
        }'''

        return error_dict

    def check_priorita(self, df_mapping):
        print("start checking priorità e tipologie di accesso")
        error_dict = {}

        error_prime_visite = Check_priorita.ck_prime_visite(self, df_mapping, error_dict)
        error_controlli = Check_priorita.ck_controlli(self, df_mapping, error_prime_visite)
        error_esami_strumentali =Check_priorita.ck_esami_strumentali(self, df_mapping, error_controlli)

        error_dict = error_esami_strumentali
        '''error_list = {
            "error_prime_visite": error_prime_visite,
            "error_controlli": error_controlli,
            "error_esami_strumentali": error_esami_strumentali
        }'''

        return error_dict

    def check_univocita_prestazione(self, df_mapping):
        print("start checking univocità delle prestazioni")
        error_dict = Check_univocita_prestazione.ck_casi_1n(self, df_mapping, {})

        return error_dict


    def _validation(self, error_dict):
        print("questi sono gli errori indivuduati e separati per categoria:\n %s", error_dict)
        
        '''df = pd.DataFrame(rows_list)
        with pd.ExcelWriter(self.file_name, engine='openpyxl', mode='a') as writer:
            df.to_excel(writer, sheet_name='new_mapping', index=False)'''
        print("\nPer osservare i risultati ottenuti, controllare il file prodotto: check_excel_result.txt")
        file = open("check_excel_result.txt", "w") 
        #file.write(json.dumps(error_dict)) 
        file.write(self.output_message + "\n" + json.dumps(error_dict))
        file.close() 

    def findCell(self, sh, searchedValue, start_col):
        result_coord = []
        result_value = []

        for row in range(sh.nrows):
            for col in range(start_col, sh.ncols):
                myCell = sh.cell(row, col)
                myValue = sh.cell(row, self.work_index_codice_prestazione_SISS)
                abilita = sh.cell(row, self.work_index_abilitazione_esposizione_SISS)
                if myCell.value == searchedValue and abilita.value == "S":
                    result_coord.append(str(row) + "#" + str(col))
                    result_value.append(myValue.value)
                    #return row, col#xl_rowcol_to_cell(row, col)

        if result_coord == []:
            return -1
        return result_coord, result_value

    def findCell_agenda(self, sh, searchedValue, start_col):
        result_coord = []
        result_value = []

        for row in range(sh.nrows):
            for col in range(start_col-1, start_col):
                myCell = sh.cell(row, col)
                myValue = sh.cell(row, self.work_index_codice_SISS_agenda-1) #Codice SISS agenda 15
                #abilita = sh.cell(row, self.work_index_abilitazione_esposizione_SISS-1) #abilitazione esposizione SISS 28
                if myCell.value == searchedValue: # and abilita.value == "S":
                    result_coord.append(str(row) + "#" + str(col))
                    result_value.append(myValue.value)
                    #return row, col#xl_rowcol_to_cell(row, col)

        if result_coord == []:
            return -1
        return result_coord#, result_value

    def update_list_in_dict(self, dictio, index, element):
        if index in dictio.keys():
            dictio[index].append(element)
        else:
            dictio[index] = [element]
        return dictio

k = Check_action()

k.import_file()
    