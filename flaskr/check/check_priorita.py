import pandas as pd
import numpy as np
import requests
import xlrd
from xlsxwriter.utility import xl_rowcol_to_cell
import yaml
import logging
import re

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

class Check_priorita():

    file_name = ""
    output_message = ""
    error_list = {}

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
        with open("./flaskr/config_validator_PSM.yml", "rt", encoding='utf8') as yamlfile:
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

    
    def ck_prime_visite(self, df_mapping, error_dict): 
        #Descrizione Prestazione SISS
        print("start checking if prestazione prime visite is correct")
        error_dict.update({'error_prime_visite': []})
        
        str_check = "PRIMA VISITA"
        for index, row in df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss] == "S":
                if str_check in row[self.work_descrizione_prestazione_siss]:
                    logging.info("Prestazione PRIMA VISITA da controllare all'indice: " + str(int(index)+2))
                    if row[self.work_priorita_U] == "N" and row[self.work_priorita_primo_accesso_D] == "N" and row[self.work_priorita_primo_accesso_P] == "N" and row[self.work_priorita_primo_accesso_B] == "N": 
                        logging.error("trovato anomalia in check prime visite all'indice: " + str(int(index)+2))
                        error_dict["error_prime_visite"].append(str(int(index)+2))
        
        out1 = ", \n".join(error_dict['error_prime_visite'])
        self.output_message = self.output_message + "\nerror_prime_visite: \n" + "at index: \n" + out1
        
        return error_dict

    '''Nel caso di prestazione visita di controllo, verificare se è presente il campo Accesso programmabile ZP'''
    def ck_controlli(self, df_mapping, error_dict):
        print("start checking if prestazione controlli is correct")
        error_dict.update({'error_controlli': []})

        str_check = "CONTROLLO"
        for index, row in df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss] == "S":
                if str_check in row[self.work_descrizione_prestazione_siss]:
                    logging.info("Prestazione CONTROLLO da controllare all'indice: " + str(int(index)+2))
                    if row[self.work_accesso_programmabile_ZP] == "N": 
                        logging.error("trovato anomalia in check prestazione controllo all'indice: " + str(int(index)+2))
                        error_dict["error_controlli"].append(str(int(index)+2))
        
        out1 = ", \n".join(error_dict['error_controlli'])
        self.output_message = self.output_message + "\nerror_controlli: \n" + "at index: \n" + out1
        
        return error_dict

    '''Nel caso di prestazioni per esami strumentale, controllare se le priorità sono definite'''
    def ck_esami_strumentali(self, df_mapping, error_dict):
        print("start checking if prestazione esami is correct")
        error_dict.update({'error_esami_strumentali': []})

        for index, row in df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss] == "S":
                if row[self.work_priorita_U] == "N" and row[self.work_priorita_primo_accesso_D] == "N" and row[self.work_priorita_primo_accesso_P] == "N" and row[self.work_priorita_primo_accesso_B] == "N":
                    logging.info("Prestazione CONTROLLO da controllare all'indice: " + str(int(index)+2))
                    if row[self.work_accesso_programmabile_ZP] == "N": 
                        logging.error("trovato anomalia in check esami strumentali all'indice: " + str(int(index)+2))
                        error_dict["error_esami_strumentali"].append(str(int(index)+2))
        
        out1 = ", \n".join(error_dict['error_esami_strumentali'])
        self.output_message = self.output_message + "\nerror_esami_strumentali: \n" + "at index: \n" + out1
        
        return error_dict