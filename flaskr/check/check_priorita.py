import pandas as pd
import numpy as np
import requests
import xlrd
from xlsxwriter.utility import xl_rowcol_to_cell
import yaml
import logging
import re
import openpyxl 

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
    work_index_codice_SISS_agenda = 0
    work_index_abilitazione_esposizione_SISS = 0
    work_index_codice_prestazione_SISS = 0
    work_index_operatore_logico_distretto = 0
    work_index_codici_disciplina_catalogo = 0

    
    def ck_prime_visite(self, df_mapping, error_dict): 
        #Descrizione Prestazione SISS
        print("start checking if prestazione prime visite is correct")
        error_dict.update({'error_prime_visite': []})
        
        xfile = openpyxl.load_workbook(self.file_data) #recupero file excel da file system
        sheet = xfile.get_sheet_by_name(self.work_sheet) #recupero sheet excel

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
        
        out_message = ""
        for ind in error_dict['error_prime_visite']:
            out_message = "__> Rilevato errore di priorità per prestazione PRIMA VISITA"
            if sheet[self.work_alert_column+ind].value is not None:
                sheet[self.work_alert_column+ind] = str(sheet[self.work_alert_column+ind].value) + "; \n" + out_message 
            else:
                sheet[self.work_alert_column+ind] = out_message

        xfile.save(self.file_data)  
        return error_dict

    '''Nel caso di prestazione visita di controllo, verificare se è presente il campo Accesso programmabile ZP'''
    def ck_controlli(self, df_mapping, error_dict):
        print("start checking if prestazione controlli is correct")
        error_dict.update({'error_controlli': []})

        xfile = openpyxl.load_workbook(self.file_data) #recupero file excel da file system
        sheet = xfile.get_sheet_by_name(self.work_sheet) #recupero sheet excel

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
        
        out_message = ""
        for ind in error_dict['error_controlli']:
            out_message = "__> Rilevato possibile errore di priorità per prestazione DI CONTROLLO"
            out_message = out_message + "\n _> controllare che l'accesso programmabile ZP non sia a N"
            if sheet[self.work_alert_column+ind].value is not None:
                sheet[self.work_alert_column+ind] = str(sheet[self.work_alert_column+ind].value) + "; \n" + out_message 
            else:
                sheet[self.work_alert_column+ind] = out_message

        xfile.save(self.file_data) 
        return error_dict

    '''Nel caso di prestazioni per esami strumentale, controllare se le priorità sono definite'''
    def ck_esami_strumentali(self, df_mapping, error_dict):
        print("start checking if prestazione esami is correct")
        error_dict.update({'error_esami_strumentali': []})

        xfile = openpyxl.load_workbook(self.file_data) #recupero file excel da file system
        sheet = xfile.get_sheet_by_name(self.work_sheet) #recupero sheet excel

        str_check = "VISITA"

        for index, row in df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss] == "S":
                if str_check not in row[self.work_descrizione_prestazione_siss]:  
                    if row[self.work_priorita_U] == "N" and row[self.work_priorita_primo_accesso_D] == "N" and row[self.work_priorita_primo_accesso_P] == "N" and row[self.work_priorita_primo_accesso_B] == "N":
                        logging.info("Prestazione ESAME da controllare all'indice: " + str(int(index)+2))
                        if row[self.work_accesso_programmabile_ZP] == "N": 
                            logging.error("trovato anomalia in check esami strumentali all'indice: " + str(int(index)+2))
                            error_dict["error_esami_strumentali"].append(str(int(index)+2))
        
        out1 = ", \n".join(error_dict['error_esami_strumentali'])
        self.output_message = self.output_message + "\nerror_esami_strumentali: \n" + "at index: \n" + out1
        
        out_message = ""
        for ind in error_dict['error_esami_strumentali']:
            out_message = "__> Rilevato errore di priorità per prestazione di tipo esami strumentali"
            if sheet[self.work_alert_column+ind].value is not None:
                sheet[self.work_alert_column+ind] = str(sheet[self.work_alert_column+ind].value) + "; \n" + out_message 
            else:
                sheet[self.work_alert_column+ind] = out_message

        xfile.save(self.file_data)
        return error_dict