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

class Check_canali():

    file_name = ""
    output_message = ""
    error_list = {}

    work_sheet = "" #sheet di lavoro di df_mapping
    work_codice_prestazione_siss = ""
    work_descrizione_prestazione_siss = ""
    work_codice_agenda_siss = ""
    work_casi_1_n = ""
    work_abilitazione_esposizione_siss = ""
    work_prenotabile_siss = ""
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
    work_inviante = ""
    work_accesso_farmacia = ""
    work_accesso_CCR = ""
    work_accesso_cittadino = ""
    work_accesso_MMG = ""
    work_accesso_amministrativo = ""
    work_accesso_PAI = ""
    work_gg_preparazione = ""
    work_gg_refertazione = ""

    work_index_codice_QD = 0
    work_index_codice_SISS_agenda = 0
    work_index_abilitazione_esposizione_SISS = 0
    work_index_codice_prestazione_SISS = 0
    work_index_operatore_logico_distretto = 0
    work_index_codici_disciplina_catalogo = 0
    
    def ck_canali_vuoti(self, df_mapping, error_dict):
        print("start checking canali di prenotazione configurati")
        error_dict.update({'error_canali_vuoti': []})

        xfile = openpyxl.load_workbook(self.file_data) #recupero file excel da file system
        sheet = xfile.get_sheet_by_name(self.work_sheet) #recupero sheet excel
        
        type_error = {
            1: "farmacia",
            2: "ccr",
            3: "cittadino",
            4: "mmg",
            5: "amministrativo",
            6: "pai"
        }
        list_error = []

        for index, row in df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss] == "S" and row[self.work_prenotabile_siss] == "S": #se prestazione esposta e prenotabile
                if row[self.work_accesso_farmacia] == "":
                    list_error.append(type_error[1])
                if row[self.work_accesso_CCR] == "":
                    list_error.append(type_error[2])
                if row[self.work_accesso_cittadino] == "":
                    list_error.append(type_error[3])
                if row[self.work_accesso_MMG] == "":
                    list_error.append(type_error[4])
                if row[self.work_accesso_amministrativo] == "":
                    list_error.append(type_error[5])
                if row[self.work_accesso_PAI] == "":
                    list_error.append(type_error[6])

                if list_error != []:
                    error_dict["error_canali_vuoti"].append(str(int(index)+2))
                    ind = str(int(index)+2)
                    out_message = ""
                    out_message = "__> I seguenti canali di accesso non sono valorizzati: '{}' ".format(", ".join(list_error))
                    if sheet[self.work_alert_column+ind].value is not None:
                        sheet[self.work_alert_column+ind] = str(sheet[self.work_alert_column+ind].value) + "; \n" + out_message
                    else:
                        sheet[self.work_alert_column+ind] = out_message

            list_error = []

        xfile.save(self.file_data)  
        return error_dict

    def ck_canali_PAI(self, df_mapping, error_dict):
        print("start checking canale PAI")
        error_dict.update({'error_canale_PAI': []})

        xfile = openpyxl.load_workbook(self.file_data) #recupero file excel da file system
        sheet = xfile.get_sheet_by_name(self.work_sheet) #recupero sheet excel

        for index, row in df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss] == "S" and row[self.work_prenotabile_siss] == "S": #se prestazione esposta e prenotabile
                if row[self.work_accesso_PAI] == "S":
                    if row[self.work_accesso_farmacia] == "S" or row[self.work_accesso_CCR] == "S" or row[self.work_accesso_cittadino] == "S" or row[self.work_accesso_MMG] == "S" or row[self.work_accesso_amministrativo] == "S":
                        error_dict["error_canale_PAI"].append(str(int(index)+2))
            
        out_message = ""
        for ind in error_dict['error_canale_PAI']:
            out_message = "__> Rilevati canali di accesso abilitati contemporaneamente al canale CREG PAI"
            out_message = out_message + "\n _> lasciare abilitato solo il canale CREG PAI o disabilitarlo."
            if sheet[self.work_alert_column+ind].value is not None:
                sheet[self.work_alert_column+ind] = str(sheet[self.work_alert_column+ind].value) + "; \n" + out_message #modificare colonna alert
            else:
                sheet[self.work_alert_column+ind] = out_message

        xfile.save(self.file_data)  
        return error_dict

    def ck_canali_abilitati(self, df_mapping, error_dict):
        print("start checking canali abilitati")
        error_dict.update({'error_canali_abilitati': []})

        xfile = openpyxl.load_workbook(self.file_data) #recupero file excel da file system
        sheet = xfile.get_sheet_by_name(self.work_sheet) #recupero sheet excel

        for index, row in df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss] == "S" and row[self.work_prenotabile_siss] == "S": #se prestazione esposta e prenotabile
                if row[self.work_accesso_PAI] == "S":
                    if row[self.work_accesso_farmacia] == "N" and row[self.work_accesso_CCR] == "N" and row[self.work_accesso_cittadino] == "N" and row[self.work_accesso_MMG] == "N" and row[self.work_accesso_amministrativo] == "N" and row[self.work_accesso_PAI] == "N":
                        error_dict["error_canali_abilitati"].append(str(int(index)+2))
            
        out_message = ""
        for ind in error_dict['error_canali_abilitati']:
            out_message = "__> Rilevati canali di accesso tutti disabilitati per prestazione esposta e prenotabile"
            out_message = out_message + "\n _> abilitare almeno uno dei canali di prenotazione proposti."
            if sheet[self.work_alert_column+ind].value is not None:
                sheet[self.work_alert_column+ind] = str(sheet[self.work_alert_column+ind].value) + "; \n" + out_message #modificare colonna alert
            else:
                sheet[self.work_alert_column+ind] = out_message

        xfile.save(self.file_data)  
        return error_dict