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

class Check_univocita_prestazione():

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

    '''def __init__(self):
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
        self.work_index_codici_disciplina_catalogo = data[1]["work_index"]["work_index_codici_disciplina_catalogo"]'''

    
    def ck_casi_1n(self, df_mapping, error_dict):
        print("start checking if casi 1:n is correct")
        error_dict.update({'error_casi_1n': []})

        xfile = openpyxl.load_workbook(self.file_data) #recupero file excel da file system
        sheet = xfile.get_sheet_by_name(self.work_sheet) #recupero sheet excel

        casi_1n_dict_error = {}

        agende_list = [] #Codice SISS Agenda
        prestazioni_list = [] #Codice Prestazione SISS
        agenda_prestazione_list = []
        metodica_distretti_dict = {} #dict delle metodiche e distretti delle prestazioni messe in lista
        for index, row in df_mapping.iterrows():
            #if row["Abilititazione Esposizione SISS"] == "S":
            if row[self.work_casi_1_n] != "OK" or row[self.work_casi_1_n] in "1:N":
                a_p = str(row[self.work_codice_agenda_siss]) + "_" + str(row[self.work_codice_prestazione_siss])
                m_d = row[self.work_codice_metodica] + "_" + row[self.work_codice_distretto]
                if a_p not in metodica_distretti_dict.keys() and row[self.work_abilitazione_esposizione_siss] == "S": 
                    #if m_d not in metodica_distretti_dict[a_p]:
                        #print("CASO 1:N corretto momentaneamente, all'indice:" + str(int(index)+2)) 
                        #print("A_P1: " + a_p + ", Abilititazione Esposizione SISS: " + row["Abilititazione Esposizione SISS"])
                        #agenda_prestazione_list.append(a_p)
                        #metodica_distretti_list.append(m_d)
                    metodica_distretti_dict = self.update_list_in_dict(metodica_distretti_dict, a_p, m_d)
                elif a_p in metodica_distretti_dict.keys() and m_d in metodica_distretti_dict[a_p] and row[self.work_abilitazione_esposizione_siss] == "S":
                    metodica_distretti_dict = self.update_list_in_dict(metodica_distretti_dict, a_p, m_d)
                    for md in metodica_distretti_dict[a_p]:
                        if md.split("_")[1] == "":
                            error_dict["error_casi_1n"].append(str(int(index)+1))
                            casi_1n_dict_error = self.update_list_in_dict(casi_1n_dict_error, str(int(index)+2), a_p + ": " + "caso 1:n con distretto vuoto")
                    error_dict["error_casi_1n"].append(str(int(index)+2))
                    casi_1n_dict_error = self.update_list_in_dict(casi_1n_dict_error, str(int(index)+2), a_p + ": " + ", ".join(metodica_distretti_dict[a_p]))
                    #print("trovato caso 1:n per la coppia agenda-prestazione, all'indice: " + str(int(index)+2))
                elif m_d.split("_")[1] == "" and a_p in metodica_distretti_dict.keys() and row[self.work_abilitazione_esposizione_siss] == "S":
                    metodica_distretti_dict = self.update_list_in_dict(metodica_distretti_dict, a_p, m_d)
                    error_dict["error_casi_1n"].append(str(int(index)+2))
                    casi_1n_dict_error = self.update_list_in_dict(casi_1n_dict_error, str(int(index)+2), a_p + ": " + "caso 1:n con distretto vuoto")
                elif a_p in metodica_distretti_dict.keys() and m_d not in metodica_distretti_dict[a_p] and row[self.work_abilitazione_esposizione_siss] == "S":
                    metodica_distretti_dict = self.update_list_in_dict(metodica_distretti_dict, a_p, m_d)
                else:
                    logging.info("trovato caso 1:n con abilitazione SISS a N corretta, all'indice: " + str(int(index)+2))
                    #print("trovato caso 1:n con abilitazione SISS a N corretta, all'indice: " + str(int(index)+2))
                    #print("A_P2: " + a_p + ", Abilititazione Esposizione SISS: " + row["Abilititazione Esposizione SISS"])
        
        out1 = ""
        out_message = ""
        for ind in error_dict['error_casi_1n']:
            out_message = "Casi 1:N: rilevato per la coppia prestazione/agenda: '{}'".format(", ".join(casi_1n_dict_error[ind]))
            out1 = out1 + "at index: " + ind + ", on agenda_prestazione: " + ", ".join(casi_1n_dict_error[ind]) + ", \n"
            if sheet["BY"+ind].value is not None:
                sheet["BY"+ind] = str(sheet["BY"+ind].value) + "; \n" + out_message #modificare colonna alert
            else:
                sheet["BY"+ind] = out_message

        self.output_message = self.output_message + "\nerror_casi_1n: \n" + "at index: \n" + out1
            
        xfile.save(self.file_data)  
        return error_dict