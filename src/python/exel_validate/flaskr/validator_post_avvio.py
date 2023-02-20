import argparse
import itertools
import json
import logging
from operator import le
import random
import re
import time
from collections import OrderedDict
from functools import reduce
from pathlib import Path
from typing import Dict, List
from matplotlib.pyplot import flag


#import openpyxl 
import pandas as pd
import numpy as np
import requests
import xlrd
from xlsxwriter.utility import xl_rowcol_to_cell
import yaml
#import xlsxwriter
#from openpyxl.utils import get_column_letter

from flaskr.check.check_post_avvio import Check_post_avvio

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

RESULT_VALIDATION = "..\check_excel_result.txt"


class Check_action():

    file_name = ""
    file_data = {}
    file_rivisto = {}
    configurazione_rivisto = {}
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
    work_combinata = ""
    work_codice_agenda_interno = ""
    work_codice_prestazione_interno = ""
    work_inviante = ""
    work_accesso_farmacia = ""
    work_accesso_CCR = ""
    work_accesso_cittadino = ""
    work_accesso_MMG = ""
    work_accesso_amministrativo = ""
    work_accesso_PAI = ""
    work_gg_preparazione = ""
    work_gg_refertazione = ""
    work_nota_operatore = ""
    work_nota_preparazione = ""
    work_nota_agenda = ""
    work_nota_revoca = ""
    work_risorsa = ""

    work_index_codice_QD = 0
    work_index_codice_SISS_agenda = 0
    work_index_abilitazione_esposizione_SISS = 0
    work_index_codice_prestazione_SISS = 0
    work_index_operatore_logico_distretto = 0
    work_index_codici_disciplina_catalogo = 0
    work_index_operatore_logico_QD = 0

    work_alert_column = ""
    work_delimiter = ""


    def __init__(self, data, excel_file):
        #self.output_message = ""
        #with open("./flaskr/config_validator.yml", "rt", encoding='utf8') as yamlfile:
        #    data = yaml.load(yamlfile, Loader=yaml.FullLoader)
        #logger.debug(data)
        self.work_sheet = data[0]["work_column"]["work_sheet"] 
        self.work_codice_prestazione_siss = data[0]["work_column"]["work_codice_prestazione_siss"]
        self.work_descrizione_prestazione_siss = data[0]["work_column"]["work_descrizione_prestazione_siss"]
        self.work_codice_agenda_siss = data[0]["work_column"]["work_codice_agenda_siss"]
        self.work_casi_1_n = data[0]["work_column"]["work_casi_1_n"]
        self.work_abilitazione_esposizione_siss = data[0]["work_column"]["work_abilitazione_esposizione_siss"]
        self.work_prenotabile_siss = data[0]["work_column"]["work_prenotabile_siss"]
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
        self.work_combinata = data[0]["work_column"]["work_combinata"]
        self.work_codice_agenda_interno = data[0]["work_column"]["work_codice_agenda_interno"]
        self.work_codice_prestazione_interno = data[0]["work_column"]["work_codice_prestazione_interno"]
        self.work_inviante = data[0]["work_column"]["work_inviante"]
        self.work_accesso_farmacia = data[0]["work_column"]["work_accesso_farmacia"]
        self.work_accesso_CCR = data[0]["work_column"]["work_accesso_CCR"]
        self.work_accesso_cittadino = data[0]["work_column"]["work_accesso_cittadino"]
        self.work_accesso_MMG = data[0]["work_column"]["work_accesso_MMG"]
        self.work_accesso_amministrativo = data[0]["work_column"]["work_accesso_amministrativo"]
        self.work_accesso_PAI = data[0]["work_column"]["work_accesso_PAI"]
        self.work_gg_preparazione = data[0]["work_column"]["work_gg_preparazione"]
        self.work_gg_refertazione = data[0]["work_column"]["work_gg_refertazione"]
        self.work_nota_operatore = data[0]["work_column"]["work_nota_operatore"]
        self.work_nota_preparazione = data[0]["work_column"]["work_nota_preparazione"]
        self.work_nota_agenda = data[0]["work_column"]["work_nota_agenda"]
        self.work_nota_revoca = data[0]["work_column"]["work_nota_revoca"]

        self.work_risorsa = data[0]["work_column"]["work_risorsa"]

        self.work_alert_column = data[1]["work_index"]["work_alert_column"]
        self.work_map_value_column = data[1]["work_index"]["work_map_value_column"]
        try:
            self.work_delimiter = data[2]["work_separator"]["work_delimiter"]
        except:
            self.work_delimiter = "," #valore di default
        self.file_data = excel_file


    def initializer(self, file_path_rivisto, checked_dict):
        self.file_rivisto = file_path_rivisto #file rivisto
        self.configurazione_rivisto = checked_dict #nomi colonne del rivisto
        print(checked_dict)
        df_mapping = pd.read_excel(self.file_data, sheet_name=self.work_sheet, converters={self.work_codici_disciplina_catalogo: str, self.work_codice_prestazione_siss: str, self.work_codice_agenda_siss: str, self.work_codice_prestazione_interno: str, self.work_combinata: str}).replace(np.nan, '', regex=True)
        df_rivisto = pd.read_excel(self.file_rivisto, sheet_name=checked_dict["Sheet"], converters={"CD_AGENDA": str, "CD_PRESTAZIONE_SISS": str, "CD_INTERNO_PRESTAZIONE": str, "ID_COMBINATA": str}).replace(np.nan, '', regex=True)
        
        error = self.analizer(df_mapping, df_rivisto)

        error_dict = {
            "error_post_avvio": error
        }
        self._validation(error_dict)


    def analizer(self, df_mapping, df_rivisto):
        Post_avvio_error = Check_post_avvio.ck_post_avvio(self, df_mapping, df_rivisto, {})
        return Post_avvio_error


    def _validation(self, error_dict):
        print("questi sono gli errori indivuduati e separati per categoria:\n %s", error_dict)
        #self.output_message = self.output_message + "\n" + json.dumps(error_dict)
        '''df = pd.DataFrame(rows_list)
        with pd.ExcelWriter(self.file_name, engine='openpyxl', mode='a') as writer:
            df.to_excel(writer, sheet_name='new_mapping', index=False)'''
        print("\nPer osservare i risultati ottenuti, controllare il file prodotto: check_excel_result.txt")
        file = open(RESULT_VALIDATION, "w") 
        file.write(self.output_message)
        file.close() 
        

    def findCell(self, sh, searchedValue, start_col):
        result_coord = []
        result_value = []

        for row in range(sh.nrows):
            for col in range(start_col, start_col+1):
                myCell = sh.cell(row, col)
                myValue = sh.cell(row, self.work_index_codice_prestazione_SISS)
                abilita = sh.cell(row, self.work_index_abilitazione_esposizione_SISS)
                if myCell.value == searchedValue and abilita.value == "S":
                    result_coord.append(str(row) + "#" + str(col))
                    result_value.append(myValue.value)
                    #return row, col#xl_rowcol_to_cell(row, col)

        if result_coord == []:
            return -1
        return result_coord#, result_value

    def findCell_agenda(self, sh, searchedValue, start_col):
        result_coord = []
        result_value = []

        for row in range(sh.nrows):
            for col in range(start_col, start_col+1):
                myCell = sh.cell(row, col)
                myValue = sh.cell(row, self.work_index_codice_SISS_agenda) #Codice SISS agenda 15
                #abilita = sh.cell(row, self.work_index_abilitazione_esposizione_SISS-1) #abilitazione esposizione SISS 28
                if myCell.value == searchedValue: # and abilita.value == "S":
                    result_coord.append(str(row) + "|" + str(col))
                    result_value.append(myValue.value)
                    #return row, col#xl_rowcol_to_cell(row, col)

        if result_coord == []:
            return -1
        return result_coord#, result_value

    def findCell_dataframe_MAP_string(self, df, searchedValue, key_rivisto, column_name):
        result_coord = ""
        flag_find = False

        #print("start findcell dataframe")
        for index, row in df.iterrows():
            mapping_key = str(row[self.work_codice_agenda_siss]).strip()+"|"+str(row[self.work_codice_prestazione_siss]).strip()+"|"+str(row[self.work_codice_prestazione_interno]).strip()
            #print("iterate mapping: " + mapping_key)
            #print("trovata corrisponenza key: " + searchedValue)
            if mapping_key == key_rivisto and row[self.work_abilitazione_esposizione_siss] == "S":
                #print("trovata corrisponenza key: " + row[column_name] + " e " + searchedValue)
                m_row_value = str(row[column_name])
                r_row_value = str(searchedValue)
                #print("m_row_value", m_row_value)
                #print("r_row_value", r_row_value)
                flag_find = True
                
                if m_row_value != r_row_value: #check the two row
                    #result_coord.append(column_name + ": " + str(row[column_name]))
                    result_coord = column_name + ": " + str(row[column_name])
                    print("errore corripondenza valori")
        
        if flag_find == False: #coppia non esistente in rivisto
            return -2
        elif result_coord == "": #non c'è stata corrispondenza 
            return -1
        return result_coord #nessun errore trovato, c'è stata corrispondenza


    def findCell_dataframe_RIV_string(self, df, searchedValue, key_mapping, column_name):
        result_coord = ""
        flag_find = False

        #print("start findcell dataframe")
        for index, row in df.iterrows():
            rivisto_key = str(row[self.configurazione_rivisto["Agenda"]]).strip()+"|"+str(row[self.configurazione_rivisto["PrestazioneSISS"]]).strip()+"|"+str(row[self.configurazione_rivisto["PrestazioneInterna"]]).strip()
            #print("iterate mapping: " + rivisto_key)
            #print("trovata corrisponenza key: " + searchedValue)
            if rivisto_key == key_mapping:
                #print("trovata corrisponenza key: " + row[column_name] + " e " + searchedValue)
                r_row_value = str(row[column_name])
                m_row_value = str(searchedValue)
                #print("m_row_value", m_row_value)
                #print("r_row_value", r_row_value)
                flag_find = True
                
                if m_row_value != r_row_value: #check the two row
                    #result_coord.append(column_name + ": " + str(row[column_name])) #modifico index con row[column_name]
                    result_coord = column_name + ": " + str(row[column_name])
                    print("errore corripondenza valori")
                
        if flag_find == False: #coppia non esistente in rivisto
            return -2
        elif result_coord == "": #non c'è stata corrispondenza 
            return -1
        return result_coord #nessun errore trovato, c'è stata corrispondenza


    def findCell_dataframe_MAP(self, df, searchedValue, key_rivisto, column_name):
        result_coord = ""
        flag_find = False

        #print("start findcell dataframe")
        for index, row in df.iterrows():
            mapping_key = str(row[self.work_codice_agenda_siss]).strip()+"|"+str(row[self.work_codice_prestazione_siss]).strip()+"|"+str(row[self.work_codice_prestazione_interno]).strip()
            #print("iterate mapping: " + mapping_key)
            #print("trovata corrisponenza key: " + searchedValue)
            if mapping_key == key_rivisto and row[self.work_abilitazione_esposizione_siss] == "S":
                #print("trovata corrisponenza key: " + row[column_name] + " e " + searchedValue)
                m_row_value = str(row[column_name]).split(self.work_delimiter)
                r_row_value = str(searchedValue).split(self.work_delimiter)
                #print("m_row_value", m_row_value)
                #print("r_row_value", r_row_value)
                #m_row_value = m_row_value.sort(key = str)
                #r_row_value = r_row_value.sort(key = str)
                flag_find = True
                
                result1 =  all(elem in r_row_value  for elem in m_row_value) #check list1 in list2
                result2 =  all(elem in m_row_value  for elem in r_row_value) #check list2 in list1
                if not result1 and not result2 and len(m_row_value)!=len(r_row_value): #check result and lenght
                    #result_coord.append(column_name + ": " + str(row[column_name]))
                    result_coord = column_name + ": " + str(row[column_name])
                    print("errore corripondenza valori")
                    #print("m_row_value", m_row_value)
                    #print("r_row_value", r_row_value)
                
        if flag_find == False: #coppia non esistente in rivisto
            return -2
        elif result_coord == "": #non c'è stata corrispondenza 
            return -1
        return result_coord #nessun errore trovato, c'è stata corrispondenza


    def findCell_dataframe_RIV(self, df, searchedValue, key_mapping, column_name):
        result_coord = ""
        flag_find = False

        #print("start findcell dataframe")
        for index, row in df.iterrows():
            rivisto_key = str(row[self.configurazione_rivisto["Agenda"]]).strip()+"|"+str(row[self.configurazione_rivisto["PrestazioneSISS"]]).strip()+"|"+str(row[self.configurazione_rivisto["PrestazioneInterna"]]).strip()
            #print("iterate mapping: " + rivisto_key)
            #print("trovata corrisponenza key: " + searchedValue)
            if rivisto_key == key_mapping:
                #print("trovata corrisponenza key: " + row[column_name] + " e " + searchedValue)
                r_row_value = str(row[column_name]).split(self.work_delimiter) #list
                m_row_value = str(searchedValue).split(self.work_delimiter) #list
                #print("m_row_value", m_row_value)
                #print("r_row_value", r_row_value)
                flag_find = True
                
                result1 =  all(elem in m_row_value  for elem in r_row_value) #check list1 in list2
                result2 =  all(elem in r_row_value  for elem in m_row_value) #check list2 in list1
                if not result1 and not result2 and len(m_row_value)!=len(r_row_value): #check result and lenght
                    #result_coord.append(column_name + ": " + str(row[column_name]))
                    result_coord = column_name + ": " + str(row[column_name])
                    print("errore corripondenza valori")
                    #m_row_value = m_row_value.sort(key = str)
                    #r_row_value = r_row_value.sort(key = str)
                    #rint("m_row_value", m_row_value)
                    #print("r_row_value", r_row_value)
                
        if flag_find == False: #coppia non esistente in rivisto
            return -2
        elif result_coord == "": #non c'è stata corrispondenza 
            return -1
        return result_coord #nessun errore trovato, c'è stata corrispondenza

    '''Metodo che aggiunge elemento in una lista esistente o crea la lista nel caso
    non fosse presente'''
    def update_list_in_dict(self, dictio, index, element):
        if index in dictio.keys():
            dictio[index].append(element)
        else:
            dictio[index] = [element]
        return dictio


    def list_duplicates(self, seq):
        seen = set()
        seen_add = seen.add
        # adds all elements it doesn't know yet to seen and all other to seen_twice
        seen_twice = set( x for x in seq if x in seen or seen_add(x) )
        # turn the set into a list (as requested)
        return list( seen_twice )


    