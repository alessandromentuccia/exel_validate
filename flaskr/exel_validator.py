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


class PhraseTemplate(object):
    def __init__(self, original_line: str, string_to_fill: str, slot_list: List[List[str]]):
        self.original_line = original_line
        self.string_to_fill = string_to_fill
        self.slot_list = slot_list


class Check_action():

    file_name = ""
    file_data = {}
    catalogo = OrderedDict()
    flag_check_list = []
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
        with open("config_validator.yml", "r") as yamlfile:
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

        self.work_index_codice_QD = data[0]["work_index"]["work_index_codice_QD"]
        self.work_index_op_logic_distretto = data[0]["work_index"]["work_index_op_logic_distretto"]
        self.work_index_codice_SISS_agenda = data[0]["work_index"]["work_index_codice_SISS_agenda"]
        self.work_index_abilitazione_esposizione_SISS = data[0]["work_index"]["work_index_abilitazione_esposizione_SISS"]
        self.work_index_codice_prestazione_SISS = data[0]["work_index"]["work_index_codice_prestazione_SISS"]
        self.work_index_operatore_logico_distretto = data[0]["work_index"]["work_index_operatore_logico_distretto"]
        self.work_index_codici_disciplina_catalogo = data[0]["work_index"]["work_index_codici_disciplina_catalogo"]

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
        df_mapping = pd.read_excel(template_file, 0, converters={self.work_codici_disciplina_catalogo: str}).replace(np.nan, '', regex=True)
        #print ("print JSON")
        #print(sh)
        
        catalogo_dir = "c:\\Users\\aless\\exel_validate\\CCR-BO-CATGP#01_Codifiche_attributi_catalogo GP++_201910.xls"

        sheet_QD = pd.read_excel(catalogo_dir, sheet_name='QD' )
        sheet_Metodiche = pd.read_excel(catalogo_dir, sheet_name='METODICHE' )
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

        error_dict = {
            "QD_error": QD_error,
            "metodiche_error": metodiche_error,
            "distretti_error": distretti_error,
            "priorita_error": priorita_error,    
            "univocita_prestazione_error": univocita_prestazione_error
        }

        self._validation(error_dict)

    def check_column_name(self, df_mapping):
        print("check the used column name of the excel file")

    def check_qd(self, df_mapping, sheet_QD):
        print("start checking QD") #Codice Quesito Diagnostico
        #controllo se per ogni Agenda sono inseriti gli stessi QD
        error_dict = {}
        
        error_QD_sintassi = self.ck_QD_sintassi(df_mapping, error_dict)
        error_QD_agenda = self.ck_QD_agenda(df_mapping, error_QD_sintassi)
        error_QD_disciplina_agenda = self.ck_QD_disciplina_agenda(df_mapping, sheet_QD, error_QD_agenda)
        error_QD_descrizione = self.ck_QD_descrizione(df_mapping, sheet_QD, error_QD_disciplina_agenda)
        error_QD_operatori_logici = self.ck_QD_operatori_logici(df_mapping, error_QD_descrizione)

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

        error_metodica_sintassi = self.ck_metodica_sintassi(df_mapping, error_dict)
        error_metodica_inprestazione = self.ck_metodica_inprestazione(df_mapping, sheet_Metodiche, error_metodica_sintassi)
        error_metodica_descrizione = self.ck_metodica_descrizione(df_mapping, sheet_Metodiche, error_metodica_inprestazione)

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

        error_distretti_sintassi = self.ck_distretti_sintassi(df_mapping, error_dict)
        error_distretti_inprestazione = self.ck_distretti_inprestazione(df_mapping, sheet_Distretti, error_distretti_sintassi)
        error_distretti_descrizione = self.ck_distretti_descrizione(df_mapping, sheet_Distretti, error_distretti_inprestazione)
        error_distretti_operatori_logici = self.ck_distretti_operatori_logici(df_mapping, error_distretti_descrizione)

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

        error_prime_visite = self.ck_prime_visite(df_mapping, error_dict)
        error_controlli = self.ck_controlli(df_mapping, error_prime_visite)
        error_esami_strumentali =self.ck_esami_strumentali(df_mapping, error_controlli)

        error_dict = error_esami_strumentali
        '''error_list = {
            "error_prime_visite": error_prime_visite,
            "error_controlli": error_controlli,
            "error_esami_strumentali": error_esami_strumentali
        }'''

        return error_dict

    def check_univocita_prestazione(self, df_mapping):
        print("start checking univocità delle prestazioni")
        error_dict = self.ck_casi_1n(df_mapping, {})

        return error_dict

    def ck_QD_agenda(self, df_mapping, error_dict):
        print("start checking if foreach agenda there are the same QD")
        
        error_dict.update({'error_QD_agenda': []})
        
        agenda = df_mapping[self.work_codice_agenda_siss].iloc[2]
        last_QD = df_mapping[self.work_codice_QD].iloc[2]
        for index, row in df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss] == "S":
                if row[self.work_codice_agenda_siss] == agenda:
                    #print("- same agenda -")
                    if row[self.work_codice_QD] != last_QD:
                        print("error QD at index:" +  str(int(index)+2))
                        #error_list.append(str(int(index)+2))
                        error_dict['error_QD_agenda'].append(str(int(index)+2))
                    #else: 
                    #    print("correct QD")
                else: 
                    #print("- the agenda is changed -")
                    agenda = row[self.work_codice_agenda_siss]
                    last_QD = row[self.work_codice_QD]

        print("error_dict: %s", error_dict)
        return error_dict

    def ck_QD_disciplina_agenda(self, df_mapping, sheet_QD, error_dict):
        print("start checking if foreach agenda there is the same Disciplina for all the QD")
        #tutti i QD di un agenda hanno la stessa disciplina
        error_dict.update({
            'error_QD_disciplina_agenda': [],
            'error_disciplina_mancante' : []
        }) 

        wb = xlrd.open_workbook(self.file_name)
        sheet_mapping = wb.sheet_by_index(0)
        print("sheet caricato")
        
        #disciplina_QD_column = sheet_QD[['Cod Disciplina','Codice Quesito']]
        #print("disciplina_QD_column: %s", disciplina_QD_column)
        agende_viewed = []
        for index, row in df_mapping.iterrows():
            disci_flag = False
            if row[self.work_codice_agenda_siss] not in agende_viewed:
                searchedAgenda = row[self.work_codice_agenda_siss]
                result = self.findCell_agenda(sheet_mapping, searchedAgenda, self.work_index_codice_prestazione_SISS) #prendo tutte le righe con questa agenda

                if result != -1:
                    result_disciplina_last = ""
                    agende_error_list = []
                    for res in result: #per ogni risultato controllo che ci sia la stessa disciplina
                        r = res.split("#")[0] #row agenda
                        c = res.split("#")[1] #column agenda
                        result_disciplina = sheet_mapping.cell(int(r), self.work_index_codici_disciplina_catalogo).value #disciplina da catalogo
                        if result_disciplina != "":
                            if result_disciplina_last != "": #se non è la prima iterazione
                                if result_disciplina != result_disciplina_last:
                                    disci_flag = True
                                    agende_error_list.append(str(int(r)+1))
                                    print("result_disciplina: " + result_disciplina + ", result_disciplina_last: " + result_disciplina_last)
                                else: 
                                    result_disciplina_last = result_disciplina    
                else:
                    error_dict['error_disciplina_mancante'].append(str(int(index)+2))     #inserisco la riga senza disciplina negli errori

                    if disci_flag == True: #se durante il mapping con la sua disciplina, questa non viene rilevata, allora è errore
                        for age in agende_error_list:
                            error_dict['error_QD_disciplina_agenda'].append(age)

        '''for index, row in df_mapping.iterrows():
            QD_list = row["Codice Quesito Diagnostico"].split(",")
            disci = ""
            if row["Abilititazione Esposizione SISS"] == "S": 
                disci_flag = 0
                if QD_list is not None:

                    for QD in QD_list:
                        if QD != "":
                            short_sheet = sheet_QD.loc[sheet_QD["Codice Quesito"] == QD.strip()]  #[str(i) for i in                      
                            try:
                                disciplina_mapping_row = row[self.work_codici_disciplina_catalogo]
                            except:    
                                disciplina_mapping_row = ""
                                disci_flag = -1

                            if disciplina_mapping_row != "":
                                print("disciplina in catalogo disciplina 11:" + disciplina_mapping_row + " + " + disci)
                                
                                for d in short_sheet["Cod Disciplina"]: 
                                    if disciplina_mapping_row == d: 
                                        disci_flag = 1
                                        print("disciplina in catalogo disciplina 22:" + str(disciplina_mapping_row) + " + " + disci)
                                        if str(disciplina_mapping_row) == disci or disci == "": # controllo se disciplina è uguale a quella precedente
                                            print("correct Disciplina")
                                        else: # se non è uguale è errore
                                            print("error QD on index:" + str(int(index)+2))
                                            error_dict['error_QD_disciplina_agenda'].append(str(int(index)+2))
                                    else: 
                                        print("disciplina diversa da quella di QD: " + str(d))
                            else:
                                disci_flag = -1
                        else:
                            disci_flag = 1

                if disci_flag == 0: #se durante il mapping con la sua disciplina, questa non viene rilevata, allora è errore
                    error_dict['error_QD_disciplina_agenda'].append(str(int(index)+2))
                elif disci_flag == -1:
                    error_dict['error_disciplina_mancante'].append(str(int(index)+2))
                disci = str(disciplina_mapping_row)'''

        return error_dict
            
    def ck_QD_sintassi(self, df_mapping, error_dict):
        print("start checking if there is ',' separator between each QD defined")
        error_dict.update({
            'error_QD_caratteri_non_consentiti': [],
            'error_QD_spazio_bordi': [],
            'error_QD_spazio_internamente': [],
        })
        string_check = re.compile('1234567890,Q') #lista caratteri ammessi

        for index, row in df_mapping.iterrows():
            #print("QD: " + row["Codice Quesito Diagnostico"])
            if row[self.work_abilitazione_esposizione_siss] == "S": 
                #flag_error = False
                if row[self.work_codice_QD] is not None:
                    row_replace = row[self.work_codice_QD].replace(" ", "")
                    if " " in row[self.work_codice_QD]:
                        if " " in row_replace.strip():
                            #print("string contain space inside the string")
                            error_dict['error_QD_spazio_internamente'].append(str(int(index)+2))
                        else:
                            #print("string contain space in the border")
                            error_dict['error_QD_spazio_bordi'].append(str(int(index)+2))
                    elif(string_check.search(row_replace) != None):
                        #print("String contains other Characters.")
                        error_dict['error_QD_caratteri_non_consentiti'].append(str(int(index)+2))
                        #flag_error = True
                    #elif " " in r:
                    #    print("string contain space")
                    #    error_dict['error_QD_trovato_spazio'].append(str(int(index)+2))
                    #else: 
                    #    print("String does not contain other characters") 

            #if flag_error == True:
            #    error_dict['error_QD_caratteri_non_consentiti'].append(str(int(index)+2))

        return error_dict

    def ck_QD_descrizione(self, df_mapping, sheet_QD, error_dict):
        print("start checking if there are the relative QD description")
        error_dict.update({
            'error_QD_descrizione': [],
            'error_QD_descrizione_space_bordo': [],
            'error_QD_descrizione_space_interno': []
        })
        
        for index, row in df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss] == "S":
                QD_string = row[self.work_codice_QD].split(",")
                description_list = row[self.work_descrizione_QD]#.split(",")
                flag_error = False
                if len(QD_string) != len(row[self.work_descrizione_QD].split(",")):
                    print("il numero di descrizioni è diverso dal numero di QD all'indice " + str(index))
                    flag_error = True

                if QD_string is not None:
                    for QD in QD_string:
                        if QD != "":
                            QD = QD.strip()
                            QD_catalogo = sheet_QD.loc[sheet_QD["Codice Quesito"] == QD]  
                            #print("QD: " + str(QD)) 
                            #try:
                            
                            if description_list != description_list.strip(): #there is a space in the beginning or in the end
                                error_dict['error_QD_descrizione_space_bordo'].append(str(int(index)+2))
                                logging.error("ERROR SPACE BORDI: controllare QD: " + QD + " all'indice: " + str(int(index)+2))
                                description_list = description_list.strip()

                            if " ," in description_list or ", " in description_list:
                                #print("print QD_catalogo2:" + QD_catalogo)
                                #print("controllare manualmente qual'è il problema")
                                print("QD: " + QD + ", Quesiti Diagnostici size:" + str(QD_catalogo.size) + ", description_list: %s", description_list)
                                logging.error("ERROR SPACE INTERNO: controllare QD: " + QD + " all'indice: " + str(int(index)+2))
                                error_dict['error_QD_descrizione_space_interno'].append(str(int(index)+2))
                                description_list = description_list.replace(", ", ",")
                                description_list = description_list.replace(" ,", ",")
                            try:
                                if QD_catalogo["Quesiti Diagnostici"].values[0] not in description_list.split(","):
                                    print("la descrizione QD non è presente all'indice " + str(int(index)+2))
                                    #print("QD: " + QD + ", Quesiti Diagnostici: " + QD_catalogo["Quesiti Diagnostici"].values[0] + ", Description_list: %s", description_list)
                                    logging.error("ERROR DESCRIZIONE: controllare QD: " + QD + " all'indice: " + str(int(index)+2))
                                    flag_error = True
                            except: #togliere try/catch e gestire gli spazi nell'if sopra
                                #print("print QD_catalogo2:" + QD_catalogo)
                                print("controllare manualmente qual'è il problema all'indice: " + str(int(index)+2))
                                #print("QD: " + QD + ", Quesiti Diagnostici size:" + str(QD_catalogo.size) + ", Description_list: %s", description_list)
                                #logging.error("controllare manualmente il QD: " + QD + " all'indice: " + str(int(index)+2))
                                #error_dict['error_QD_descrizione_space_interno'].append(str(int(index)+2))  
                                      
                            
                if flag_error == True:
                   error_dict['error_QD_descrizione'].append(str(int(index)+2))  

        return error_dict

    def ck_QD_operatori_logici(self, df_mapping, error_dict):
        print("start checking if there are the same logic op. for each agenda")
        error_dict.update({'error_QD_operatori_logici': []})
        
        agenda = df_mapping[self.work_codice_agenda_siss].iloc[2]
        last_OP = df_mapping[self.work_operatore_logico_QD].iloc[2]
        for index, row in df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss] == "S":
                if row[self.work_codice_agenda_siss] == agenda:
                    if row[self.work_operatore_logico_QD] != last_OP:
                        print("error OP at index:" +  str(int(index)+2))
                        error_dict['error_QD_operatori_logici'].append(str(int(index)+2))
                else: 
                    agenda = row[self.work_codice_agenda_siss]
                    last_OP = row[self.work_operatore_logico_QD]

        #print("error_dict: %s", error_dict)
        return error_dict


    def ck_metodica_inprestazione(self, df_mapping, sheet_Metodiche, error_dict):
        print("start checking if metodica are correct")
        error_dict.update({'error_metodica_inprestazione': []})
        
        for index, row in df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss] == "S":
                Metodica_string = row[self.work_codice_metodica].split(",")

                siss = "" #
                siss_flag = False
                if Metodica_string is not None:
                    for metodica in Metodica_string:
                        if metodica != "":
                            short_sheet = sheet_Metodiche.loc[sheet_Metodiche["Codice Metodica"] == metodica]                      
                            
                            #print("disciplina " + str(disciplina["Cod Disciplina"]) + " " + str(disciplina["Codice Quesito"]))
                            print("prestazione della metodica in mapping:" + row[self.work_codice_prestazione_siss] + " + " + siss)
                            for cod_SISS_cat in short_sheet["Codice SISS"]: 
                                if row[self.work_codice_prestazione_siss] == cod_SISS_cat: 
                                    siss_flag = True
                                    print("prestazione della metodica in mapping 22:" + str(row[self.work_codice_prestazione_siss]) + " + " + siss)
                                    if str(row[self.work_codice_prestazione_siss]) == siss or siss == "": # controllo se disciplina è uguale a quella precedente
                                        print("correct Metodica")
                                    else: # se non è uguale è errore
                                        print("error metodica on index:" + str(int(index)+2))
                                        error_dict['error_metodica_inprestazione'].append(str(int(index)+2))
                                else:
                                    print("disciplina diversa da quella di metodica: " + str(cod_SISS_cat))
                        else:
                            siss_flag = True

                if siss_flag == False: #se durante il mapping con la sua disciplina, questa non viene rilevata, allora è errore
                    error_dict['error_metodica_inprestazione'].append(str(int(index)+2))
                siss = str(row[self.work_codice_prestazione_siss])

        return error_dict

    def ck_metodica_sintassi(self, df_mapping, error_dict):
        print("start checking if there is ',' separator between each metodiche defined")
        error_dict.update({
            'error_metodica_caratteri_non_consentiti': [],
            'error_metodica_spazio_bordi': [],
            'error_metodica_spazio_internamente': []
        })
        string_check = re.compile('1234567890,M')

        for index, row in df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss] == "S":
                print("Metodica: " + row[self.work_codice_metodica])
                #flag_error = False
                if row[self.work_codice_metodica] is not None:
                    '''r = row["Codice Metodica"].strip()
                    if(string_check.search(row["Codice Metodica"]) != None):
                        print("String contains other Characters.")
                        #flag_error = True
                        error_dict['error_metodica_caratteri_non_consentiti'].append(str(int(index)+2))
                    elif " " in r:
                        print("string contain space")
                        error_dict['error_metodica_trovato_spazio'].append(str(int(index)+2))
                    '''
                    row_replace = row[self.work_codice_metodica].replace(" ", "")
                    if " " in row[self.work_codice_metodica]:
                        if " " in row_replace.strip():
                            print("string contain space inside the string")
                            error_dict['error_metodica_spazio_internamente'].append(str(int(index)+2))
                        else:
                            print("string contain space in the border")
                            error_dict['error_metodica_spazio_bordi'].append(str(int(index)+2))
                    elif(string_check.search(row_replace) != None):
                        print("String contains other Characters.")
                        error_dict['error_metodica_caratteri_non_consentiti'].append(str(int(index)+2))

            #if flag_error == True:
            #    error_dict['error_metodica_separatore'].append(str(int(index)+2))
        return error_dict

    def ck_metodica_descrizione(self, df_mapping, sheet_Metodiche, error_dict):
        print("start checking if metodica have the correct description")
        error_dict.update({'error_metodica_descrizione': []})
        
        for index, row in df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss] == "S":
                Metodica_string = row[self.work_codice_metodica].split(",")
                Description_list = row[self.work_descrizione_metodica].split(",")
                flag_error = False
                if len(Metodica_string) != len(Description_list):
                    print("il numero di descrizioni è diverso dal numero di metodiche all'indice " + str(index))
                    flag_error = True

                if Metodica_string is not None:
                    for metodica in Metodica_string:
                        if metodica != "":
                            metodica_catalogo = sheet_Metodiche.loc[sheet_Metodiche["Codice Metodica"] == metodica]                    
                            
                            try:
                                if metodica_catalogo["Metodica Rilevata"].values[0] not in Description_list:
                                    print("la descrizione metodica non è presente all'indice " + str(int(index)+2))
                                    flag_error = True
                            except:
                                if metodica_catalogo.size == 0:
                                    #print("print metodica_catalogo2:" + metodica_catalogo)
                                    print("exception avvennuta: controllare manualmente qual'è il problema")
                                    logging.error("controllare manualmente la metodica: " + metodica + " all'indice: " + str(int(index)+2))
                                    flag_error = True
                if flag_error == True:
                    error_dict['error_metodica_separatore'].append(str(int(index)+2))

        return error_dict


    def ck_distretti_inprestazione(self, df_mapping, sheet_Distretti, error_dict):
        print("start checking if distretti are correct")
        error_dict.update({'error_distretti_inprestazione': []})

        for index, row in df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss] == "S":
                Distretto_string = row[self.work_codice_distretto].split(",")

                siss = ""
                siss_flag = False
                if Distretto_string is not None:
                    for distretto in Distretto_string:
                        if distretto != "":
                            short_sheet = sheet_Distretti.loc[sheet_Distretti["Codice Distretto"] == distretto]                      
                            
                            #print("disciplina " + str(disciplina["Cod Disciplina"]) + " " + str(disciplina["Codice Quesito"]))
                            print("disciplina in catalogo disciplina 11:" + row[self.work_codice_prestazione_siss] + " + " + siss)
                            for cod_SISS_cat in short_sheet["Codice SISS"]: 
                                if row[self.work_codice_prestazione_siss] == cod_SISS_cat: 
                                    siss_flag = True
                                    print("disciplina in catalogo disciplina 22:" + str(row[self.work_codice_prestazione_siss]) + " + " + siss)
                                    if str(row[self.work_codice_prestazione_siss]) == siss or siss == "": # controllo se disciplina è uguale a quella precedente
                                        print("correct Distretto")
                                    else: # se non è uguale è errore
                                        print("error distretto on index:" + str(int(index)+2))
                                        error_dict['error_distretti_inprestazione'].append(str(int(index)+2))
                                else:
                                    print("disciplina diversa da quella di distretto: " + str(cod_SISS_cat))
                        else:
                            siss_flag = True

                if siss_flag == False: #se durante il mapping con la sua disciplina, questa non viene rilevata, allora è errore
                    error_dict['error_distretti_inprestazione'].append(str(int(index)+2))
                siss = str(row[self.work_codice_prestazione_siss])

        return error_dict

    def ck_distretti_sintassi(self, df_mapping, error_dict):
        print("start checking if there is ',' separator between each distretti defined")
        error_dict.update({
            'error_distretti_caratteri_non_consentiti': [],
            'error_distretti_trovato_spazio': []
        })
        string_check = re.compile('1234567890,M')

        for index, row in df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss] == "S":
                #print("Distretto: " + row["Codice Distretto"])
                if row[self.work_codice_distretto] is not None:
                    r = row[self.work_codice_distretto].strip()
                    if(string_check.search(row[self.work_codice_distretto]) != None):
                        print("String contains other Characters.")
                        error_dict['error_distretti_caratteri_non_consentiti'].append(str(int(index)+2))
                    elif " " in r: 
                        print("string contain space")
                        error_dict['error_distretti_trovato_spazio'].append(str(int(index)+2))
    
        return error_dict

    def ck_distretti_descrizione(self, df_mapping, sheet_Distretti, error_dict):
        print("start checking if distretti have the correct description")
        error_dict.update({'error_distretti_descrizione': []})

        for index, row in df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss] == "S":
                Distretto_string = row[self.work_codice_distretto].split(",")
                Description_list = row[self.work_descrizione_distretto].split(",")
                flag_error = False
                if len(Distretto_string) != len(Description_list):
                    print("il numero di descrizioni è diverso dal numero di distretti all'indice " + str(index))
                    flag_error = True

                if Distretto_string is not None:
                    for distretto in Distretto_string:
                        if distretto != "":
                            distretto_catalogo = sheet_Distretti.loc[sheet_Distretti["Codice Distretto"] == distretto]                    
                            
                            try:
                                if distretto_catalogo["Distretti"].values[0] not in Description_list:
                                    print("la descrizione distretto non è presente all'indice " + str(int(index)+2))
                                    flag_error = True
                            except:
                                if distretto_catalogo.size == 0:
                                    print("print distretto_catalogo2:" + distretto_catalogo)
                                    print("controllare manualmente qual'è il problema")
                                    logging.error("controllare manualmente il distretto: " + distretto + " all'indice: " + str(int(index)+2))
                                    flag_error = True
                if flag_error == True:
                    error_dict['error_distretti_descrizione'].append(str(int(index)+2))

        return error_dict

    def ck_distretti_operatori_logici(self, df_mapping, error_dict):
        print("start checking if there are the same logic op. for each prestazione")
        error_dict.update({'error_distretti_operatori_logici': []})

        #catalogo_dir = "c:\\Users\\aless\\exel_validate\\CCR-BO-CATGP#01_Codifiche_attributi_catalogo GP++_201910.xls"
        wb = xlrd.open_workbook(self.file_name)
        sheet_mapping = wb.sheet_by_index(0)
        print("sheet caricato")

        #problema: dovrei ordinare il file con la colonna prestazioni, ma così mi perderei 
        # l'ordine per mostrare i risultati. Stesso problema se andassi a filtrarmi le prestazioni nel file.
        # Possibile soluzione: filtro su prestazione, check OP, se è errore vado a ricercare 
        # l'indice del record.
        #last_prestazione = df_mapping['Codice Prestazione SISS'].iloc[2]
        #last_OP = df_mapping['Operatore Logico Distretto'].iloc[2]
        prestazione_checked = []
        flag_error = False

        for index, row in df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss] == "S":
                if row[self.work_operatore_logico_distretto] is not "":

                    searchedProdotto = row[self.work_codice_prestazione_siss].strip()
                    if searchedProdotto not in prestazione_checked:
                        result, result_Value = self.findCell(sheet_mapping, searchedProdotto, self.work_index_codice_prestazione_SISS)
                        print("Prestazione SISS: " + searchedProdotto + ", all'indice: " + str(int(index)+2) + " result: %s", result)
                        
                        '''if result != -1:
                            for res in result_Value:     
                                print("resultOP: " + res + ", operatore logico df_mapping: " + row[self.work_operatore_logico_distretto])
                                if row[self.work_operatore_logico_distretto] != res:
                                    flag_error = True'''
                        if result != -1:
                            for res in result:
                                r = res.split("#")[0]
                                c = res.split("#")[1]
                                resultOP = sheet_mapping.cell(int(r), self.work_index_operatore_logico_distretto).value
                                #resultOP = df_mapping['Operatore Logico Distretto'].values[int(res[0])+1]
                                print("resultOP: " + resultOP + ", operatore logico df_mapping: " + row[self.work_operatore_logico_distretto])
                                if row[self.work_operatore_logico_distretto] != resultOP:
                                    flag_error = True

                    if flag_error == True:
                        error_dict['error_distretti_operatori_logici'].append(str(int(index)+2))
                        print("error OP at index:" +  str(int(index)+2))
                    prestazione_checked.append(searchedProdotto)

        #print("error_dict: %s", error_dict)
        return error_dict   

    '''Nel caso di prestazioni prime visite, controllare se è definita almeno una priorità'''
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
        
        return error_dict

    def ck_casi_1n(self, df_mapping, error_dict):
        print("start checking if casi 1:n is correct")
        error_dict.update({'error_casi_1n': []})

        agende_list = [] #Codice SISS Agenda
        prestazioni_list = [] #Codice Prestazione SISS
        agenda_prestazione_list = []
        metodica_distretti_list = [] #lista delle metodiche e distretti delle prestazioni messe in lista
        for index, row in df_mapping.iterrows():
            #if row["Abilititazione Esposizione SISS"] == "S":
            if row[self.work_casi_1_n] != "OK":
                a_p = row[self.work_codice_agenda_siss] + "_" + row[self.work_codice_prestazione_siss]
                m_d = row[self.work_codice_metodica] + "_" + row[self.work_codice_distretto]
                if a_p not in agenda_prestazione_list and m_d not in metodica_distretti_list and row[self.work_abilitazione_esposizione_siss] == "S":
                    #print("CASO 1:N corretto momentaneamente, all'indice:" + str(int(index)+2)) 
                    #print("A_P1: " + a_p + ", Abilititazione Esposizione SISS: " + row["Abilititazione Esposizione SISS"])
                    agenda_prestazione_list.append(a_p)
                    metodica_distretti_list.append(m_d)
                elif a_p in agenda_prestazione_list and m_d in metodica_distretti_list and row[self.work_abilitazione_esposizione_siss] == "S":
                    error_dict["error_casi_1n"].append(str(int(index)+2))
                    #print("trovato caso 1:n per la coppia agenda-prestazione, all'indice: " + str(int(index)+2))
                else:
                    logging.info("trovato caso 1:n con abilitazione SISS a N corretta, all'indice: " + str(int(index)+2))
                    #print("trovato caso 1:n con abilitazione SISS a N corretta, all'indice: " + str(int(index)+2))
                    #print("A_P2: " + a_p + ", Abilititazione Esposizione SISS: " + row["Abilititazione Esposizione SISS"])
        return error_dict

    def _validation(self, error_dict):
        print("questi sono gli errori indivuduati e separati per categoria:\n %s", error_dict)
        
        '''df = pd.DataFrame(rows_list)
        with pd.ExcelWriter(self.file_name, engine='openpyxl', mode='a') as writer:
            df.to_excel(writer, sheet_name='new_mapping', index=False)'''
        print("\nPer osservare i risultati ottenuti, controllare il file prodotto: check_excel_result.txt")
        file = open("check_excel_result.txt", "w") 
        file.write(json.dumps(error_dict)) 
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
            for col in range(start_col, sh.ncols):
                myCell = sh.cell(row, col)
                myValue = sh.cell(row, self.work_index_codice_SISS_agenda) #Codice SISS agenda 15
                abilita = sh.cell(row, self.work_index_abilitazione_esposizione_SISS) #abilitazione esposizione SISS 28
                if myCell.value == searchedValue and abilita.value == "S":
                    result_coord.append(str(row) + "#" + str(col))
                    result_value.append(myValue.value)
                    #return row, col#xl_rowcol_to_cell(row, col)

        if result_coord == []:
            return -1
        return result_coord, result_value

k = Check_action()

k.import_file()
    