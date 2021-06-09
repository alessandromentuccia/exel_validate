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
#import xlrd
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

EXAMPLE_FILE_PATH = 'generated_examples.md'

class PhraseTemplate(object):
    def __init__(self, original_line: str, string_to_fill: str, slot_list: List[List[str]]):
        self.original_line = original_line
        self.string_to_fill = string_to_fill
        self.slot_list = slot_list


class knowledge_action():

    file_name = ""
    file_data = {}
    catalogo = OrderedDict()
    flag_check_list = []

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
        df_mapping = pd.read_excel(template_file, 0, converters={'Disciplina Agenda': str}).replace(np.nan, '', regex=True)
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

        print('Start analisys:\n', df_mapping)

        print("Fase 1") #FASE 1: CONTROLLO I QUESITI DIAGNOSTICI
        QD_error = self.check_qd(df_mapping, sheet_QD)
        print("Fase 2") #FASE 2: CONTROLLO LE METODICHE
        metodiche_error = self.check_metodiche(df_mapping, sheet_Metodiche)
        print("Fase 3") #FASE 3: CONTROLLO I DISTRETTI
        distretti_error = self.check_distretti(df_mapping, sheet_Distretti)
        print("Fase 4") #FASE 4: CONTROLLO LE PRIORITA'
        priorita_error = self.check_priorita(df_mapping)

        error_dict = {
            "QD_error": QD_error,
            "metodiche_error": metodiche_error,
            "distretti_error": distretti_error,
            "priorita_error": priorita_error,    
        }

        self._validation(error_dict)

    def check_qd(self, df_mapping, sheet_QD):
        print("start checking QD") #Codice Quesito Diagnostico
        #controllo se per ogni Agenda sono inseriti gli stessi QD
        
        error_QD_agenda = self.ck_QD_agenda(df_mapping)
        error_QD_separatore = self.ck_QD_separatore(df_mapping)
        error_QD_disciplina_agenda = self.ck_QD_disciplina_agenda(df_mapping, sheet_QD)
        error_QD_disciplina_descrizione = self.ck_QD_disciplina_descrizione(df_mapping, sheet_QD)
        

        error_list = {
            "error_QD_agenda": error_QD_agenda,
            "error_QD_disciplina_agenda": error_QD_disciplina_agenda,
            "error_QD_separatore": error_QD_separatore,
            "error_QD_disciplina_descrizione": error_QD_disciplina_descrizione
        }

        return error_list

    def check_metodiche(self, df_mapping, sheet_Metodiche):
        print("start checking Metodiche")

        error_metodica_inprestazione = self.ck_metodica_inprestazione(df_mapping, sheet_Metodiche)
        error_metodica_separatore = self.ck_metodica_separatore(df_mapping)
        error_metodica_descrizione = self.ck_metodica_descrizione(df_mapping, sheet_Metodiche)

        error_list = {
            "error_metodica_inprestazione": error_metodica_inprestazione,
            "error_metodica_separatore": error_metodica_separatore,
            "error_metodica_descrizione": error_metodica_descrizione
        }

        print("error_list: %s", error_list)
        return error_list

    def check_distretti(self, df_mapping, sheet_Distretti):
        print("start checking Distretti")

        error_distretti_inprestazione = self.ck_distretti_inprestazione(df_mapping, sheet_Distretti)
        error_distretti_separatore = self.ck_distretti_separatore(df_mapping)
        error_distretti_descrizione = self.ck_distretti_descrizione(df_mapping, sheet_Distretti)

        error_list = {
            "error_distretti_inprestazione": error_distretti_inprestazione,
            "error_distretti_separatore": error_distretti_separatore,
            "error_distretti_descrizione": error_distretti_descrizione
        }

        print("error_list: %s", error_list)
        return error_list

    def check_priorita(self, df_mapping):
        print("start checking priorità e tipologie di accesso")

        error_prime_visite = self.ck_prime_visite(df_mapping)
        error_controlli = self.ck_controlli(df_mapping)
        error_esami_strumentali =self.ck_esami_strumentali(df_mapping)

        error_list = {
            "error_prime_visite": error_prime_visite,
            "error_controlli": error_controlli,
            "error_esami_strumentali": error_esami_strumentali
        }

        print("error_list: %s", error_list)
        return error_list

    def ck_QD_agenda(self, df_mapping):
        print("start checking if foreach agenda there are the same QD")

        error_list = []
        
        agenda = df_mapping['Codice SISS Agenda'].iloc[2]
        last_QD = df_mapping['Codice Quesito Diagnostico'].iloc[2]
        for index, row in df_mapping.iterrows():
            if row["Codice SISS Agenda"] is agenda:
                #print("- same agenda -")
                if row["Codice Quesito Diagnostico"] is not last_QD:
                    print("error QD")
                    error_list.append(str(index))
                #else: 
                #    print("correct QD")
            else: 
                #print("- the agenda is changed -")
                agenda = row["Codice SISS Agenda"]
                last_QD = row["Codice Quesito Diagnostico"]

        print("error_list: %s", error_list)
        return error_list

    def ck_QD_disciplina_agenda(self, df_mapping, sheet_QD):
        print("start checking if foreach agenda there is the same Disciplina for all the QD")
        error_list = [] 
        
        #disciplina_QD_column = sheet_QD[['Cod Disciplina','Codice Quesito']]
        #print("disciplina_QD_column: %s", disciplina_QD_column)
        for index, row in df_mapping.iterrows():
            QD_string = row["Codice Quesito Diagnostico"].split(",")
            disci = ""
            disci_flag = False
            if QD_string is not None:
                for QD in QD_string:
                    if QD != "":
                        disciplina = sheet_QD.loc[sheet_QD["Codice Quesito"] == QD]  #[str(i) for i in                      
                        #print("QD: " + QD)
                        #print("disciplina " + str(disciplina["Cod Disciplina"]) + " " + str(disciplina["Codice Quesito"]))
                        print("disciplina in catalogo disciplina 11:" + row["Disciplina Agenda"] + " + " + disci)
                        for d in disciplina["Cod Disciplina"]: 
                            if row["Disciplina Agenda"] == d: 
                                disci_flag = True
                                print("disciplina in catalogo disciplina 22:" + str(row["Disciplina Agenda"]) + " + " + disci)
                                if str(row["Disciplina Agenda"]) == disci or disci == "": # controllo se disciplina è uguale a quella precedente
                                    print("correct Disciplina")
                                else: # se non è uguale è errore
                                    print("error QD on index:" + str(int(index)+2))
                                    error_list.append(str(int(index)+2))
                            else:
                                print("disciplina diversa da quella di QD: " + str(d))
                    else:
                        disci_flag = True

            if disci_flag == False: #se durante il mapping con la sua disciplina, questa non viene rilevata, allora è errore
                error_list.append(str(int(index)+2))
            disci = str(row["Disciplina Agenda"])

        return error_list
            
    def ck_QD_separatore(self, df_mapping):
        print("start checking if there is ',' separator between each QD defined")
        error_list = []
        string_check = re.compile('1234567890,Q')

        for index, row in df_mapping.iterrows():
            print("QD: " + row["Codice Quesito Diagnostico"])
            flag_error = False
            if row["Codice Quesito Diagnostico"] is not None:
                if(string_check.search(row["Codice Quesito Diagnostico"]) == None):
                    print("String does not contain other characters") 
                else: 
                    print("String contains other Characters.")
                    flag_error = True

            if flag_error == True:
                error_list.append(str(int(index)+2))

        return error_list

    def ck_QD_disciplina_descrizione(self, df_mapping, sheet_QD):
        print("start checking if there are the relative QD description")
        error_list = []
        
        for index, row in df_mapping.iterrows():
            QD_string = row["Codice Quesito Diagnostico"].split(",")
            Description_list = row["Descrizione Quesito Diagnostico"].split(",")
            flag_error = False
            if len(QD_string) != len(Description_list):
                print("il numero di descrizioni è diverso dal numero di QD all'indice " + str(index))
                flag_error = True

            if QD_string is not None:
                for QD in QD_string:
                    if QD != "":
                        QD_catalogo = sheet_QD.loc[sheet_QD["Codice Quesito"] == QD]  
                        
                        #print("QD: " + str(QD)) 
                        try:
                            if QD_catalogo["Quesiti Diagnostici"].values[0] not in Description_list:
                                print("la descrizione QD non è presente all'indice " + str(int(index)+2))
                                flag_error = True
                        except:
                            if QD_catalogo.size == 0:
                                print("print QD_catalogo2:" + QD_catalogo)
                                print("controllare manualmente qual'è il problema")
                                logging.error("controllare manualmente il QD: " + QD + " all'indice: " + str(int(index)+2))

                        
                        #print("QD_catalogo: " + QD_catalogo["Quesiti Diagnostici"].values[0])     
                        #print("Description_list: %s", Description_list)            
                        

            if flag_error == True:
                error_list.append(str(int(index)+2))

        return error_list

    def ck_metodica_inprestazione(self, df_mapping, sheet_Metodiche):
        print("start checking if metodica are correct")
        error_list = []
        
        for index, row in df_mapping.iterrows():
            Metodica_string = row["Codice Metodica"].split(",")

            disci = ""
            disci_flag = False
            if Metodica_string is not None:
                for metodica in Metodica_string:
                    if metodica != "":
                        disciplina = sheet_Metodiche.loc[sheet_Metodiche["Codice Metodica"] == metodica]                      
                        
                        #print("disciplina " + str(disciplina["Cod Disciplina"]) + " " + str(disciplina["Codice Quesito"]))
                        print("disciplina in catalogo disciplina 11:" + row["Disciplina Agenda"] + " + " + disci)
                        for d in disciplina["Cod Disciplina"]: 
                            if row["Disciplina Agenda"] == d: 
                                disci_flag = True
                                print("disciplina in catalogo disciplina 22:" + str(row["Disciplina Agenda"]) + " + " + disci)
                                if str(row["Disciplina Agenda"]) == disci or disci == "": # controllo se disciplina è uguale a quella precedente
                                    print("correct Disciplina")
                                else: # se non è uguale è errore
                                    print("error metodica on index:" + str(int(index)+2))
                                    error_list.append(str(int(index)+2))
                            else:
                                print("disciplina diversa da quella di metodica: " + str(d))
                    else:
                        disci_flag = True

            if disci_flag == False: #se durante il mapping con la sua disciplina, questa non viene rilevata, allora è errore
                error_list.append(str(int(index)+2))
            disci = str(row["Disciplina Agenda"])

        return error_list

    def ck_metodica_separatore(self, df_mapping):
        print("start checking if there is ',' separator between each metodiche defined")
        error_list = []
        string_check = re.compile('1234567890,M')

        for index, row in df_mapping.iterrows():
            print("Metodica: " + row["Codice Metodica"])
            flag_error = False
            if row["Codice Metodica"] is not None:
                if(string_check.search(row["Codice Metodica"]) == None):
                    print("String does not contain other characters") 
                else: 
                    print("String contains other Characters.")
                    flag_error = True

            if flag_error == True:
                error_list.append(str(int(index)+2))

        return error_list

    def ck_metodica_descrizione(self, df_mapping, sheet_Metodiche):
        print("start checking if metodica have the correct description")
        error_list = []
        
        for index, row in df_mapping.iterrows():
            Metodica_string = row["Codice Metodica"].split(",")
            Description_list = row["Descrizione Metodica"].split(",")
            flag_error = False
            if len(Metodica_string) != len(Description_list):
                print("il numero di descrizioni è diverso dal numero di metodiche all'indice " + str(index))
                flag_error = True

            if Metodica_string is not None:
                for metodica in Metodica_string:
                    if metodica != "":
                        metodica_catalogo = sheet_Metodiche.loc[sheet_Metodiche["Codice Metodica"] == metodica]                    
                        if metodica_catalogo["Metodica Rilevata"].values[0] not in Description_list:
                            print("la descrizione metodica non è presente all'indice " + str(index))
                            flag_error = True

            if flag_error == True:
                error_list.append(str(int(index)+2))

        return error_list

    def ck_distretti_inprestazione(self, df_mapping, sheet_Distretti):
        print("start checking if distretti are correct")
        error_list = []

        for index, row in df_mapping.iterrows():
            Distretto_string = row["Codice Distretto"].split(",")

            disci = ""
            disci_flag = False
            if Distretto_string is not None:
                for distretto in Distretto_string:
                    if distretto != "":
                        disciplina = sheet_Distretti.loc[sheet_Distretti["Codice Distretto"] == distretto]                      
                        
                        #print("disciplina " + str(disciplina["Cod Disciplina"]) + " " + str(disciplina["Codice Quesito"]))
                        print("disciplina in catalogo disciplina 11:" + row["Disciplina Agenda"] + " + " + disci)
                        for d in disciplina["Cod Disciplina"]: 
                            if row["Disciplina Agenda"] == d: 
                                disci_flag = True
                                print("disciplina in catalogo disciplina 22:" + str(row["Disciplina Agenda"]) + " + " + disci)
                                if str(row["Disciplina Agenda"]) == disci or disci == "": # controllo se disciplina è uguale a quella precedente
                                    print("correct Disciplina")
                                else: # se non è uguale è errore
                                    print("error distretto on index:" + str(int(index)+2))
                                    error_list.append(str(int(index)+2))
                            else:
                                print("disciplina diversa da quella di distretto: " + str(d))
                    else:
                        disci_flag = True

            if disci_flag == False: #se durante il mapping con la sua disciplina, questa non viene rilevata, allora è errore
                error_list.append(str(int(index)+2))
            disci = str(row["Disciplina Agenda"])

        return error_list

    def ck_distretti_separatore(self, df_mapping):
        print("start checking if there is ',' separator between each distretti defined")
        error_list = []
        string_check = re.compile('1234567890,M')

        for index, row in df_mapping.iterrows():
            print("Distretto: " + row["Codice Distretto"])
            flag_error = False
            if row["Codice Distretto"] is not None:
                if(string_check.search(row["Codice Distretto"]) == None):
                    print("String does not contain other characters") 
                else: 
                    print("String contains other Characters.")
                    flag_error = True

            if flag_error == True:
                error_list.append(str(int(index)+2))

        return error_list

    def ck_distretti_descrizione(self, df_mapping, sheet_Distretti):
        print("start checking if distretti have the correct description")
        error_list = []

        for index, row in df_mapping.iterrows():
            Distretto_string = row["Codice Distretto"].split(",")
            Description_list = row["Descrizione Distretto"].split(",")
            flag_error = False
            if len(Distretto_string) != len(Description_list):
                print("il numero di descrizioni è diverso dal numero di distretti all'indice " + str(index))
                flag_error = True

            if Distretto_string is not None:
                for distretto in Distretto_string:
                    if distretto != "":
                        distretto_catalogo = sheet_Distretti.loc[sheet_Distretti["Codice Distretto"] == distretto]                    
                        if distretto_catalogo["Distretti"].values[0] not in Description_list:
                            print("la descrizione distretto non è presente all'indice " + str(index))
                            flag_error = True

            if flag_error == True:
                error_list.append(str(int(index)+2))

        return error_list

    def ck_prime_visite(self, df_mapping):
        error_list = []
        return error_list

    def ck_controlli(self, df_mapping):
        error_list = []
        return error_list

    def ck_esami_strumentali(self, df_mapping):
        error_list = []
        return error_list

    def _validation(self, error_list):
        print("questi sono gli errori indivuduati e separati per categoria:\n %s", error_list)
        
        '''df = pd.DataFrame(rows_list)
        with pd.ExcelWriter(self.file_name, engine='openpyxl', mode='a') as writer:
            df.to_excel(writer, sheet_name='new_mapping', index=False)'''


k = knowledge_action()

k.import_file()
