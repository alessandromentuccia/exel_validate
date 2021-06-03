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
        df_mapping = pd.read_excel(template_file, 0, converters={'question_number': str}).replace(np.nan, '', regex=True)
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
        metodiche_error = self.check_metodiche(df_mapping)
        print("Fase 3") #FASE 3: CONTROLLO I DISTRETTI
        distretti_error = self.check_distretti(df_mapping)
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
        error_QD_disciplina_agenda = self.ck_QD_disciplina_agenda(df_mapping, sheet_QD)
        error_QD_separatore = self.ck_QD_separatore(df_mapping)

        error_list = {
            "error_QD_agenda": error_QD_agenda,
            "error_QD_disciplina_agenda": error_QD_disciplina_agenda,
            "error_QD_separatore": error_QD_separatore
        }

        return error_list

    def check_metodiche(self, df_mapping):
        print("start checking Metodiche")

    def check_distretti(self, df_mapping):
        print("start checking Distretti")
    
    def check_priorita(self, df_mapping):
        print("start checking priorità e tipologie di accesso")

    def ck_QD_agenda(self, df_mapping):
        print("start checking if foreach agenda there are the same QD")

        error_list = []
        
        agenda = df_mapping['Codice SISS Agenda'].iloc[2]
        last_QD = df_mapping['Codice Quesito Diagnostico'].iloc[2]
        for index, row in df_mapping.iterrows():
            if row["Codice SISS Agenda"] is agenda:
                print("- same agenda -")
                if row["Codice Quesito Diagnostico"] is last_QD:
                    print("correct QD")
                else: 
                    print("error QD")
                    error_list.append(str(index))
            else:
                print("- the agenda is changed -")
                agenda = row["Codice SISS Agenda"]
                last_QD = row["Codice Quesito Diagnostico"]

        print("error_list: %s", error_list)
        return error_list

    def ck_QD_disciplina_agenda(self, df_mapping, sheet_):
        print("start checking if foreach agenda there is the same Disciplina for all the QD")
        error_list = []

        return error_list
            
    def ck_QD_separatore(self, df_mapping):
        print("start checking if there is ',' separator between each QD defined")
        error_list = []

        return error_list

    def _validation(self):
        # reproducible randomization in future runs
        reproducible_random = random.Random(1)
        examples_count = 0

        #define the first raw
        rows_list = {
                        "question_type" : [], 
                        "question_number" : [], 
                        "question" : [], 
                        "answer" : [], 
                        "note" : []
                    }

        '''for intent_name, phrase_templates in intent_templates.items():
                logger.info(f"Generating examples of intent '{intent_name}' ...")
                for template in phrase_templates:
                    #output_file.write(f"<!-- {template.original_line} -->\n")
                    expanded_slot_lists = self.__expand_references_inside_slots(template)
                        
                    
                    # fill strings with slots
                    filled_template_list = []
                    for filling_words in filling_words_list:
                        filled_template_list.append(
                            template.string_to_fill.format(*filling_words))
                    # print on file
                    for filled in sorted(filled_template_list):
                        rows_list["question_type"].append(self.file_data[intent_name][0]) 
                        rows_list["question_number"].append(intent_name)
                        rows_list["question"].append(filled)
                        rows_list["answer"].append(self.file_data[intent_name][3])
                        rows_list["note"].append(self.file_data[intent_name][4])
                    logger.info(
                        f"  {len(filled_template_list)} examples of template \"{template.original_line}\"")
                    examples_count += len(filled_template_list)
        logger.info(f"Generation completed - total examples: {examples_count}")

        df = pd.DataFrame(rows_list)
        with pd.ExcelWriter(self.file_name, engine='openpyxl', mode='a') as writer:
            df.to_excel(writer, sheet_name='new_mapping', index=False)'''


k = knowledge_action()

k.import_file()
