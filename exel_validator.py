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

    def import_file(self):
        logging.warning("import CSV")

        template_file = input("Enter your mapping file.xlsx: ") ##insert mapping file
        print(template_file) 
        template_file = Path(template_file)
        self.file_name = template_file
        if template_file.is_file(): #C:\Users\aless\csi-progetti\FaqBot\faqbot-09112020.xlsx
            self.read_exel_file(template_file)
        else:
            print("Il file non esiste, prova a ricaricare il file con la directory corretta.\n")

    def read_exel_file(self, template_file):
        df = pd.read_excel(template_file, '0', converters={'question_number': str}).replace(np.nan, '', regex=True)
        #print ("print JSON")
        #print(sh)
        data_list = {'question_number': {}}
        
        for index, row in df.iterrows():
            data = OrderedDict()
            if row["question_number"] not in data_list["question_number"]:
                data[row[1]] = [row["question_template"]]
                self.file_data.update({ row["question_number"]: [row["question_type"], row["question_number"], row["question_template"], row["answer"], row["note"]]})
                data_list["question_number"].update(data)
                print("data_list: %s", data_list)
            else: 
                data_list["question_number"][row["question_number"]].append(row["question_template"])
                self.file_data.update({ row["question_number"]: [row["question_type"], row["question_number"], row["question_template"], row["answer"], row["note"]]})

        #with open("RulesJson.json", "w", encoding="utf-8") as writeJsonfile:
        #    json.dump(data_list, writeJsonfile, indent=4,default=str) 
        logging.warning("initiate generation")
        self.analizer(data_list)

    def analizer(self, template_file):

        print('Start analisys:\n', template_file)

        parsed_intent_templates = []

        self.__build_examples(parsed_intent_templates)

 
    def __build_examples(self, intent_templates: Dict[str, List[PhraseTemplate]]) -> List[str]:
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

        for intent_name, phrase_templates in intent_templates.items():
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
            df.to_excel(writer, sheet_name='new_mapping', index=False)


k = knowledge_action()

k.import_file()
