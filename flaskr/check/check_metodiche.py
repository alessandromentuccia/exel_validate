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

class Check_metodiche():

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

    def ck_metodica_inprestazione(self, df_mapping, sheet_Metodiche, error_dict):
        print("start checking if metodica are correct")
        error_dict.update({'error_metodica_inprestazione': []})
        #self.output_message = self.output_message + "\nerror_metodica_inprestazione: "
        metodica_dict_error = {}
        for index, row in df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss] == "S":
                Metodica_string = row[self.work_codice_metodica].split(",")

                siss = "" 
                siss_flag = False
                if Metodica_string is not None:
                    cod_pre_siss = str(row[self.work_codice_prestazione_siss])
                    for metodica in Metodica_string:
                        if metodica != "":
                            short_sheet = sheet_Metodiche.loc[sheet_Metodiche["Codice Metodica"] == metodica]                      
                            
                            #print("disciplina " + str(disciplina["Cod Disciplina"]) + " " + str(disciplina["Codice Quesito"]))
                            print("prestazione della metodica in mapping:" + cod_pre_siss + " + " + siss)
                            
                            if cod_pre_siss not in short_sheet["Codice SISS"].values:
                                siss_flag = True
                                metodica_dict_error = self.update_list_in_dict(metodica_dict_error, str(int(index)+2), metodica)
                                print("error metodica on index:" + str(int(index)+2))
                            else:
                                if cod_pre_siss != siss and siss != "":
                                    print("error metodica on index:" + str(int(index)+2))
                                    siss_flag = True
                                    metodica_dict_error = self.update_list_in_dict(metodica_dict_error, str(int(index)+2), metodica)

                            '''for cod_SISS_cat in short_sheet["Codice SISS"]: 
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
                            siss_flag = True'''

                if siss_flag == True: #se durante il mapping con la sua disciplina, questa non viene rilevata, allora è errore
                    error_dict['error_metodica_inprestazione'].append(str(int(index)+2))
                siss = str(row[self.work_codice_prestazione_siss])
        out1 = ""
        for ind in error_dict['error_metodica_inprestazione']:
            out1 = out1 + "at index: " + ind + ", on metodica: " + ", ".join(metodica_dict_error[ind]) + ", \n"
        
        self.output_message = self.output_message + "\nerror_metodica_inprestazione: \n" + out1
            
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

        out1 = ", \n".join(error_dict['error_metodica_caratteri_non_consentiti'])
        self.output_message = self.output_message + "\nerror_metodica_caratteri_non_consentiti: \n" + "at index: \n" + out1
        out2 = ", \n".join(error_dict['error_metodica_spazio_bordi'])
        self.output_message = self.output_message + "\nerror_metodica_spazio_bordi: \n" + "at index: \n" + out2
        out3 = ", \n".join(error_dict['error_metodica_spazio_internamente'])
        self.output_message = self.output_message + "\nerror_metodica_spazio_internamente: \n" + "at index: \n" + out3
            
        return error_dict
    
    def ck_metodica_descrizione(self, df_mapping, sheet_Metodiche, error_dict):
        print("start checking if metodica have the correct description")
        error_dict.update({
            'error_metodica_descrizione': [],
            'error_metodica_separatore': []
            })
        metodica_dict_error = {}
        
        for index, row in df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss] == "S":
                Metodica_string = row[self.work_codice_metodica].split(",")
                Description_list = row[self.work_descrizione_metodica].split(",")
                flag_error = False
                if len(Metodica_string) != len(Description_list):
                    print("il numero di descrizioni è diverso dal numero di metodiche all'indice " + str(index))
                    flag_error = True
                    metodica_dict_error = self.update_list_in_dict(metodica_dict_error, str(int(index)+2), "Manca un Codice Metodica o una Descrizione")

                if Metodica_string is not None:
                    for metodica in Metodica_string:
                        if metodica != "":
                            metodica_catalogo = sheet_Metodiche.loc[sheet_Metodiche["Codice Metodica"] == metodica]                    
                            
                            try:
                                if metodica_catalogo["Metodica Rilevata"].values[0] not in Description_list:
                                    print("la descrizione metodica non è presente all'indice " + str(int(index)+2))
                                    flag_error = True
                                    metodica_dict_error = self.update_list_in_dict(metodica_dict_error, str(int(index)+2), metodica)
                            except:
                                if metodica_catalogo.size == 0:
                                    #print("print metodica_catalogo2:" + metodica_catalogo)
                                    print("exception avvennuta: controllare manualmente qual'è il problema")
                                    logging.error("controllare manualmente la metodica: " + metodica + " all'indice: " + str(int(index)+2))
                                    flag_error = True
                                    metodica_dict_error = self.update_list_in_dict(metodica_dict_error, str(int(index)+2), metodica)
                if flag_error == True:
                    error_dict['error_metodica_descrizione'].append(str(int(index)+2))

        out1 = ""
        for ind in error_dict['error_metodica_descrizione']:
            out1 = out1 + "at index: " + ind + ", on metodica: " + ", ".join(metodica_dict_error[ind]) + ", \n"
        
        #out1 = ", \n".join(error_dict['error_metodica_descrizione'])
        self.output_message = self.output_message + "\nerror_metodica_descrizione: \n" + out1
        out2 = ", \n".join(error_dict['error_metodica_separatore'])
        self.output_message = self.output_message + "\nerror_metodica_separatore: \n" + out2

        return error_dict