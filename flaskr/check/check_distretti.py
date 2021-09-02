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

class Check_distretti():

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

    
    def ck_distretti_inprestazione(self, df_mapping, sheet_Distretti, error_dict):
        print("start checking if distretti are correct")
        error_dict.update({'error_distretti_inprestazione': []})
        distretto_dict_error = {}

        for index, row in df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss] == "S":
                Distretto_string = row[self.work_codice_distretto].split(",")

                siss = ""
                siss_flag = False
                if Distretto_string is not None:
                    cod_pre_siss = str(row[self.work_codice_prestazione_siss])
                    for distretto in Distretto_string:
                        if distretto != "":
                            short_sheet = sheet_Distretti.loc[sheet_Distretti["Codice Distretto"] == distretto] #filtro catalogo sul codice distretto                     
                            
                            #print("disciplina " + str(disciplina["Cod Disciplina"]) + " " + str(disciplina["Codice Quesito"]))
                            print("prestazione del distretto in mapping 11:" + cod_pre_siss + " + " + siss)
                            
                            if cod_pre_siss not in short_sheet["Codice SISS"].values:
                                siss_flag = True
                                print("error distretto on index:" + str(int(index)+2))
                                distretto_dict_error = self.update_list_in_dict(distretto_dict_error, str(int(index)+2), distretto)
                            else:
                                if cod_pre_siss != siss and siss != "":
                                    print("error distretto on index:" + str(int(index)+2))
                                    siss_flag = True
                                    distretto_dict_error = self.update_list_in_dict(distretto_dict_error, str(int(index)+2), distretto)

                if siss_flag == True: #se durante il mapping con la sua prestazione, questa non viene rilevata, allora è errore
                    error_dict['error_distretti_inprestazione'].append(str(int(index)+2))
                siss = cod_pre_siss

        out1 = ""
        for ind in error_dict['error_distretti_inprestazione']:
            out1 = out1 + "at index: " + ind + ", on distretti: " + ", ".join(distretto_dict_error[ind]) + ", \n"
        self.output_message = self.output_message + "\nerror_distretti_inprestazione: \n" + out1
            
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
    
        out1 = ", \n".join(error_dict['error_distretti_caratteri_non_consentiti'])
        self.output_message = self.output_message + "\nerror_distretti_caratteri_non_consentiti: \n" + "at index: \n" + out1
        out2 = ", \n".join(error_dict['error_distretti_trovato_spazio'])
        self.output_message = self.output_message + "\nerror_distretti_trovato_spazio: \n" + "at index: \n" + out2
        
        return error_dict

    def ck_distretti_descrizione(self, df_mapping, sheet_Distretti, error_dict):
        print("start checking if distretti have the correct description")
        error_dict.update({'error_distretti_descrizione': []})

        distretti_dict_error = {}

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
                                    distretti_dict_error = self.update_list_in_dict(distretti_dict_error, str(int(index)+2), distretto)
                            except:
                                if distretto_catalogo.size == 0:
                                    print("print distretto_catalogo2:" + distretto_catalogo)
                                    print("controllare manualmente qual'è il problema")
                                    logging.error("controllare manualmente il distretto: " + distretto + " all'indice: " + str(int(index)+2))
                                    flag_error = True
                                    distretti_dict_error = self.update_list_in_dict(distretti_dict_error, str(int(index)+2), distretto)
                if flag_error == True:
                    error_dict['error_distretti_descrizione'].append(str(int(index)+2))

        out1 = ""
        for ind in error_dict['error_distretti_descrizione']:
            out1 = out1 + "at index: " + ind + ", on distretto: " + ", ".join(distretti_dict_error[ind]) + ", \n"
        self.output_message = self.output_message + "\nerror_distretti_descrizione: \n" + out1
        
        return error_dict

    def ck_distretti_operatori_logici(self, df_mapping, error_dict):
        print("start checking if there are the same logic op. for each prestazione")
        error_dict.update({
            'error_distretti_operatori_logici': [],
            'error_distretti_operatori_logici_mancante': [] })
        distretti_dict_error = {}
        #catalogo_dir = "c:\\Users\\aless\\exel_validate\\CCR-BO-CATGP#01_Codifiche_attributi_catalogo GP++_201910.xls"
        wb = None
        if self.file_name != "":
            wb = xlrd.open_workbook(self.file_name)
        else:
            wb = xlrd.open_workbook(file_contents=self.file_data) #file_contents=self.file_data.read()
        sheet_mapping = wb.sheet_by_index(self.work_index_sheet)
        print("sheet caricato")

        #problema: dovrei ordinare il file con la colonna prestazioni, ma così mi perderei 
        # l'ordine per mostrare i risultati. Stesso problema se andassi a filtrarmi le prestazioni nel file.
        # Possibile soluzione: filtro su prestazione, check OP, se è errore vado a ricercare 
        # l'indice del record.
        #last_prestazione = df_mapping['Codice Prestazione SISS'].iloc[2]
        #last_OP = df_mapping['Operatore Logico Distretto'].iloc[2]
        prestazione_checked = []

        for index, row in df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss] == "S":
                if row[self.work_operatore_logico_distretto] is not "":
                    cod_pre_siss = str(row[self.work_codice_prestazione_siss])
                    searchedProdotto = cod_pre_siss.strip()
                    if searchedProdotto not in prestazione_checked:
                        result = self.findCell(sheet_mapping, searchedProdotto, self.work_index_codice_prestazione_SISS)
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
                                    distretti_dict_error = self.update_list_in_dict(distretti_dict_error, str(int(r)+1), "OP diversi per le prestazioni:" + searchedProdotto)
                                    error_dict['error_distretti_operatori_logici'].append(str(int(r)+1))
                                    print("error OP at index:" +  str(int(r)+1))
                    prestazione_checked.append(searchedProdotto)

                elif row[self.work_operatore_logico_distretto] is "" and row[self.work_codice_distretto] is not "": 
                    error_dict['error_distretti_operatori_logici_mancante'].append(str(int(index)+2))
                    
        
        out1 = ""
        for ind in error_dict['error_distretti_operatori_logici']:
            out1 = out1 + "at index: " + ind + ", on distretto: " + ", ".join(distretti_dict_error[ind]) + ", \n"
        self.output_message = self.output_message + "\nerror_distretti_operatori_logici: \n" + "at index: \n" + out1
        out2 = ", \n".join(error_dict['error_distretti_operatori_logici_mancante'])
        self.output_message = self.output_message + "\nerror_distretti_operatori_logici_mancante: \n" + "at index: \n" + out2
        return error_dict   