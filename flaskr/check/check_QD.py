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

class Check_QD():

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
    work_index_codici_descri_disciplina_catalogo = 0

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
        self.work_index_codici_descri_disciplina_catalogo = data[1]["work_index"]["work_index_codici_descri_disciplina_catalogo"]

    def ck_QD_agenda(self, df_mapping, error_dict):
        print("start checking if foreach agenda there are the same QD")
        
        error_dict.update({'error_QD_agenda': []})
        QD_dict_error = {}
        
        agenda = df_mapping[self.work_codice_agenda_siss].iloc[2]
        last_QD = df_mapping[self.work_codice_QD].iloc[2]
        for index, row in df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss] == "S":
                if row[self.work_codice_agenda_siss] == agenda:
                    if row[self.work_codice_QD] != last_QD:
                        print("error QD at index:" +  str(int(index)+2))
                        #error_list.append(str(int(index)+2))
                        error_dict['error_QD_agenda'].append(str(int(index)+2))
                        QD_dict_error = self.update_list_in_dict(QD_dict_error, str(int(index)+2), "il QD: "+ row[self.work_codice_QD] + "è diverso per la stessa agenda")
                else: 
                    agenda = row[self.work_codice_agenda_siss]
                    last_QD = row[self.work_codice_QD]

        out1 = ""
        for ind in error_dict['error_QD_agenda']:
            out1 = out1 + "at index: " + ind + ", ".join(QD_dict_error[ind]) + ", \n"
        self.output_message = self.output_message + "\nerror_QD_agenda: \n" + out1

        return error_dict

    def ck_QD_disciplina_agenda(self, df_mapping, sheet_QD, error_dict):
        print("start checking if foreach agenda there is the same Disciplina for all the QD")
        #tutti i QD di un agenda hanno la stessa disciplina
        error_dict.update({
            'error_QD_disciplina_agenda': [],
            'error_QD_descrizione_disciplina_agenda': [],
            'error_disciplina_mancante' : [],
            'error_discipline_agende_diverse': []
        }) 
        QD_disci_dict_error = {}
        QD_descri_disci_dict_error = {}
        agenda_disci_dict_error = {}

        wb = xlrd.open_workbook(self.file_name)
        sheet_mapping = wb.sheet_by_index(self.work_index_sheet)
        print("sheet caricato")
        
        #disciplina_QD_column = sheet_QD[['Cod Disciplina','Codice Quesito']]
        #print("disciplina_QD_column: %s", disciplina_QD_column)
        agende_viewed = []
        for index, row in df_mapping.iterrows():
            disci_flag_QD = False
            descri_disci_flag_QD = False
            disci_flag_agenda = False
            if row[self.work_codice_agenda_siss] not in agende_viewed and row[self.work_codice_QD] is not None:
                searchedAgenda = row[self.work_codice_agenda_siss]
                disciplina_mapping_row = row[self.work_codici_disciplina_catalogo]
                descrizione_disciplina_mapping_row = row[self.work_descrizione_disciplina_catalogo]
                #prendo tutte le agende con lo stesso codice
                result = self.findCell_agenda(sheet_mapping, searchedAgenda, self.work_index_codice_SISS_agenda) #prendo tutte le righe con questa agenda

                if result != -1:
                    result_disciplina_last = ""
                    agende_error_list = []
                    descrizione_discipline_error_list = []
                    discipline_error_list = []
                    for res in result: #per ogni risultato controllo che ci sia la stessa disciplina
                        r = res.split("#")[0] #row agenda
                        c = res.split("#")[1] #column agenda

                        result_QD = sheet_mapping.cell(int(r), self.work_index_codice_QD-1).value.split(",") #QD
                        if result_QD is not None:
                            for QD in result_QD:
                                if QD != "":
                                    short_sheet = sheet_QD.loc[sheet_QD["Codice Quesito"] == QD.strip()]
                                    if disciplina_mapping_row not in short_sheet["Cod Disciplina"].values:
                                        disci_flag_QD = True
                                        if str(int(r)+1) not in agende_error_list:
                                            agende_error_list.append(str(int(r)+1))
                                            QD_disci_dict_error[str(int(r)+1)] = "per il QD: "+ QD + " non c'è la disciplina: " + disciplina_mapping_row
                                        print("disciplina_mapping_row: " + disciplina_mapping_row + ", QD not in disciplina: " + QD)
                                    if descrizione_disciplina_mapping_row not in short_sheet["Descrizione disciplina"].values:
                                        descri_disci_flag_QD = True
                                        if str(int(r)+1) not in descrizione_discipline_error_list:
                                            descrizione_discipline_error_list.append(str(int(r)+1))
                                            QD_descri_disci_dict_error[str(int(r)+1)] = "per il QD: "+ QD + " non c'è la descrizione disciplina: " + descrizione_disciplina_mapping_row

                        result_disciplina = sheet_mapping.cell(int(r), self.work_index_codici_disciplina_catalogo).value #disciplina da catalogo
                        if result_disciplina != "":
                            if result_disciplina_last != "": #se non è la prima iterazione
                                if result_disciplina != result_disciplina_last:
                                    disci_flag_agenda = True
                                    discipline_error_list.append(str(int(r)+1))
                                    agenda_disci_dict_error = self.update_list_in_dict(agenda_disci_dict_error, str(int(r)+1), "agenda con discipline diverse: "+ result_disciplina + " e " + result_disciplina_last)
                                    print("result_disciplina: " + result_disciplina + ", result_disciplina_last: " + result_disciplina_last)
                                else: 
                                    result_disciplina_last = result_disciplina    
                
                    if disci_flag_QD == True: #se la disciplina non è rilevata nel catalogo allora è errore
                        for age in agende_error_list:
                            error_dict['error_QD_disciplina_agenda'].append(age)
                            
                    if disci_flag_agenda == True: #se la disciplina è diversa in una stessa agenda, allora è errore
                        for age in discipline_error_list:
                            error_dict['error_discipline_agende_diverse'].append(age)
                else:
                    error_dict['error_disciplina_mancante'].append(str(int(index)+2))     #inserisco la riga senza disciplina negli errori

        out1 = ""
        for ind in error_dict['error_QD_disciplina_agenda']:
            out1 = out1 + "at index: " + ind + ", error: " + QD_disci_dict_error[ind] + ", \n"
        out2 = ""
        for ind in error_dict['error_discipline_agende_diverse']:
            out2 = out2 + "at index: " + ind + ", error: " + ", ".join(agenda_disci_dict_error[ind]) + ", \n"
        out3 = ""
        for ind in error_dict['error_disciplina_mancante']:
            out3 = out3 + "at index: " + ind + "disciplina mancante, \n"
        out4 = ""
        for ind in error_dict['error_discipline_agende_diverse']:
            out4 = out4 + "at index: " + ind + ", error: " + QD_descri_disci_dict_error[ind] + ", \n"
            
    
        self.output_message = self.output_message + "\nerror_QD_disciplina_agenda: \n" + out1
        self.output_message = self.output_message + "\nerror_discipline_agende_diverse: \n" + out2
        self.output_message = self.output_message + "\nerror_disciplina_mancante: \n" + out3
        self.output_message = self.output_message + "\nerror_discipline_agende_diverse: \n" + out4

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
            if row[self.work_abilitazione_esposizione_siss] != None: 
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

        return error_dict