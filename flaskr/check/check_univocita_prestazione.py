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
        distretti_dict = {}
        
        index_dict = {}
        for index, row in df_mapping.iterrows():
            a_p = str(row[self.work_codice_agenda_siss]) + "_" + str(row[self.work_codice_prestazione_siss])
            m_d = row[self.work_codice_metodica] + "_" + row[self.work_codice_distretto]
            index_dict = self.update_list_in_dict(index_dict, a_p, str(int(index)+2))
            distretti_dict = self.update_list_in_dict(distretti_dict, str(int(index)+2), row[self.work_operatore_logico_distretto])
            if a_p not in metodica_distretti_dict.keys() and row[self.work_abilitazione_esposizione_siss] == "S":
                #primo elemento inserito nel dict
                metodica_distretti_dict = self.update_list_in_dict(metodica_distretti_dict, a_p, m_d)
            elif a_p in metodica_distretti_dict.keys() and row[self.work_abilitazione_esposizione_siss] == "S":
                #secondo e successivi elementi inseriti nel dict
                if m_d not in metodica_distretti_dict[a_p]:
                    #elemento non costituisce caso 1:N al momento
                    metodica_distretti_dict = self.update_list_in_dict(metodica_distretti_dict, a_p, m_d)
                elif m_d in metodica_distretti_dict[a_p]:
                    #elemento costituisce caso 1:N. 2 o più elementi con stesso m_d
                    metodica_distretti_dict = self.update_list_in_dict(metodica_distretti_dict, a_p, m_d)
                    #error_dict = self.update_list_in_dict(error_dict, str(int(index)+2), a_p)
                else:
                    #teoricamente non dovrebbe mai entrare in questa condizione
                    metodica_distretti_dict = self.update_list_in_dict(metodica_distretti_dict, a_p, m_d)

        print("valutazione risultati casi 1:N")
        for key, value in metodica_distretti_dict.items(): #key:a_p, value:m_d
            print(key + ":  " + ", ".join(value))
            indexAP = index_dict[key] #indici dove si trovano tutte le occorrenze di una coppia A/P
            lengthMD = len(value) #occorrenze coppie prestazioni/agenda

            set_duplicates = {}
            if lengthMD > 1: #occorrenza multipla
                print("occorrenza multipla " + str(lengthMD))
                flag_error1 = False
                flag_error2 = False
                flag_error3 = False
                flag_error4 = False
                error_1_list = [] 
                cont = 0
                #d = "" #value[0].split("_")[1]
                for v in value: #per ogni m_d
                    print("v: " + v)
                    if v in set_duplicates.keys():
                        #set_duplicates.update({v: set_duplicates[v]+1})
                        set_duplicates[v] = set_duplicates[v] + 1
                        print("1:" + str(set_duplicates[v]))
                    else:
                        set_duplicates[v] = 1
                        print("2:" + str(set_duplicates[v]))
                    
                    #set_duplicates[v] = set_duplicates[v].items + 1 #dict con key gli m_d e value sono le occorrenze
                    distrettosplit = v.split("_") #faccio split m_d
                    print("distrettosplit: " + distrettosplit[1])
                    if distrettosplit[1] != []: #controllo se 'd' c'è
                        for vv in distrettosplit[1].split(","):
                            print("d: " + vv)
                            if vv == "":
                                flag_error1 = True
                                error_1_list.append(indexAP[cont])
                        #elif v.split("_")[1] == d:
                        #    flag_error3 = True
                    cont = cont + 1
                
                for k, v in set_duplicates.items(): #verifico le occorrenze
                    if v > 1: #se le occorrenze sono > 1 allora caso 1:N
                        flag_error2 = True

                if flag_error1 == True: #se un distretto è vuoto
                    for ind in error_1_list:
                        print("flag error 1: " + ind)
                        print(error_dict.keys())
                        error_dict["error_casi_1n"].append(ind) #ind ha già sommato + 2
                        casi_1n_dict_error = self.update_list_in_dict(casi_1n_dict_error, ind, key + ": trovato caso 1:N con caso di distretto vuoto")
                if flag_error2 == True: #se una delle m_d è multipla allora è errore
                    for ind in indexAP:
                        print("flag error 2: " + ind)
                        error_dict["error_casi_1n"].append(ind) #ind ha già sommato + 2
                        casi_1n_dict_error = self.update_list_in_dict(casi_1n_dict_error, ind, key + ": " + ", ".join(value))
                
                opd = "A"
                for ind in indexAP: 
                    distretto_op = distretti_dict[ind] #vedo se gli operatori logici sono giusti
                    if opd != distretto_op[0] and df_mapping.at[int(ind)-2,self.work_codice_distretto] != "" and opd != "A":
                        print("dis:" + df_mapping.at[int(ind)-2,self.work_codice_distretto])
                        flag_error4 = True
                        print("flag error 4: " + ind)
                        error_dict["error_casi_1n"].append(ind) #ind ha già sommato + 2
                        casi_1n_dict_error = self.update_list_in_dict(casi_1n_dict_error, ind, key + ": OP non conforme all'interno della coppia agenda/prestazione")
                    opd = distretto_op[0]
                
                '''if flag_error3 == True: #errore se un distretto è uguale a quello precedente
                    for ind in index_dict[key]:
                        error_dict["error_casi_1n"].append(ind) #ind ha già sommato + 2
                        casi_1n_dict_error = self.update_list_in_dict(casi_1n_dict_error, ind, key + ": trovato caso 1:N con caso di distretto vuoto")'''
                


            #if row[self.work_abilitazione_esposizione_siss] == "S":
            '''if row[self.work_casi_1_n] != "OK" or row[self.work_casi_1_n] in "1:N":
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
                    ''''''for md in metodica_distretti_dict[a_p]:
                        if md.split("_")[1] == "":
                            error_dict["error_casi_1n"].append(str(int(index)+1))
                            casi_1n_dict_error = self.update_list_in_dict(casi_1n_dict_error, str(int(index)+2), a_p + ": " + "caso 1:n con distretto vuoto")''''''
                    error_dict["error_casi_1n"].append(str(int(index)+2))
                    casi_1n_dict_error = self.update_list_in_dict(casi_1n_dict_error, str(int(index)+2), a_p + ": " + ", ".join(metodica_distretti_dict[a_p]))
                    #print("trovato caso 1:n per la coppia agenda-prestazione, all'indice: " + str(int(index)+2))
                elif m_d.split("_")[1] == "" and a_p in metodica_distretti_dict.keys() and len(metodica_distretti_dict[a_p]) > 1 and row[self.work_abilitazione_esposizione_siss] == "S":
                    metodica_distretti_dict = self.update_list_in_dict(metodica_distretti_dict, a_p, m_d)
                    error_dict["error_casi_1n"].append(str(int(index)+2))
                    casi_1n_dict_error = self.update_list_in_dict(casi_1n_dict_error, str(int(index)+2), a_p + ": " + "caso 1:n con distretto vuoto")
                elif a_p in metodica_distretti_dict.keys() and m_d not in metodica_distretti_dict[a_p] and row[self.work_abilitazione_esposizione_siss] == "S":
                    ''''''for md in metodica_distretti_dict[a_p]:
                        if md.split("_")[1] == "":
                            error_dict["error_casi_1n"].append(str(int(index)+2))
                            casi_1n_dict_error = self.update_list_in_dict(casi_1n_dict_error, str(int(index)+2), a_p + ": " + "caso 1:n un possibile distretto vuoto in una prestazione")''''''
                    metodica_distretti_dict = self.update_list_in_dict(metodica_distretti_dict, a_p, m_d)
                else:
                    logging.info("trovato caso 1:n con abilitazione SISS a N corretta, all'indice: " + str(int(index)+2))
                    #print("trovato caso 1:n con abilitazione SISS a N corretta, all'indice: " + str(int(index)+2))
                    #print("A_P2: " + a_p + ", Abilititazione Esposizione SISS: " + row["Abilititazione Esposizione SISS"])'''
        
        #modificare controllo andando a verificare a valle la lunghezza del dict metodica_distretti_dict.
        #se per ogni coppia a_p ci sono più di un m_d uguale, allora è errore.
        #se uno di questi m_d ha d vuoto, allora manca d e bisogna segnalare
        print("start definizione output casi 1:n")
        out1 = ""
        out_messsage = ""
        for ind in error_dict['error_casi_1n']:
            out_message = "Casi 1:N: rilevato per la coppia prestazione/agenda: '{}'".format(", ".join(casi_1n_dict_error[ind]))
            out1 = out1 + "at index: " + ind + ", on agenda_prestazione: " + ", ".join(casi_1n_dict_error[ind]) + ", \n"
            if sheet[self.work_alert_column+ind].value is not None:
                sheet[self.work_alert_column+ind] = str(sheet[self.work_alert_column+ind].value) + "; \n" + out_message #modificare colonna alert
            else:
                sheet[self.work_alert_column+ind] = out_message

        self.output_message = self.output_message + "\nerror_casi_1n: \n" + "at index: \n" + out1
            
        xfile.save(self.file_data)  
        return error_dict