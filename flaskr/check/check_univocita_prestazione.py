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
    
    def ck_casi_1n(self, df_mapping, error_dict):
        print("start checking if casi 1:n is correct")
        error_dict.update({'error_casi_1n': []})

        xfile = openpyxl.load_workbook(self.file_data) #recupero file excel da file system
        sheet = xfile.get_sheet_by_name(self.work_sheet) #recupero sheet excel

        casi_1n_dict_error = {}
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

            set_duplicates = {} #dict con key gli m_d e value sono le occorrenze
            if lengthMD > 1: #occorrenza multipla, rilevo possibile caso 1:N se lenght > 1
                print("occorrenza multipla " + str(lengthMD))
                flag_error1 = False
                flag_error2 = False
                flag_error3 = False
                error_1_list = [] 
                error_2_list = [] 
                cont = 0
                dict_items_cont = {} #inserisco chiave:m_d e value: a che indice di indexMD è
                
                for v in value: #per ogni m_d in value
                    print("v: " + v)
                    #definisco un dict per conteggiare le occorrenze degli m_d
                    if v in set_duplicates.keys(): 
                        set_duplicates[v] = set_duplicates[v] + 1
                        dict_items_cont[v].append(cont)
                        #print("1:" + str(set_duplicates[v]))
                    else:
                        set_duplicates[v] = 1
                        dict_items_cont[v] = [cont]
                        #print("2:" + str(set_duplicates[v]))
                    
                    splitMD_list = v.split("_") #splitto m_d in una lista di due elementi
                    print("splitMD_list: " + splitMD_list[1])
                    #verifico se nei casi 1:N, i distretto non sono vuoti
                    if splitMD_list[1] != []: 
                        for distretto in splitMD_list[1].split(","):
                            print("d: " + distretto)
                            if distretto == "":
                                flag_error1 = True
                                error_1_list.append(indexAP[cont])
                    cont = cont + 1
                
                #verifico le occorrenze degli m_d
                #se le occorrenze di un m_d sono > 1, allora caso 1:N non risolto
                for MD, occorrenze in set_duplicates.items(): 
                    if occorrenze > 1: #se le occorrenze sono > 1 allora caso 1:N
                        flag_error2 = True
                        for md in dict_items_cont[MD]:
                            error_2_list.append(indexAP[md])
                
                #se flag_error1 è True, allora trovato errore distretto vuoto
                if flag_error1 == True: 
                    for ind in error_1_list:
                        print("flag error 1: " + ind)
                        casi_1n_dict_error = self.update_list_in_dict(casi_1n_dict_error, ind, key + ": caso 1:N con distretto vuoto")
                        if ind not in error_dict["error_casi_1n"]:
                            error_dict["error_casi_1n"].append(ind) #ind ha già sommato + 2
                #se flag_error2 è True, allora trovato errore chiave m_d multipla
                if flag_error2 == True: 
                    for ind in error_2_list:
                        print("flag error 2: " + ind)
                        casi_1n_dict_error = self.update_list_in_dict(casi_1n_dict_error, ind, key + " con metodica_distretto: " + ", ".join(value))
                        if ind not in error_dict["error_casi_1n"]:
                            error_dict["error_casi_1n"].append(ind) #ind ha già sommato + 2
                         
                #verifico se gli operatori logici sono conformi e univoci per ogni coppia A/P nei casi 1:N
                opd = "A"
                for ind in indexAP: 
                    distretto_op = distretti_dict[ind] 
                    if opd != distretto_op[0] and df_mapping.at[int(ind)-2,self.work_codice_distretto] != "" and opd != "A":
                        print("dis:" + df_mapping.at[int(ind)-2,self.work_codice_distretto])
                        flag_error3 = True
                        print("flag error 4: " + ind)
                        error_dict["error_casi_1n"].append(ind) #ind ha già sommato + 2
                        casi_1n_dict_error = self.update_list_in_dict(casi_1n_dict_error, ind, key + ": OP non conforme all'interno della coppia agenda/prestazione")
                    opd = distretto_op[0]
    
                #taggo i casi 1:N risolti, cioè che non presentano errori
                for i in indexAP: 
                    if i not in error_dict["error_casi_1n"]:
                        out_message = "caso 1:N risolto"
                        if sheet[self.work_alert_column+i].value is not None:
                            sheet[self.work_alert_column+i] = str(sheet[self.work_alert_column+i].value) + "; \n" + out_message #modificare colonna alert
                        else:
                            sheet[self.work_alert_column+i] = out_message

        #creazione del messaggio di alert riportato nel file excel
        print("start definizione output casi 1:n")
        out1 = ""
        out_message = ""
        for ind in error_dict['error_casi_1n']:
            out_message = "Casi 1:N: rilevato per la coppia prestazione/agenda: '{}'".format(", ".join(casi_1n_dict_error[ind]))
            if "S" in df_mapping.at[int(ind)-2, self.work_combinata]: #verifico se c'è combinata
                out_message = out_message + "; \nCasi 1:N: rilevata possibile risoluzione tramite combinata"
            out1 = out1 + "at index: " + ind + ", on agenda_prestazione: " + ", ".join(casi_1n_dict_error[ind]) + ", \n"
            if sheet[self.work_alert_column+ind].value is not None:
                sheet[self.work_alert_column+ind] = str(sheet[self.work_alert_column+ind].value) + "; \n" + out_message #modificare colonna alert
            else:
                sheet[self.work_alert_column+ind] = out_message

        self.output_message = self.output_message + "\nerror_casi_1n: \n" + "at index: \n" + out1
            
        xfile.save(self.file_data)  
        return error_dict