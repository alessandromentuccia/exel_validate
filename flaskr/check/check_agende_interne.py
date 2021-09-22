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

class Check_agende_interne():

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
    
    def ck_agende_interne(self, df_mapping, error_dict):
        print("start checking agende interne")
        error_dict.update({'error_agende_interne': []})

        xfile = openpyxl.load_workbook(self.file_data) #recupero file excel da file system
        sheet = xfile.get_sheet_by_name(self.work_sheet) #recupero sheet excel

        agende_interne_dict = {}
        agende_siss_dict = {}
        index_dict = {}
        
        for index, row in df_mapping.iterrows():
            #print("iterate excel row")
            codice_agenda_interno = row[self.work_codice_agenda_interno]
            codice_agenda_siss = row[self.work_codice_agenda_siss]
            codice_prestazione_siss = row[self.work_codice_prestazione_siss]
            agende_interne_dict = self.update_list_in_dict(agende_interne_dict, codice_agenda_interno, str(int(index)+2) + "|" + codice_agenda_siss)
            agende_siss_dict = self.update_list_in_dict(agende_siss_dict, codice_agenda_siss, str(int(index)+2) + "|" + codice_prestazione_siss)
            index_dict = self.update_list_in_dict(index_dict, codice_agenda_siss + "|" + codice_prestazione_siss, str(int(index)+2))
        
        for key, value in agende_interne_dict.items(): #key:agende_interne, value:indici
            #print("agenda interna: ", value)
            prestazioni_list = []
            prestazione_precedenti = []
            prestazione_precedenti_precedenti = []
            #flag_error_detected = False
            cont = 0
            if len(value) > 1:
                for v in value: #ciclo per ogni record con stesso codice siss agenda
                    #print("agenda SISS: ", v)
                    v = v.split("|")
                    indice_agenda_SISS = v[0]
                    codice_agenda_SISS = v[1] 
                    prestazioni_attuali = agende_siss_dict[codice_agenda_SISS]
                    p_list = []
                    i_list = []
                    for l in prestazioni_attuali:
                        l = l.split("|")
                        prest = l[1]
                        inde = l[0]
                        p_list.append(prest)
                        i_list.append(inde)

                    if cont > 0 and len(p_list) > 1:
                        if(set(p_list) != set(prestazione_precedenti)):
                            print("Lists are not equal")
                            list_intersection = set(p_list) & set(prestazione_precedenti) #lista con elementi in comune
                            if list_intersection != {}:

                                ind = 0
                                for p in p_list:
                                    if p not in list_intersection:
                                        indice = i_list[ind]
                                        if indice not in error_dict['error_agende_interne']:
                                            error_dict['error_agende_interne'].append(indice)
                                    ind = ind + 1
                        else:
                            print("Lists are equal")

                    prestazione_precedenti = p_list
                    if cont%2 == 0:
                        prestazione_precedenti_precedenti = p_list
                    cont = cont + 1



                    #prestazione_siss = df_mapping.at[int(indice_agenda_SISS), self.work_codice_prestazione_siss]
                    #prestazioni_list = [] #.append(prestazione_siss) 
                    '''lista_appoggio = []
                    prestazioni_attuali = agende_siss_dict[codice_agenda_SISS]
                    #print("prestazioni_attuali", prestazioni_attuali)
                    
                    if cont != 0 and len(prestazioni_attuali) > 1:
                        for dizionario in prestazioni_attuali:
                            dizionario = dizionario.split("|")
                            prestazione = dizionario[1] #prestazione SISS
                            indice_prestazione = dizionario[0] #indice prestazione SISS
                            #print("indice_prestazione: " + str(indice_prestazione) + "prestazione: " + prestazione)
                            #print("prestazioni_list", prestazioni_list)
                            if prestazione not in prestazioni_list:
                                #flag_error_detected = True
                                ind = indice_prestazione #index_dict[prestazione]
                                if ind not in error_dict['error_agende_interne']:
                                    error_dict['error_agende_interne'].append(ind)
                                    print("ERRORE TROVATO")
                            lista_appoggio.append(prestazione)
                            #prestazione__past = prestazione
                            #cont = cont + 1
                    elif cont == 0:
                        for dizionario in prestazioni_attuali:
                            dizionario = dizionario.split("|")
                            prestazione = dizionario[1]
                            indice_prestazione = dizionario[0]
                            lista_appoggio.append(prestazione)
                            #prestazione__past = prestazione
                    prestazioni_list = lista_appoggio
                    cont = cont + 1'''
                    


        #creazione del messaggio di alert riportato nel file excel
        print("start definizione output agende interne")
        out1 = ""
        out_message = ""
        for ind in error_dict['error_agende_interne']:
            out_message = "__> Codice Interno Agenda:\n"
            out_message = out_message + "  _> rilevata prestazione non conforme per l'agenda interna"
            #out1 = out1 + "at index: " + ind + ", on agenda: " + ", ".join(agenda_interna_error[ind]) + ", \n"
            if sheet[self.work_alert_column+ind].value is not None:
                sheet[self.work_alert_column+ind] = str(sheet[self.work_alert_column+ind].value) + "; \n" + out_message #modificare colonna alert
            else:
                sheet[self.work_alert_column+ind] = out_message

        self.output_message = self.output_message + "\nerror_agenda_interna: \n" + "at index: \n" + out1
            
        xfile.save(self.file_data)  
        return error_dict

    