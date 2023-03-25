from itertools import count
from ntpath import join
from typing import Any
import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell
import logging
import re
import openpyxl 
from pathlib import Path
#from flaskr.validator import Check_action

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

class Report_Creation(): #Check_action

    file_name = ""
    output_message = ""
    error_list = {}
    file_data = {}

    work_N1 = ""
    work_N2 = ""
    work_descrizione_N1 = ""
    work_descrizione_N2 = ""
    work_nota_agenda = ""

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

    #df_mapping = pd.DataFrame
    #xfile = openpyxl.Workbook
    
    #count_ROW = 1
    #count_COLUMN = "A"

    def __init__(self,df_mapping, file_data, work_sheet, work_N1, work_N2, 
                  work_descrizione_N1, work_descrizione_N2, work_codice_prestazione_siss,
                  work_descrizione_prestazione_siss, work_codice_agenda_siss, work_casi_1_n,         
                  work_abilitazione_esposizione_siss, work_prenotabile_siss, work_codici_disciplina_catalogo,         
                  work_descrizione_disciplina_catalogo, work_codice_QD, work_codice_metodica,   
                  work_codice_distretto, work_priorita_U, work_priorita_primo_accesso_D,       
                  work_priorita_primo_accesso_P, work_priorita_primo_accesso_B, work_accesso_programmabile_ZP,   
                  work_combinata, work_codice_agenda_interno, work_codice_prestazione_interno,  
                  work_inviante, work_accesso_farmacia, work_accesso_CCR, 
                  work_accesso_cittadino, work_accesso_MMG, work_accesso_amministrativo, 
                  work_accesso_PAI, work_gg_preparazione, work_gg_refertazione,            
                  work_nota_operatore, work_nota_agenda, work_nota_revoca, work_disciplina,
                  work_sesso, work_eta_min, work_eta_max, work_alert_column, work_delimiter, error_dict):#                 
                         
                                  
        self.df_mapping = df_mapping
        self.file_data = file_data
        self.work_sheet = work_sheet
        self.work_N1 = work_N1
        self.work_N2 = work_N2
        self.work_descrizione_N1 = work_descrizione_N1
        self.work_descrizione_N2 = work_descrizione_N2
        self.work_codice_prestazione_siss = work_codice_prestazione_siss
        self.work_descrizione_prestazione_siss = work_descrizione_prestazione_siss
        self.work_codice_agenda_siss = work_codice_agenda_siss
        self.work_casi_1_n = work_casi_1_n
        self.work_abilitazione_esposizione_siss = work_abilitazione_esposizione_siss
        self.work_prenotabile_siss = work_prenotabile_siss
        self.work_codici_disciplina_catalogo = work_codici_disciplina_catalogo
        self.work_descrizione_disciplina_catalogo = work_descrizione_disciplina_catalogo
        self.work_codice_QD = work_codice_QD
        self.work_codice_metodica = work_codice_metodica
        self.work_codice_distretto = work_codice_distretto
        self.work_priorita_U = work_priorita_U
        self.work_priorita_primo_accesso_D = work_priorita_primo_accesso_D
        self.work_priorita_primo_accesso_P = work_priorita_primo_accesso_P
        self.work_priorita_primo_accesso_B = work_priorita_primo_accesso_B
        self.work_accesso_programmabile_ZP = work_accesso_programmabile_ZP
        self.work_combinata = work_combinata
        self.work_codice_agenda_interno = work_codice_agenda_interno
        self.work_codice_prestazione_interno = work_codice_prestazione_interno
        self.work_inviante = work_inviante
        self.work_accesso_farmacia = work_accesso_farmacia
        self.work_accesso_CCR = work_accesso_CCR
        self.work_accesso_cittadino = work_accesso_cittadino 
        self.work_accesso_MMG = work_accesso_MMG
        self.work_accesso_amministrativo = work_accesso_amministrativo
        self.work_accesso_PAI = work_accesso_PAI
        self.work_gg_preparazione = work_gg_preparazione
        self.work_gg_refertazione = work_gg_refertazione
        self.work_nota_operatore = work_nota_operatore 
        self.work_nota_agenda = work_nota_agenda
        self.work_nota_revoca = work_nota_revoca
        self.work_disciplina = work_disciplina
        self.work_sesso = work_sesso
        self.work_eta_min = work_eta_min
        self.work_eta_max = work_eta_max
        self.work_alert_column = work_alert_column
        self.work_delimiter = work_delimiter
        self.error_dict = error_dict

        self.count_ROW = 1
        self.count_COLUMN = "A"
        self.sheet_report = Any

        #numero_agende = 0
        #super().__init__(self)

    '''def __call__(self, df_mapping, file_data):
        self.get_report(self) 
        self.df_mapping = df_mapping
        self.file_data = file_data'''

    def get_report(self):
        print("start Validation Report")
        
        try:
            #self.df_mapping = pd.read_excel(self.file_data, sheet_name=self.work_sheet, converters={self.work_codici_disciplina_catalogo: str, self.work_codice_prestazione_siss: str, self.work_codice_metodica: str, self.work_codice_distretto: str}).replace(np.nan, '', regex=True)
            print("0:"+ str(self.file_data))
            xfile = openpyxl.load_workbook(self.file_data) #recupero file excel da file system
            print("1")
            #sheet_mapping = xfile.get_sheet_by_name(self.work_sheet) #recupero sheet excel mapping
            try:
                self.sheet_report = xfile.get_sheet_by_name('Report Validazione') #recupero sheet excel report validazione 
            except:
                #creo Report Validazione se non esiste
                print("1.1: Non esiste sheet Validazione - creazione sheet")
                self.sheet_report = xfile.create_sheet('Report Validazione')
                print("1.2")
            print("2")
            #sheet_report["A1"] = "Report Validazione1" #A1
            
            #self.sheet_report[str(self.count_COLUMN)+str(self.count_ROW)] = "Report Validazione" #A1
            print("3.1")
            #print("3.2: "  + str(self.count_COLUMN)+str(self.count_ROW))
            #self.count_ROW+=1
            self.write_row(self.count_COLUMN, "Report Validazione")
            
            print("4")
        except:
            print("non esiste file o sheet")


        self.get_N1_N2()

        self.get_Num_Prestazioni()

        numero_agende = self.get_Num_Agende()

        self.get_Num_PA_Esposte()

        self.get_Num_PA_Prenotabili()

        self.get_Combinate()

        self.get_Raggruppate()

        self.get_Nota_Amministrativa()

        self.get_Nota_Revoca()
        
        self.get_Discipline()

        self.get_Agende_con_QD(numero_agende)

        self.get_Campi_Descrittivi() #Sesso, GG di prep, GG di ref, Età min ed Età max

        self.get_Riassunto_Errori()

        xfile.save(self.file_data) 

        '''
        #for index, row in df_mapping.iterrows():
            
        print("start definizione output agenda QD")
        out1 = ""
        out_message = ""
        for ind in error_dict['error_QD_agenda']:
            out1 = out1 + "at index: , \n"
            out_message = "__> QD: diversi per le prestazioni della stessa agenda"
            if sheet_report[self.work_alert_column+ind].value is not None:
                sheet_report[self.work_alert_column+ind] = str(sheet_report[self.work_alert_column+ind].value) + "; \n" + out_message #modificare colonna alert
            else:
                sheet_report[self.work_alert_column+ind] = out_message

        self.output_message = self.output_message + "\nerror_QD_agenda: \n" + out1
        print("finish print Validation Report")  
        xfile.save(self.file_data) 
        return error_dict'''


    def get_N1_N2(self):
        print("get N1 and N2")

        CODICE_N1 = self.work_N1
        CODICE_N2 = self.work_N2
        DESCRIZIONE_N1 = self.work_descrizione_N1
        DESCRIZIONE_N2 = self.work_descrizione_N2

        cod_N1 = []
        cod_N2 = []
        desc_N1 = []
        desc_N2 = []

        try:
            for index, row in self.df_mapping.iterrows():
                if str(row[CODICE_N1]) not in cod_N1:
                    cod_N1.append(str(row[CODICE_N1]))
                if str(row[CODICE_N2]) not in cod_N2:
                    cod_N2.append(str(row[CODICE_N2]))
                if str(row[DESCRIZIONE_N1]) not in desc_N1:
                    desc_N1.append(str(row[DESCRIZIONE_N1]))
                if str(row[DESCRIZIONE_N2]) not in desc_N2:
                    desc_N2.append(str(row[DESCRIZIONE_N2]))
        except Exception as e:
            logger.error('Failed to get N1 and N2: '+ str(e))
            if cod_N1 == []:
                cod_N1.append("valore assente")
            if cod_N2 == []:
                cod_N2.append("valore assente")
            if desc_N1 == []:
                desc_N1.append("valore assente")
            if desc_N2 == []:
                desc_N2.append("valore assente")

        self.write_row(self.count_COLUMN, "N1: " + ", ".join(cod_N1), 1) #lascio uno spazio dalla riga precedente
        self.write_row(self.count_COLUMN, "Descrizione N1: " + ", ".join(desc_N1))
        self.write_row(self.count_COLUMN, "N2: " + ", ".join(cod_N2))
        self.write_row(self.count_COLUMN, "Descrizione N2: " + ", ".join(desc_N2))
        #self.sheet_report[str(self.count_COLUMN)+str(self.count_ROW)] = "N1: " + ", ".join(N1) 
        

    def get_Num_Prestazioni(self):
        print("get Numero coppie prestazioni/agende")

        contatore = 0

        for index, row in self.df_mapping.iterrows():
            contatore += 1

        self.write_row(self.count_COLUMN, "Numero di coppie prestazione/agenda: " + str(contatore), 1)


    def get_Num_Agende(self):
        print("get Numero agende")

        agende_list = []
        contatore = 0

        for index, row in self.df_mapping.iterrows():
            if row[self.work_codice_agenda_siss] not in agende_list:
                contatore += 1
                agende_list.append(row[self.work_codice_agenda_siss]) 

        self.write_row(self.count_COLUMN, "Numero di agende nell'offerta sanitaria: " + str(contatore), 1)
        return contatore

    def get_Num_PA_Esposte(self):
        print("get Numero coppie PA esposte")

        contatore = 0
        contatore2 = 0
        col_abilitazione = self.work_abilitazione_esposizione_siss

        for index, row in self.df_mapping.iterrows():
            contatore2 += 1
            if row[col_abilitazione] == "S":
                contatore += 1

        self.write_row(self.count_COLUMN, "Numero di coppie prestazione/agenda abilitate all'esposizione: " + str(contatore) + " su " + str(contatore2), 1)

    def get_Num_PA_Prenotabili(self):
        print("get Numero coppie PA esposte e prenotabili")

        contatore = 0
        contatore2 = 0
        col_abilitazione = self.work_abilitazione_esposizione_siss
        col_prenotabile = self.work_prenotabile_siss

        for index, row in self.df_mapping.iterrows():
            contatore2 += 1
            if row[col_abilitazione] == "S" and row[col_prenotabile] == "S":
                contatore += 1

        self.write_row(self.count_COLUMN, "Numero di coppie prestazione/agenda abilitate all'esposizione e prenotabili: " + str(contatore) + " su " + str(contatore2), 0)

    def get_Combinate(self):
        print("get Numero coppie PA combinate")
        
        contatore = 0

        for index, row in self.df_mapping.iterrows():
            if row[self.work_combinata] != "" and row[self.work_combinata] != "N" and row[self.work_abilitazione_esposizione_siss] =="S":
                contatore += 1

        self.write_row(self.count_COLUMN, "Numero di coppie prestazione/agenda in combinata: " + str(contatore), 1)

    def get_Raggruppate(self):
        print("get Numero coppie PA raggruppate")
    '''
        contatore = 0

        for index, row in self.df_mapping.iterrows():
            if row[self.work_ragg] != "":
                contatore += 1

        self.write_row(self.count_COLUMN, "Numero di coppie prestazione/agenda in combinata: " + str(contatore), 1)
    '''

    def get_Nota_Amministrativa(self):
        print("get Nota Amministrativa")
        #occorrenze_nota = 0
        contatore = {}
        note_agenda_list = []

        for index, row in self.df_mapping.iterrows():
            if row[self.work_nota_agenda] not in note_agenda_list:
                contatore[row[self.work_nota_agenda]] = 1
                note_agenda_list.append(row[self.work_nota_agenda]) 
            else:
                occorrenze_nota = contatore.get(row[self.work_nota_agenda])
                contatore[row[self.work_nota_agenda]] = occorrenze_nota + 1
        
        self.write_row(self.count_COLUMN, str(len(contatore)) + " Note amministrative configurate: ", 1)

        for note, cont in contatore.items():
            #if contatore == 1:
            self.write_row("B", " per " + str(cont) + " coppie prestazione/agenda è presente la nota: " + str(note), 0)
            #else:
            #    self.write_row("B", str(note), 0)
            #    contatore -= 1
        

    def get_Nota_Revoca(self):
        print("get Nota Revoca")

        contatore = 0
        note_revoca_list = []

        for index, row in self.df_mapping.iterrows():
            if row[self.work_nota_revoca] not in note_revoca_list:
                contatore += 1
                note_revoca_list.append(row[self.work_nota_revoca]) 
        
        self.write_row(self.count_COLUMN, str(contatore) + " Note revoca configurate: ", 1)

        for note in note_revoca_list:
            #if contatore == 1:
            self.write_row("B", str(note), 0)


    def get_Discipline(self):
        print("get Discipline")

        discipline_list = []

        for index, row in self.df_mapping.iterrows():
            if str(row[self.work_disciplina]) not in discipline_list:
                discipline_list.append(str(row[self.work_disciplina]))

        self.write_row(self.count_COLUMN, "Discipline dell'offerta sanitaria: " + ", ".join(discipline_list), 1)
       
    
    def get_Agende_con_QD(self, num_agende):
        print("get agende con QD")

        list_agende = []
        contatore = 0 #conta le agende con QD

        for index, row in self.df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss]=="S":
                if str(row[self.work_codice_QD]) != "" and str(row[self.work_codice_agenda_siss]) not in list_agende:
                    contatore += 1
                    list_agende.append(str(row[self.work_codice_agenda_siss]))

        contatore_agende_no_QD = num_agende - contatore

        string_OUT = "Sono presenti " + str(contatore) + " AGENDE con QD configurati e " + str(contatore_agende_no_QD) + " AGENDE sprovviste di QD"
        self.write_row(self.count_COLUMN, string_OUT, 1)


    def get_Campi_Descrittivi(self): #Sesso, GG di prep, GG di ref, Età min ed Età max
        print("get Sesso, GG di prep, GG di ref, Età min ed Età max")

        sesso_list = []
        gg_prep_list = []
        gg_ref_list = []
        eta_min_list = []
        eta_max_list = []
        eta_min_max_list = []

        for index, row in self.df_mapping.iterrows():
            if str(row[self.work_sesso]) not in sesso_list:
                sesso_list.append(str(row[self.work_sesso]))
            if str(row[self.work_gg_preparazione]) not in gg_prep_list:
                gg_prep_list.append(str(row[self.work_gg_preparazione]))
            if str(row[self.work_gg_refertazione]) not in gg_ref_list:
                gg_ref_list.append(str(row[self.work_gg_refertazione]))

            eta_min_max = str(row[self.work_eta_min]) + " <->" + str(row[self.work_eta_max])
            if eta_min_max not in eta_min_max_list:
                eta_min_max_list.append(str(eta_min_max))

        self.write_row(self.count_COLUMN, "Sesso nell'offerta sanitaria: " + ", ".join(sesso_list), 1)
        self.write_row(self.count_COLUMN, "Giorni di preparazione definiti nell'offerta sanitaria: " + ", ".join(gg_prep_list))
        self.write_row(self.count_COLUMN, "Giorni di refertazione definiti nell'offerta sanitaria: " + ", ".join(gg_ref_list))
        self.write_row(self.count_COLUMN, "Range di Età minime e massime definite nell'offerta sanitaria: " + ", ".join(eta_min_max_list))
        #self.write_row(self.count_COLUMN, "Età massime definite nell'offerta sanitaria: " + ", ".join(eta_max_list))

    def get_Riassunto_Errori(self):
        print("get Riassunto errori rilevati")
        
        #self.write_row(self.count_COLUMN, "Errori trovati: " + ", ".join(self.error_dict), 1)
        self.write_row(self.count_COLUMN, "Errori trovati: ", 1)
        for key, value in self.error_dict.items():
            print(key, ' : ', value)
            self.write_row(self.count_COLUMN, key + ": ", 0)
            for k, v in value.items():
                self.write_row(self.count_COLUMN, k + ' : ' + ", ".join(v), 0)

    def write_row(self, column, message, row=0):
        self.count_ROW += row #lascio uno spazio dalla riga prima
        self.sheet_report[str(column)+str(self.count_ROW)] = message
        self.count_ROW += 1 #vado a riga successiva per messaggio dopo