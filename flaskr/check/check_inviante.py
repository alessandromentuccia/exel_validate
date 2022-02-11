import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell
import logging
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

class Check_inviante():

    file_name = ""
    output_message = ""
    error_list = {}

    work_sheet = "" #sheet di lavoro di df_mapping
    work_codice_prestazione_siss = ""
    work_descrizione_prestazione_siss = ""
    work_codice_agenda_siss = ""
    work_casi_1_n = ""
    work_abilitazione_esposizione_siss = ""
    work_prenotabile_siss = ""
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
    work_inviante = ""
    work_accesso_farmacia = ""
    work_accesso_CCR = ""
    work_accesso_cittadino = ""
    work_accesso_MMG = ""
    work_accesso_amministrativo = ""
    work_accesso_PAI = ""
    work_gg_preparazione = ""
    work_gg_refertazione = ""

    work_index_codice_QD = 0
    work_index_codice_SISS_agenda = 0
    work_index_abilitazione_esposizione_SISS = 0
    work_index_codice_prestazione_SISS = 0
    work_index_operatore_logico_distretto = 0
    work_index_codici_disciplina_catalogo = 0
    
    def ck_inviante(self, df_mapping, error_dict):
        print("start checking inviante if empty or in error")
        error_dict.update({
            'error_invianti_vuoti': [],
            'error_inviante_controllo': []})

        xfile = openpyxl.load_workbook(self.file_data) #recupero file excel da file system
        sheet = xfile.get_sheet_by_name(self.work_sheet) #recupero sheet excel
        
        str_check = "CONTROLLO"
        for index, row in df_mapping.iterrows():
            if row[self.work_abilitazione_esposizione_siss] == "S": #se prestazione esposta
                if row[self.work_inviante] == "": #controllo se inviante è stato configurato
                    error_dict['error_invianti_vuoti'].append(str(int(index)+2)) 
                elif row[self.work_inviante] == "0" and str_check in row[self.work_codice_prestazione_siss]: #controllo se inviante non è 0 quando prestazione di controllo
                    error_dict['error_inviante_controllo'].append(str(int(index)+2))
                
        out_message = ""
        for ind in error_dict['error_invianti_vuoti']:
            out_message = "__> Rilevato inviante non configurato: inserire valori 0,1,2,3 a seconda delle esigenze"
            if sheet[self.work_alert_column+ind].value is not None:
                sheet[self.work_alert_column+ind] = str(sheet[self.work_alert_column+ind].value) + "; \n" + out_message #modificare colonna alert
            else:
                sheet[self.work_alert_column+ind] = out_message
        for ind in error_dict['error_inviante_controllo']:
            out_message = "__> ALERT: Rilevato inviante a 0 per una visita di controllo. Verificare se è opportuno non avere vincoli sul prescrittore della prestazione di controllo"
            if sheet[self.work_alert_column+ind].value is not None:
                sheet[self.work_alert_column+ind] = str(sheet[self.work_alert_column+ind].value) + "; \n" + out_message #modificare colonna alert
            else:
                sheet[self.work_alert_column+ind] = out_message

        xfile.save(self.file_data)              
        xfile.save(self.file_data)  
        return error_dict