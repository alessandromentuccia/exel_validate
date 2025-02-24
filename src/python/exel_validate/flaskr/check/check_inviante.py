import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell
import logging
import openpyxl
from flaskr.check_action import Check_action

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

class Check_inviante(Check_action):

    output_message = ""
    error_list = {}

    def __init__(self, file):
        #pass
        super().__init__(file)
        self.file = file
    
    def ck_inviante(self, error_dict):
        print("start checking inviante if empty or in error")
        error_dict.update({
            'error_invianti_vuoti': [],
            'error_inviante_controllo': []})

        xfile = openpyxl.load_workbook(self.file.file_data) #recupero file excel da file system
        sheet = xfile[self.file.work_sheet] #recupero sheet excel
        
        str_check = "CONTROLLO"
        for index, row in self.file.df_mapping.iterrows():
            if row[self.file.work_abilitazione_esposizione_siss] == "S": #se prestazione esposta
                if row[self.file.work_inviante] == "": #controllo se inviante è stato configurato
                    error_dict['error_invianti_vuoti'].append(str(int(index)+2)) 
                    print("error_invianti_vuoti")
                elif row[self.file.work_inviante] == "0" and str_check in row[self.file.work_codice_prestazione_siss]: #controllo se inviante non è 0 quando prestazione di controllo
                    error_dict['error_inviante_controllo'].append(str(int(index)+2))
                    print("error_inviante_controllo")
                
        out_message = ""
        for ind in error_dict['error_invianti_vuoti']:
            out_message = "__> Rilevato inviante non configurato: inserire valori 0,1,2,3 a seconda delle esigenze"
            if sheet[self.file.work_alert_column+ind].value is not None:
                sheet[self.file.work_alert_column+ind] = str(sheet[self.file.work_alert_column+ind].value) + "; \n" + out_message #modificare colonna alert
            else:
                sheet[self.file.work_alert_column+ind] = out_message
        for ind in error_dict['error_inviante_controllo']:
            out_message = "__> ALERT: Rilevato inviante a 0 per una visita di controllo. Verificare se è opportuno non avere vincoli sul prescrittore della prestazione di controllo"
            if sheet[self.file.work_alert_column+ind].value is not None:
                sheet[self.file.work_alert_column+ind] = str(sheet[self.file.work_alert_column+ind].value) + "; \n" + out_message #modificare colonna alert
            else:
                sheet[self.file.work_alert_column+ind] = out_message

        xfile.save(self.file.file_data)
        xfile.save(self.file.file_data)
        return error_dict