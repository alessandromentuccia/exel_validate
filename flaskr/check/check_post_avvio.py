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

class Check_post_avvio():
    
    def ck_post_avvio(self, df_mapping, df_rivisto, error_dict):
        print("start checking post avvio")
        error_dict.update({'error_post_avvio': []})

        xfile = openpyxl.load_workbook(self.file_data) #recupero file excel da file system
        sheet = xfile.get_sheet_by_name(self.work_sheet) #recupero sheet excel

        if self.configurazione_rivisto["Quesiti"] != "":
            error_dict = Check_post_avvio.ck_QD(self, df_mapping, df_rivisto, error_dict)
        '''if self.configurazione_rivisto["OperatoreQD"] != "":
            error_dict = Check_post_avvio.ck_Operatore_QD(self, df_mapping, df_rivisto, error_dict)
        if self.configurazione_rivisto["Distretti"] != "":
            error_dict = Check_post_avvio.ck_Distretti(self, df_mapping, df_rivisto, error_dict)
        if self.configurazione_rivisto["OperatoreDistretto"] != "":
            error_dict = Check_post_avvio.ck_Operatore_Distretti(self, df_mapping, df_rivisto, error_dict)
        if self.configurazione_rivisto["Metodiche"] != "":
            error_dict = Check_post_avvio.ck_Metodiche(self, df_mapping, df_rivisto, error_dict)    
        if self.configurazione_rivisto["Inviante"] != "":
            error_dict = Check_post_avvio.ck_Invianti(self, df_mapping, df_rivisto, error_dict)
        if self.configurazione_rivisto["Risorsa"] != "":
            error_dict = Check_post_avvio.ck_Risorsa(self, df_mapping, df_rivisto, error_dict)
        if self.configurazione_rivisto["Canaliabilitati"] != "":
            error_dict = Check_post_avvio.ck_Canali_abilitati(self, df_mapping, df_rivisto, error_dict)
        #for index, row in df_mapping.iterrows():'''

        #creazione del messaggio di alert riportato nel file excel
        print("start definizione output controlli post avvio")
        '''out1 = ""
        out_message = ""
        for ind in error_dict['error_post_avvio']:
            out_message = "Casi 1:N: rilevato per la coppia prestazione/agenda: "
            out1 = out1 + "at index: " + ind + ", on agenda_prestazione: "
            if sheet[self.work_alert_column+ind].value is not None:
                sheet[self.work_alert_column+ind] = str(sheet[self.work_alert_column+ind].value) + "; \n" + out_message #modificare colonna alert
            else:
                sheet[self.work_alert_column+ind] = out_message

        self.output_message = self.output_message + "\nerror_post_avvio: \n" + "at index: \n" + out1
        
        xfile.save(self.file_data) ''' 
        return error_dict

    def ck_QD(self, df_mapping, df_rivisto, error_dict):
        error_dict.update({
            'error_ck_QD': [] })
        
        xfile = openpyxl.load_workbook(self.file_data) #recupero file excel da file system
        sheet = xfile.get_sheet_by_name(self.work_sheet)

        file_rivisto = self.file_rivisto 
        configurazione_rivisto = self.configurazione_rivisto

        for index, row in df_rivisto.iterrows():
            rivisto_key = str(row["CD_AGENDA"].strip())+"|"+str(row["CD_PRESTAZIONE_SISS"].strip())+"|"+str(row["CD_INTERNO_PRESTAZIONE"].strip())
            #print("iterate row: " + rivisto_key)
            #print("iterate row: " + row[self.configurazione_rivisto["Quesiti"]])
            searchedValue = row[self.configurazione_rivisto["Quesiti"]] #modificare
            result_QD = self.findCell_dataframe(df_mapping, searchedValue, rivisto_key, self.work_codice_QD) #modificare
            if result_QD == -1:
                error_dict["error_ck_QD"].append(str(int(index)+2))
                print("trovato errore su Quesiti")
                
        xfile.save(self.file_data)     
        return error_dict

    def ck_Operatore_QD(self, df_mapping, df_rivisto, error_dict):
        error_dict.update({
            'error_ck_Operatore_QD': [] })
        
        xfile = openpyxl.load_workbook(self.file_data) #recupero file excel da file system
        sheet = xfile.get_sheet_by_name(self.work_sheet)

        file_rivisto = self.file_rivisto 
        configurazione_rivisto = self.configurazione_rivisto

        for index, row in df_rivisto.iterrows():
            rivisto_key = str(row["CD_AGENDA"])+"|"+str(row["CD_PRESTAZIONE_SISS"])+"|"+str(row["CD_INTERNO_PRESTAZIONE"])
            #print("iterate row: " + rivisto_key)
            #print("iterate row: " + row[self.configurazione_rivisto["Quesiti"]])
            searchedValue = row[self.configurazione_rivisto["Quesiti"]] #modificare
            result_QD = self.findCell_dataframe(df_mapping, searchedValue, rivisto_key, self.work_codice_QD) #modificare
            if result_QD == -1:
                error_dict["error_ck_QD"].append(str(int(index)+2))

        xfile.save(self.file_data)     
        return error_dict

    def ck_Distretti(self, df_mapping, df_rivisto, error_dict):
        error_dict.update({
            'error_ck_Distretti': [] })
        
        xfile = openpyxl.load_workbook(self.file_data) #recupero file excel da file system
        sheet = xfile.get_sheet_by_name(self.work_sheet)

        file_rivisto = self.file_rivisto 
        configurazione_rivisto = self.configurazione_rivisto

        for index, row in df_rivisto.iterrows():
            rivisto_key = str(row["CD_AGENDA"])+"|"+str(row["CD_PRESTAZIONE_SISS"])+"|"+str(row["CD_INTERNO_PRESTAZIONE"])
            #print("iterate row: " + rivisto_key)
            #print("iterate row: " + row[self.configurazione_rivisto["Quesiti"]])
            searchedValue = row[self.configurazione_rivisto["Quesiti"]] #modificare
            result_QD = self.findCell_dataframe(df_mapping, searchedValue, rivisto_key, self.work_codice_QD) #modificare
            if result_QD == -1:
                error_dict["error_ck_QD"].append(str(int(index)+2))

        xfile.save(self.file_data)     
        return error_dict

    def ck_Operatore_Distretti(self, df_mapping, df_rivisto, error_dict):
        error_dict.update({
            'error_ck_Operatore_Distretti': [] })
        
        xfile = openpyxl.load_workbook(self.file_data) #recupero file excel da file system
        sheet = xfile.get_sheet_by_name(self.work_sheet)

        file_rivisto = self.file_rivisto 
        configurazione_rivisto = self.configurazione_rivisto

        for index, row in df_rivisto.iterrows():
            rivisto_key = str(row["CD_AGENDA"])+"|"+str(row["CD_PRESTAZIONE_SISS"])+"|"+str(row["CD_INTERNO_PRESTAZIONE"])
            #print("iterate row: " + rivisto_key)
            #print("iterate row: " + row[self.configurazione_rivisto["Quesiti"]])
            searchedValue = row[self.configurazione_rivisto["Quesiti"]] #modificare
            result_QD = self.findCell_dataframe(df_mapping, searchedValue, rivisto_key, self.work_codice_QD) #modificare
            if result_QD == -1:
                error_dict["error_ck_QD"].append(str(int(index)+2))

        xfile.save(self.file_data)     
        return error_dict

    def ck_Metodiche(self, df_mapping, df_rivisto, error_dict):
        error_dict.update({
            'error_ck_Metodiche': [] })
        
        xfile = openpyxl.load_workbook(self.file_data) #recupero file excel da file system
        sheet = xfile.get_sheet_by_name(self.work_sheet)

        file_rivisto = self.file_rivisto 
        configurazione_rivisto = self.configurazione_rivisto

        for index, row in df_rivisto.iterrows():
            rivisto_key = str(row["CD_AGENDA"])+"|"+str(row["CD_PRESTAZIONE_SISS"])+"|"+str(row["CD_INTERNO_PRESTAZIONE"])
            #print("iterate row: " + rivisto_key)
            #print("iterate row: " + row[self.configurazione_rivisto["Quesiti"]])
            searchedValue = row[self.configurazione_rivisto["Quesiti"]] #modificare
            result_QD = self.findCell_dataframe(df_mapping, searchedValue, rivisto_key, self.work_codice_QD) #modificare
            if result_QD == -1:
                error_dict["error_ck_QD"].append(str(int(index)+2))

        xfile.save(self.file_data)     
        return error_dict
    
    def ck_Invianti(self, df_mapping, df_rivisto, error_dict):
        error_dict.update({
            'error_ck_Invianti': [] })
        
        xfile = openpyxl.load_workbook(self.file_data) #recupero file excel da file system
        sheet = xfile.get_sheet_by_name(self.work_sheet)

        file_rivisto = self.file_rivisto 
        configurazione_rivisto = self.configurazione_rivisto

        for index, row in df_rivisto.iterrows():
            rivisto_key = str(row["CD_AGENDA"])+"|"+str(row["CD_PRESTAZIONE_SISS"])+"|"+str(row["CD_INTERNO_PRESTAZIONE"])
            #print("iterate row: " + rivisto_key)
            #print("iterate row: " + row[self.configurazione_rivisto["Quesiti"]])
            searchedValue = row[self.configurazione_rivisto["Quesiti"]] #modificare
            result_QD = self.findCell_dataframe(df_mapping, searchedValue, rivisto_key, self.work_codice_QD) #modificare
            if result_QD == -1:
                error_dict["error_ck_QD"].append(str(int(index)+2))

        xfile.save(self.file_data)     
        return error_dict

    def ck_Risorsa(self, df_mapping, df_rivisto, error_dict):
        error_dict.update({
            'error_ck_Risorsa': [] })
        
        xfile = openpyxl.load_workbook(self.file_data) #recupero file excel da file system
        sheet = xfile.get_sheet_by_name(self.work_sheet)

        file_rivisto = self.file_rivisto 
        configurazione_rivisto = self.configurazione_rivisto

        for index, row in df_rivisto.iterrows():
            rivisto_key = str(row["CD_AGENDA"])+"|"+str(row["CD_PRESTAZIONE_SISS"])+"|"+str(row["CD_INTERNO_PRESTAZIONE"])
            #print("iterate row: " + rivisto_key)
            #print("iterate row: " + row[self.configurazione_rivisto["Quesiti"]])
            searchedValue = row[self.configurazione_rivisto["Quesiti"]] #modificare
            result_QD = self.findCell_dataframe(df_mapping, searchedValue, rivisto_key, self.work_codice_QD) #modificare
            if result_QD == -1:
                error_dict["error_ck_QD"].append(str(int(index)+2))

        xfile.save(self.file_data)     
        return error_dict

    def ck_Canali_abilitati(self, df_mapping, df_rivisto, error_dict):
        error_dict.update({
            'error_ck_Canali_abilitati': [] })
        
        xfile = openpyxl.load_workbook(self.file_data) #recupero file excel da file system
        sheet = xfile.get_sheet_by_name(self.work_sheet)

        file_rivisto = self.file_rivisto 
        configurazione_rivisto = self.configurazione_rivisto

        for index, row in df_rivisto.iterrows():
            rivisto_key = str(row["CD_AGENDA"])+"|"+str(row["CD_PRESTAZIONE_SISS"])+"|"+str(row["CD_INTERNO_PRESTAZIONE"])
            #print("iterate row: " + rivisto_key)
            #print("iterate row: " + row[self.configurazione_rivisto["Quesiti"]])
            searchedValue = row[self.configurazione_rivisto["Quesiti"]] #modificare
            result_QD = self.findCell_dataframe(df_mapping, searchedValue, rivisto_key, self.work_codice_QD) #modificare
            if result_QD == -1:
                error_dict["error_ck_QD"].append(str(int(index)+2))

        xfile.save(self.file_data)     
        return error_dict
    