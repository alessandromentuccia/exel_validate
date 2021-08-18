import pandas as pd
import numpy as np
import requests
import xlrd
from xlsxwriter.utility import xl_rowcol_to_cell
import yaml


class check_QD():

    file_name = ""
    output_message = ""
    error_list = {}

    def ck_QD_agenda(self, df_mapping, error_dict):
        return ""

    def ck_QD_disciplina_agenda(self, df_mapping, sheet_QD, error_dict):
        return ""

    def ck_QD_sintassi(self, df_mapping, error_dict):
        return ""

    def ck_QD_descrizione(self, df_mapping, sheet_QD, error_dict):
        return ""

    def ck_QD_operatori_logici(self, df_mapping, error_dict):
        return ""