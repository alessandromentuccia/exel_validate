import os
#import json
#from distutils.log import error
import logging
#from collections import OrderedDict
#from functools import reduce
#from pathlib import Path
#from typing import Dict, List
import pandas as pd
#import numpy as np
#import xlrd
#import openpyxl
#from xlsxwriter.utility import xl_rowcol_to_cell
from dotenv import load_dotenv

from flaskr.check.check_QD import Check_QD
from flaskr.check.check_metodiche import Check_metodiche
from flaskr.check.check_distretti import Check_distretti
from flaskr.check.check_priorita import Check_priorita
from flaskr.check.check_prestazione import Check_prestazione
from flaskr.check.check_canali import Check_canali
from flaskr.check.check_agende_interne import Check_agende_interne
from flaskr.check.check_inviante import Check_inviante
from flaskr.report_creation.rep_creation import Report_Creation #rep_creation.
from flaskr.check_action import Check_action as ca
from flaskr.file import File

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


def inizializer(checked_dict, data, excel_file):

    file = File(data, excel_file, checked_dict)

    check_action = ca(file)#data, excel_file, checked_dict) #_init method
    check_validation_results = check_action.check_sheet(data)
    #print(str(check_validation_results))

    analizer(file, checked_dict, check_action) #activate method

    return check_validation_results


def analizer(file, checked_dict, check_action):#, sheet_QD, sheet_Metodiche, sheet_Distretti):

    print('Start analisys:\n')
    QD_error = {}
    if checked_dict["Quesiti"] == 1:
        print("Fase 1: Controllo Quesiti selezionato") #FASE 1: CONTROLLO I QUESITI DIAGNOSTICI
        QD_error = check_qd(file)#check_action.check_qd()
    metodiche_error = {}
    if checked_dict["Metodiche"] == 1:
        print("Fase 2: Controllo Metodiche selezionato") #FASE 2: CONTROLLO LE METODICHE
        metodiche_error = check_metodiche(file)
    distretti_error = {}
    if checked_dict["Distretti"] == 1:
        print("Fase 3: Controllo Distretti selezionato") #FASE 3: CONTROLLO I DISTRETTI
        distretti_error = check_distretti(file)
    priorita_error = {}
    if checked_dict["Priorita"] == 1:
        print("Fase 4: Controllo Priorita selezionato") #FASE 4: CONTROLLO LE PRIORITA'
        priorita_error = check_priorita(file)
    prestazione_error = {}
    if checked_dict["Prestazione"] == 1:
        print("Fase 5: Controllo Prestazione selezionato") #FASE 5: CONTROLLO UNIVOCITA' PRESTAZIONI e CODICI'
        prestazione_error = check_prestazione(file)
    canali_error = {}
    if checked_dict["Canali"] == 1: #FASE 6: CONTROLLO CANALI DI PRENOTAZIONE
        print("Fase 6: Controllo Canali selezionato")
        canali_error = check_canali(file)
    inviante_error = {}
    if checked_dict["Inviante"] == 1: #FASE 7: CONTROLLO INVIANTE
        print("Fase 7: Controllo Inviante selezionato")
        inviante_error = check_inviante(file)


    error_dict = {
        "QD_error": QD_error,
        "metodiche_error": metodiche_error,
        "distretti_error": distretti_error,
        "priorita_error": priorita_error,    
        "prestazione_error": prestazione_error,
        "canali_error": canali_error,
        "inviante_error": inviante_error
    }

    check_action.file_validation(error_dict)


def check_qd(file):
    print("start checking QD") #Codice Quesito Diagnostico
    #controllo se per ogni Agenda sono inseriti gli stessi QD
    error_dict = {}

    QD_validator = Check_QD(file)
    #error_QD_descrizione = QD_validator.ck_QD_agenda(error_dict)

    error_QD_sintassi = QD_validator.ck_QD_sintassi(error_dict)
    error_QD_agenda = QD_validator.ck_QD_agenda(error_QD_sintassi)
    error_QD_disciplina_agenda = QD_validator.ck_QD_disciplina_agenda(error_QD_agenda)
    error_QD_descrizione = QD_validator.ck_QD_descrizione(error_QD_disciplina_agenda)
    error_QD_operatori_logici = QD_validator.ck_QD_operatori_logici(error_QD_descrizione)

    error_dict = error_QD_operatori_logici

    return error_dict

def check_metodiche(file):
    print("start checking Metodiche")
    error_dict = {}

    Metodiche_validator = Check_metodiche(file)

    file.output_message = file.output_message + "\nErrori presenti nelle metodiche e riportate attraverso gli indici:\n"

    error_metodica_sintassi = Metodiche_validator.ck_metodica_sintassi(error_dict)
    error_metodica_inprestazione = Metodiche_validator.ck_metodica_inprestazione(error_metodica_sintassi)
    #error_metodica_descrizione = Metodiche_validator.ck_metodica_descrizione(sheet_Metodiche, error_metodica_inprestazione)

    error_dict = error_metodica_inprestazione

    #print("error_dict: %s", error_dict)
    return error_dict

def check_distretti(file):
    print("start checking Distretti")
    error_dict = {}

    Distretti_validator = Check_distretti(file)

    error_distretti_sintassi = Distretti_validator.ck_distretti_sintassi(error_dict)
    error_distretti_inprestazione = Distretti_validator.ck_distretti_inprestazione(error_distretti_sintassi)
    #error_distretti_descrizione = Distretti_validator.ck_distretti_descrizione(error_distretti_inprestazione)
    error_distretti_operatori_logici = Distretti_validator.ck_distretti_operatori_logici(error_distretti_inprestazione)

    error_dict = error_distretti_operatori_logici

    return error_dict

def check_priorita(file):
    print("start checking priorità e tipologie di accesso")

    Priorita_validator = Check_priorita(file)

    error_dict = {}

    error_prime_visite = Priorita_validator.ck_prime_visite(error_dict)
    error_controlli = Priorita_validator.ck_controlli(error_prime_visite)
    error_esami_strumentali =Priorita_validator.ck_esami_strumentali(error_controlli)

    error_dict = error_esami_strumentali

    return error_dict

def check_prestazione(file):
    print("start checking univocità delle prestazioni")

    Prestazioni_validator = Check_prestazione(file)

    error_dict = {}

    error_casi_1N = Prestazioni_validator.ck_casi_1n(error_dict)
    error_prestazione = Prestazioni_validator.ck_prestazione(error_casi_1N)
    error_prestazione_non_prenotabile = Prestazioni_validator.ck_prestazione_nonprenotabile(error_prestazione)

    error_dict = error_prestazione_non_prenotabile

    return error_dict

def check_canali(file):
    print("start checking canali di accesso")

    Canali_validator = Check_canali(file)

    error_dict = {}
    error_canali_vuoti = Canali_validator.ck_canali_vuoti(error_dict)
    error_canali_PAI = Canali_validator.ck_canali_PAI(error_canali_vuoti)
    error_canali_abilitati = Canali_validator.ck_canali_abilitati(error_canali_PAI)

    error_dict = error_canali_abilitati

    return error_dict

def check_inviante(file):
    print("start checking inviante")

    Inviante_validator = Check_inviante(file)

    error_dict = {}

    error_inviante = Inviante_validator.ck_inviante(error_dict)

    error_dict = error_inviante

    return error_dict


def initializer_check_agende_interne(configuration_file, file_path):
    
    file = File(configuration_file, file_path, checked_dict= None)
    
    error = analizer_agende_interne(file)
    error_dict = {
        "error_Agende_interne": error
    }
    
    check_action = ca(file)
    check_action.file_validation(error_dict)

def analizer_agende_interne(file):
    Agende_interne_error = Check_agende_interne.ck_agende_interne( self, file.df_mapping, {})
    
    return Agende_interne_error