import os
import glob
import logging
import json
import yaml
from time import sleep
# third party import
from flask import Flask, flash, request, redirect, url_for, jsonify, send_file, render_template, session
from flask import send_from_directory
from werkzeug.utils import secure_filename
#from flask.ext.session import Session
from flask_session import Session 
# local imports
import flaskr.validator as validator
import flaskr.validator_post_avvio as validator_post_avvio
#from openpyxl import load_workbook
dir = os.path.dirname(__file__)
DOWNLOAD_FOLDER = os.path.join(dir, 'uploads/')

# flask global vars
app = Flask(__name__)
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
app.config['JSONIFY_PRETTYPRINT_REGULAR'] = True
TEMPLATE_MINI_EXAMPLE = 'config_validator_example.yml'
TEMPLATE_WORK_EXAMPLE = 'working_example.json'

# Session
sess = Session()
app.config["SESSION_TYPE"] = "filesystem"
app.config["SECRET_KEY"] = "openspending rocks"
sess.init_app(app)

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
f_handler = logging.FileHandler('validator.log', 'a+', 'utf-8')
c_handler = logging.StreamHandler()
formatter = logging.Formatter(
    '%(asctime)s - %(name)s - %(levelname)s - %(message)s')
f_handler.setFormatter(formatter)
c_handler.setFormatter(formatter)
logger.addHandler(f_handler)
logger.addHandler(c_handler)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    return render_template('index.html')


@app.route('/generate', methods=['GET', 'POST'])
def generate():
    #sess['_flashes'].clear()
    try:
        configuration_file = _read_yml_file1(request, 'yml')
        excel_filename = _read_yml_file2(request, 'file2', 'xlsx')
        _list = request.form.getlist('check')
        checked_dict = _get_generateform_checked_value(_list)
    except ValueError as e:
        logger.error(e)
        return str(e), 400
    try:
        print(excel_filename)
        file_path = app.config['DOWNLOAD_FOLDER']  + excel_filename
        #check_actionPA.initializer(file_path_rivisto, checked_dict)
        check_validation_results = validator.inizializer(checked_dict, configuration_file, file_path)
        #check_action = validator.Check_action(configuration_file, file_path) #_init method
        #check_validation_results = check_action.initializer(checked_dict, configuration_file) #activate method
        if check_validation_results != "":
            flash(str(check_validation_results))
            return redirect(url_for('upload_file'))
            #error_message = check_validation_results
            #return error_message
        return send_file(file_path,as_attachment = True, download_name=excel_filename), 200
    except Exception as e:
        logger.error(e)
        return str(e), 400

@app.route('/validate_agende_interne', methods=['GET', 'POST'])
def validate_agende_interne():
    try:
        configuration_file = _read_yml_file1(request, 'yml')
        excel_filename = _read_yml_file2(request, 'file2', 'xlsx')
    except ValueError as e:
        logger.error(e)
        return str(e), 400

    try:
        print(excel_filename)
        file_path = app.config['DOWNLOAD_FOLDER']  + excel_filename
        #check_action = validator.initializer_check_agende_interne(configuration_file, file_path)
        #check_action.()
                               
        return send_file(file_path,as_attachment = True, download_name=excel_filename), 200
    except Exception as e:
        logger.error(e)
        return str(e), 400
    
@app.route('/send-example', methods=['GET', 'POST'])
def send_example():
    try:
        return send_file(TEMPLATE_MINI_EXAMPLE,as_attachment = True, download_name=TEMPLATE_MINI_EXAMPLE), 200
    except FileNotFoundError as e:
        logger.error(e)
        return str(e), 400

@app.route('/post-avvio-check', methods=['GET', 'POST'])
def post_avvio_service():
    try:
        #flash('check avviato')
        return render_template('post_avvio_page.html'), 200
    except Exception as e:
        logger.error(e)
        return str(e), 400


@app.route('/start-post-avvio-check', methods=['GET', 'POST'])
def post_avvio_start_check():
    try:
        configuration_file = _read_yml_file1(request, 'yml')
        excel1_filename = _read_yml_file2(request, 'fileM', 'xlsx')
        excel2_filename = _read_yml_file2(request, 'fileR', 'xlsx')
        _list = request.form.getlist('check')
        
        checked_dict = _get_form_checked_value(_list, request)
        #return render_template('index.html'), 200
    except FileNotFoundError as e:
        logger.error(e)
        return str(e), 400

    try:
        file_path_mapping = app.config['DOWNLOAD_FOLDER']  + excel1_filename
        file_path_rivisto = app.config['DOWNLOAD_FOLDER']  + excel2_filename
        check_actionPA = validator_post_avvio.Check_action(configuration_file, file_path_mapping)
        check_actionPA.initializer(file_path_rivisto, checked_dict)
        #return "lista: " + ", ".join(_list), 200
        return send_file(file_path_rivisto,as_attachment = True, download_name=excel2_filename), 200
    except Exception as e:
        logger.error(e)
        return str(e), 400

@app.route('/stream')
def stream():
    def read_log():
        with open('generator.log') as f:
            while True:
                yield f.read()
                sleep(1)

    return app.response_class(read_log(), mimetype='text/plain')

def _get_generateform_checked_value(_list):
    
    return_dict = { 
                    "Quesiti": "",
                    "Metodiche": "",
                    "Distretti": "",
                    "Priorita": "",
                    "Prestazione": "",
                    "Canali": "",
                    "Inviante": ""
                }
    for element in _list:
        return_dict[element] = 1
    return return_dict

def _get_form_checked_value(_list, request):
    _list.append("Sheet")
    _list.append("Agenda")
    _list.append("PrestazioneSISS")
    _list.append("PrestazioneInterna")
    return_dict = { 
                    "Sheet": "",
                    "Agenda": "",
                    "PrestazioneSISS": "",
                    "PrestazioneInterna": "",
                    "Quesiti": "",
                    "OperatoreQD": "", 
                    "Distretti": "",
                    "OperatoreDistretto": "",
                    "Metodiche": "",
                    "Inviante": "", 
                    "Risorsa": "",
                    "Farmacia": "",
                    "CCR": "",
                    "Cittadino": "",
                    "MMG": "",
                    "Amministrativo": "",
                    "PAI": "",
                    "NoteOperatore": "",
                    "NotePreparazione": "",
                    "NoteAmministrative": "",
                    "NoteRevoca": "",
                    "PrioritaUrgenza": "",
                    "PrioritaOB": "",
                    "PrioritaOD": "",
                    "PrioritaOP": "",
                    "AccessoProgrammabile": "",
                }
    for element in _list:
        return_dict[element] = request.form[str(element)+'text']
        #print(element + ": " + request.form[str(element)+'text'])
    return return_dict

def _read_yml_file(request, extension):
    if 'file' not in request.files:
        raise ValueError('No file in the request')
    file = request.files['file']
    template_file_dict = json.loads(file.read())
    # if not __is_filename_allowed(file.filename, extension):
    #     raise ValueError(f'File \"{file.filename}\" is not allowed, it must have {extension} extension')
    return template_file_dict

def _read_yml_file1(request, extension):
    if 'file1' not in request.files:
        raise ValueError('No file in the request')
        #raise flash('No file in the request!')
    file = request.files['file1']
    template_file_dict = yaml.load(file, Loader=yaml.FullLoader)
    # if not __is_filename_allowed(file.filename, extension):
    #     raise ValueError(f'File \"{file.filename}\" is not allowed, it must have {extension} extension')
    return template_file_dict

def _read_yml_file2(request, name, extension):
    if request.files[name].filename == '':
        raise ValueError('No file' + str(name) + 'in the request')
        #raise flash('No file in the request!')
    #file = request.files['file2'].read()
    file = request.files[name]
    filename = secure_filename(file.filename)
    file.save(DOWNLOAD_FOLDER + filename)
    return filename