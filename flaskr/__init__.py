import os
import glob
import logging
import json
import yaml
from time import sleep
# third party import
from flask import Flask, flash, request, redirect, url_for, jsonify, send_file, render_template
from flask import send_from_directory
from werkzeug.utils import secure_filename
# local imports
import flaskr.validator
from openpyxl import load_workbook
dir = os.path.dirname(__file__)
DOWNLOAD_FOLDER = os.path.join(dir, 'uploads/')

# flask global vars
app = Flask(__name__)
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
app.config['JSONIFY_PRETTYPRINT_REGULAR'] = True
TEMPLATE_MINI_EXAMPLE = 'config_validator_SUZZARA.yml'
TEMPLATE_WORK_EXAMPLE = 'working_example.json'

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
    try:
        configuration_file = _read_yml_file1(request, 'yml')
        excel_filename = _read_yml_file2(request, 'xlsx')

    except ValueError as e:
        logger.error(e)
        return str(e), 400
    try:
        print(excel_filename)
        file_path = app.config['DOWNLOAD_FOLDER']  + excel_filename
        check_action = validator.Check_action(configuration_file, file_path)
        check_action.initializer()
                               
        return send_file(file_path,as_attachment = True, attachment_filename=excel_filename), 200
    except Exception as e:
        logger.error(e)
        return str(e), 400
    
@app.route('/sendexample', methods=['GET', 'POST'])
def send_example():
    try:
        return send_file(TEMPLATE_MINI_EXAMPLE,as_attachment = True, attachment_filename=TEMPLATE_MINI_EXAMPLE), 200
    except FileNotFoundError as e:
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
    file = request.files['file1']
    template_file_dict = yaml.load(file, Loader=yaml.FullLoader)
    # if not __is_filename_allowed(file.filename, extension):
    #     raise ValueError(f'File \"{file.filename}\" is not allowed, it must have {extension} extension')
    return template_file_dict

def _read_yml_file2(request, extension):
    if 'file2' not in request.files:
        raise ValueError('No file in the request')
    #file = request.files['file2'].read()
    file = request.files['file2']
    filename = secure_filename(file.filename)
    file.save(DOWNLOAD_FOLDER + filename)
    return filename