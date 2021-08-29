import os
import glob
import logging
import json
from time import sleep
# third party import
from flask import Flask, flash, request, redirect, url_for, jsonify, send_file, render_template
from flask import send_from_directory
# local imports
import flaskr.validator
# flask global vars
app = Flask(__name__)
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
        template_file_dict = request
    except ValueError as e:
        logger.error(e)
        return str(e), 400
    try:
        generated_examples_path = generator_core.generate_examples(template_file_dict, 100)
        return send_file(generated_examples_path,as_attachment = True, attachment_filename=generated_examples_path), 200
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

@app.route('/config', methods=['GET', 'POST'])
def config():
    try:
        template_file_dict = request
    except ValueError as e:
        logger.error(e)
        return str(e), 400
    try:
        check_action = validator.Check_action()
        generated_validation = check_action.config(template_file_dict)
        return send_file(generated_validation, as_attachment = True, attachment_filename=generated_validation), 200
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

    


def _read_yml_file(request, extension):
    if 'file' not in request.files:
        raise ValueError('No file in the request')
    file = request.files['file']
    template_file_dict = json.loads(file.read())
    # if not __is_filename_allowed(file.filename, extension):
    #     raise ValueError(f'File \"{file.filename}\" is not allowed, it must have {extension} extension')
    return template_file_dict
