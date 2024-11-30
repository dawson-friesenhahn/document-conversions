from pptx_to_pdf.pptx_to_pdf import pptx_to_pdf
import os
from glob import glob
from time import sleep

from flask import Flask, request, send_from_directory, after_this_request

from conversion_server import app

@app.before_request
def before_request():
    files_to_delete = glob(os.path.join(app.config['TEMP_DIR'], "*"))
    for file in files_to_delete:
        try: 
            os.remove(file)
        except PermissionError:
            pass


@app.route("/pptx-to-pdf", methods=["POST"])
def convert():

    file= request.files['file']
    input_save_location = os.path.join(app.config['TEMP_DIR'], file.filename)
    
    file.save(input_save_location)

    output_filename= os.path.basename(pptx_to_pdf(input_save_location, app.config['TEMP_DIR']))   
    print("file should be good")
    print(output_filename)
    print(os.getcwd())

    return send_from_directory(app.config['TEMP_DIR'], output_filename)

@app.route("/")
def home():
    return "Hello, world!"
