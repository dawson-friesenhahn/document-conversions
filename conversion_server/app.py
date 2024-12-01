from pptx_to_pdf.pptx_to_pdf import pptx_to_pdf
from pptx.exc import PackageNotFoundError
import os
from glob import glob
from time import sleep

from flask import request, send_from_directory, jsonify

from conversion_server import app


@app.before_request
def before_request():
    files_to_delete = glob(os.path.join(app.config["TEMP_DIR"], "*"))
    for file in files_to_delete:
        try:
            os.remove(file)
        except PermissionError:
            pass


@app.route("/pptx-to-pdf", methods=["POST"])
def convert():
    try:
        file = request.files["file"]
    except KeyError:
        return "No file attached", 400

    extension = file.filename.split(".")[1]

    if extension != "ppt" and extension != "pptx":
        return "Bad file extension", 400

    input_save_location = os.path.join(app.config["TEMP_DIR"], file.filename)

    file.save(input_save_location)
    try:
        output_filename = os.path.basename(
            pptx_to_pdf(input_save_location, app.config["TEMP_DIR"], delete_intermediate_images=True)
        )

        return send_from_directory(app.config["TEMP_DIR"], output_filename)
    except PackageNotFoundError:
        return "Invalid Powerpoint File", 400


@app.route("/")
def home():
    return "Send powerpoint files to this port at the /pptx-to-pdf endpoint"
