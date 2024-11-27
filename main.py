from pptx_to_pdf.pptx_to_pdf import pptx_to_pdf
import socket
import os
import shutil

from flask import Flask, request, send_from_directory

app = Flask(__name__)

@app.route("/pptx-to-pdf", methods=["POST"])
def convert():
    print(request.data)
    print(request.files)

    file= request.files[0]
    input_save_location = os.path.join("temp_files", request.files[0].filename)
    
    file.save(input_save_location)

    output_filename= os.path.basename(pptx_to_pdf(input_save_location, "temp_files"))

    return send_from_directory("temp_files", output_filename)

@app.route("/")
def home():
    return "Hello, world!"




if __name__ == "__main__":
    if os.path.exists("temp_files"):
        shutil.rmtree("temp_files")
    os.mkdir("temp_files")
    app.run(debug=True)


