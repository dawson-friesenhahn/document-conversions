import os
import shutil
from flask import Flask

app = Flask(__name__)
app.config["TEMP_DIR"] = os.path.abspath(".temp_files")

if os.path.exists(app.config["TEMP_DIR"]):
    shutil.rmtree(app.config["TEMP_DIR"])

os.mkdir(app.config["TEMP_DIR"])
