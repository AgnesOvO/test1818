from app import app
from flask import render_template, request, redirect
import os
from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.styles import Font
import pandas as pd

@app.route("/")
def index():
    return render_template("public/index.html")

@app.route("/upload-xlsx", methods=["GET", "POST"])
def upload_xlsx():

    if request.method == "POST":

        if request.files:

            xlsx = request.files["xlsx"]



            print(xlsx)

            return redirect(request.url)

    return render_template("public/upload_xlsx.html")