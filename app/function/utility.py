import os
import shutil
from tkinter import messagebox
from flask import render_template

import pdfkit

def create_new_folder():
    try:
        os.mkdir("./new folder")
        os.mkdir("./new folder/External")
        os.mkdir("./new folder/Photos")
        os.mkdir("./new folder/Plot")
        os.mkdir("./new folder/SS")
        shutil.copyfile("resource/xlsx/Preliminary Calculation v2.5.xlsx",
                        "./new folder/Preliminary Calculation v2.5.xlsx")
    except:
        messagebox.showwarning(title="Error", message="The new folder already exist")


def update_asana():
    pass

def print_pdf():
    # document = ap.Document("resource/pdf/sample_pdf.pdf")
    #
    # textAbsorber = ap.text.TextFragmentAbsorber("Unit 25")
    #
    # document.pages.accept(textAbsorber)
    #
    # textFragmentCollection = textAbsorber.text_fragments
    #
    # for textFragment in textFragmentCollection:
    #     textFragment.text = """
    #     text 1
    #     text 2
    #     """
    #
    # document.save("output.pdf")
    options ={
        "--header-html": "resource/html/header.html",
        "--footer-html": "resource/html/footer.html"
    }
    name="name"
    html = render_template("resource/html/sample_html.html", name=name)
    pdfkit.from_file(html, "output.pdf")
