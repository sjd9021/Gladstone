import docx
from docx.shared import Inches, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import RGBColor
import streamlit as st
import numpy as np
from docx.shared import Pt
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.table import CT_Row, CT_Tc
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import streamlit.components.v1 as components
import base64

document = docx.Document()
style = document.styles['Normal']
font = style.font
font.name = 'Arial'
font.size = Pt(11)

def download_button(object_to_download, download_filename):
    """
    Generates a link to download the given object_to_download.
    Params:
    ------
    object_to_download:  The object to be downloaded.
    download_filename (str): filename and extension of file. e.g. mydata.docx,
    Returns:
    -------
    (str): the anchor tag to download object_to_download
    """
    try:
        # some strings <-> bytes conversions necessary here
        b64 = base64.b64encode(object_to_download.encode()).decode()

    except AttributeError as e:
        b64 = base64.b64encode(object_to_download).decode()

    dl_link = f"""
    <html>
    <head>
    <title>Start Auto Download file</title>
    <script src="http://code.jquery.com/jquery-3.2.1.min.js"></script>
    <script>
    $('<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{download_filename}">')[0].click()
    </script>
    </head>
    </html>
    """
    return dl_link


def modifyBorder(table):
    tbl = table._tbl  # get xml element in table
    for cell in tbl.iter_tcs():
        tcPr = cell.tcPr  # get tcPr element, in which we can define style of borders
        tcBorders = OxmlElement('w:tcBorders')
        top = OxmlElement('w:top')
        top.set(qn('w:val'), 'nil')

        left = OxmlElement('w:left')
        left.set(qn('w:val'), 'nil')

        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'nil')
        bottom.set(qn('w:sz'), '4')
        bottom.set(qn('w:space'), '0')
        bottom.set(qn('w:color'), 'auto')

        right = OxmlElement('w:right')
        right.set(qn('w:val'), 'nil')

        tcBorders.append(top)
        tcBorders.append(left)
        tcBorders.append(bottom)
        tcBorders.append(right)
        tcPr.append(tcBorders)


def add_imgs(imgs):
    p = document.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run()
    for x in imgs:
        # if len(imgs) == 1 :
        #     r.add_text("    ")
        r.add_picture(imgs[x], width=Cm(7.14), height=Cm(5.24))
        r.add_text("   ")
    last_paragraph = document.paragraphs[-1]
    last_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    return


def add_text(imgs):

    # r = p.add_run()
    leng = len(imgs)
    table = document.add_table(rows=1, cols=leng)
    table.style = 'TableGrid'  # single lines in all cells
    table.autofit = False
    table.allow_autofit = False

    for cell in table.columns[0].cells:
        cell.width = Cm(7.14)
    for row in table.rows:
        row.height = Cm(1)

    # r.font.color.rgb = RGBColor(255, 0, 0)
    c = 0

    for x in imgs:
        p = table.cell(0, c).paragraphs[0]
        p.alignment = docx.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER
        run = table.cell(0, c).paragraphs[0].add_run("   " + x)

        c = c + 1
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    modifyBorder(table)
    p = document.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    return

def image_converter(imgs):

    pairs = list(imgs.items())
    for i in range(0, len(pairs), 2):
        if i + 1 < len(pairs):
            dicter = {pairs[i][0]: pairs[i][1], pairs[i + 1][0]: pairs[i + 1][1]}
            add_imgs(dicter)
            add_text(dicter)

    # If the number of key-value pairs is odd, print the last key-value pair separately
    if len(pairs) % 2 != 0:
        dicter = {pairs[-1][0]:pairs[-1][1]}
        add_imgs(dicter)
        add_text(dicter)

    # print(dicter)

    # document.save('online_demo.docx')
    return document
def download_docx(data):
    edited_doc = image_converter(data)
    buff = io.BytesIO()  # create a buffer
    document.save(buff)  # write the docx to the buffers
    return buff
    # components.html(
    #     download_button(buff.getvalue(), "test.docx"),
    #     height=0,
    # )

def main():
    st.markdown("<h1 style='text-align: center; color: grey;'>Image Parser</h1>", unsafe_allow_html=True)
    c = 0
    col1, col2, col3 = st.columns([1, 3, 1])

    data = {}
    with col1:
        st.write(' ')

    with col2:
        with st.form("Number of Products"):

            numImages = st.number_input('Number Of Images', key='numImages', step=1)
            submitForm = st.form_submit_button("Set Image Number")

        if 'numImages' in st.session_state.keys():
            with st.form("Product Codes"):
                for i in range(int(st.session_state['numImages'])):
                    uploaded_files = st.file_uploader("Image", key=i + 1)
                    damage_description = st.text_input('Description', key=i * 100)
                    for x in data:
                        if damage_description == x:
                            damage_description = damage_description + " "
                    data[damage_description] = uploaded_files

                SubmitForm = st.form_submit_button("Download.docx")
                if SubmitForm:
                    x = download_docx(data)
                    value = x.getvalue()
                    c = 1
            if c==1:
                st.download_button("download docx", value, "test.docx")


    with col3:
        st.write(' ')


if __name__ == "__main__":
    main()
