import cv2
import numpy as np
import pytesseract
from pytesseract import Output
import pandas as pd
import spacy
import string

import pdf2image
import tabula
from tabulate import tabulate

from pdf2image import convert_from_bytes, convert_from_path
nlp = spacy.load('en_core_web_sm')

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


with open(r'C:\Users\i5\Desktop\AI_Proj\Invoice_pdf_1.PDF', 'rb') as pdf_file:
    images = convert_from_bytes(pdf_file.read(), poppler_path=r"C:\poppler-23.07.0\Library\bin")


ans = []
df = []

df1 = None
table_1_og = None
table_1_1 = None
table_1 = None
table_1_contents = None

for i, page in enumerate(images):
    filename = f"page_no_{i}.jpg"
    page.save(filename, "JPEG")
    image_path = filename
    img = cv2.imread(image_path)

    gray_image = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    ret, thresh1 = cv2.threshold(gray_image, 0, 255, cv2.THRESH_OTSU | cv2.THRESH_BINARY_INV)
    rect_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (18, 14))
    dilation = cv2.dilate(thresh1, rect_kernel, iterations=1)

    contours, hierarchy = cv2.findContours(dilation, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)



    if i == 0:

        table_1_og = {"company_name": [130, 40, 40, 300], "Agence": [410, 40, 170, 350],
                      "Facture no": [453, 40, 275, 250], "date": [497, 30, 130, 200], "DGI.:": [530, 40, 120, 200],
                      "Stockage du mois d'avril 2023 Voir Detail joint": [730, 28, 1400, 110],
                      "304 Manutention entree": [766, 30, 1400, 110], "305 frais etiquetage palette": [806, 30, 1400, 110],
                      "416 Stockage ": [844, 30, 1400, 110], "416 Frais de gestion mensuel ": [884, 30, 1400, 110],
                      "420 Manutention sortie ": [926, 39, 1400, 110], "305 FILMAGE PALETTES": [962, 30, 1400, 110],
                      "420 Etiquetage palette ": [1002, 30, 1400, 110], "420 Etablissement B/L": [1042, 30, 1400, 110],
                      "0 assurance sur valeur déclarée 1 142731.23 euros": [1082, 36, 1400, 110],
                      "Montant taxable :": [1790, 30, 1400, 110], "Montant non taxable : ": [1836, 30, 1400, 110],
                      "T.V.A.20.00% ": [1876, 30, 1400, 110], "Total à payer EUR": [1959, 39, 1400, 110]
                      }

        table_1_og_contents = {'1': "COMPANY NAME", '2': "Agence", '3': "Facture no", '4': "date", '5': "DGI.:",
                               '6': "Stockage du mois d'avril 2023 Voir Detail joint",
                               '7': "304 Manutention entree", '8': "305 frais etiquetage palette", '9': "416 Stockage ",
                               '10': "416 Frais de gestion mensuel ", '11': "420 Manutention sortie ",
                               '12': "305 FILMAGE PALETTES", '13': "420 Etiquetage palette ",
                               '14': "420 Etablissement B/L",'15': "0 assurance sur valeur déclarée 1 142731.23 euros",
                               '16': "Montant taxable :", '17': "Montant non taxable : ",
                               '18': "T.V.A.20.00%",'19': "Total à payer EUR"
                               }

        table_1_og_contents = pd.DataFrame([table_1_og_contents])
        table_1_contents = table_1_og_contents.T

        for j in table_1_og:

            y, h, x, w = table_1_og[j]

            cropped = gray_image[y:y + h, x:x + w]
            im2 = cv2.rectangle(gray_image, (x, y), (x + w, y + h), (0, 0, 0), 1)
            table_1_og[j] = ' '.join(str(nlp(pytesseract.image_to_string(cropped))).split())
        table_1 = pd.DataFrame([table_1_og])
        table_1_1 = table_1.T

    if i == 3:
        dfs = tabula.read_pdf(r'C:\Users\i5\Desktop\AI_Proj\Invoice_pdf_1.PDF', pages='all', encoding='latin-1')
        df1 = pd.DataFrame(dfs[i])


with pd.ExcelWriter('text.xlsx', engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
    table_1_contents.to_excel(writer, startrow=1, header=False, index=False, startcol=0)
    table_1_1.to_excel(writer, startrow=1, header=False, index=False, startcol=1)
    df1.to_excel(writer, startrow=30, header=False, index=False)




