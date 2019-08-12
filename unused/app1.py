from ocr import run_ocr
from excel import excel_coor
import pandas as pd

#data, coor, conf = run_ocr('Obligor_5_Facility_Letter_Dah_Sing.pdf', 'LR').multi_page_reader()
data, coor, conf = run_ocr('Obligor_3_DBS_Bills_Statements_Apr.pdf', 'BiS').multi_page_reader()
data.to_excel("static/temp/output_info.xlsx")
coor.to_excel("static/temp/cor_info.xlsx")
excel_coor()
