from flask import Flask, redirect, render_template, request, session, url_for, send_from_directory
from flask_dropzone import Dropzone
from flask_uploads import UploadSet, configure_uploads, IMAGES, patch_request_class
from werkzeug.utils import secure_filename
import time
import OCR_Citi_BS as obs
import excel
import threading
import pythoncom
import win32com.client as win32
import os
import OCR_DBS_billstatement as ob
import split_loan as ls



app = Flask(__name__)
dropzone = Dropzone(app)


app.config['SECRET_KEY'] = 'supersecretkeygoeshere'

# Dropzone settings
app.config['DROPZONE_UPLOAD_MULTIPLE'] = True
app.config['DROPZONE_ALLOWED_FILE_CUSTOM'] = True
app.config['DROPZONE_ALLOWED_FILE_TYPE'] = 'image/*, .pdf, .xlsm, .xlsb'
# app.config['DROPZONE_REDIRECT_VIEW'] = 'results'
# Uploads settings
#app.config['UPLOADED_PHOTOS_DEST'] = os.getcwd() + '/uploads'

#photos = UploadSet('photos', IMAGES)
#configure_uploads(app, photos)
#patch_request_class(app)  # set maximum file size, default is 16MB
files = []
APP_ROOT = os.getcwd()
output_loc = APP_ROOT + "\\static\\output"
try:
    os.remove(os.path.join(output_loc, 'output.xlsm'))
    xl = win32.gencache.EnsureDispatch('Excel.Application')
    xl.Visible = False  # change to False after development
    xl.DisplayAlerts = False
    xl.Quit()
except:
    pass


@app.route('/', methods=['GET', 'POST'])
def index():
    # set session for image results
    if "file_urls" not in session:
        session['file_urls'] = []
    # list to hold our uploaded image urls
    file_urls = session['file_urls']

    # handle image upload from Dropszone
    if request.method == 'POST':
        save_path = os.getcwd() + '\\static\\uploads'
        file_obj = request.files
        try:
            filename = request.json["name"]
            print('Delete filename:' + filename)
            file_path = os.path.join(save_path, filename)
            os.remove(file_path)
            files.remove(file_path)
        except Exception:
            pass
        for f in file_obj:
            file = request.files.get(f)
            filename = file.filename
            print('Upload filename:' + filename)
            file_path = os.path.join(save_path, filename)
            file.save(file_path)

            # append image urls
            file_urls.append(file_path)
            files.append(file_path)
            
        session['file_urls'] = file_urls
        print(files)
        return render_template('loading_1.html')
    # return dropzone template on GET request    
    return render_template('index.html')


@app.route('/load')
def load():
    
    # redirect to home if no images to display
    # if "file_urls" not in session or session['file_urls'] == []:
    #     return redirect(url_for('index'))
    #
    # # set the file_urls and remove the session variable
    file_urls = session['file_urls']
    # session.pop('file_urls', None)
    return render_template('loading_1.html', file_urls=file_urls)

@app.route('/download', methods=["GET", "POST"])
def downloads():
    save_path = os.getcwd() + '\\static\\output'
    return send_from_directory(save_path, 'output.xlsm', as_attachment=True)
    # try:
    #     return send_from_directory(files[0], 'result.pdf', as_attachment=True)
    # except Exception:
    #     return render_template('loading_1.html')

@app.route('/process')
def process():
    # redirect to home if no images to display
    # if "file_urls" not in session or session['file_urls'] == []:
    #     return redirect(url_for('index'))
    #
    time.sleep(1)
    temp_loc = "/static/temp"
    upload_loc = "/static/uploads"

    coor_file = os.path.join(temp_loc, 'cor_info.xlsx')
    data_file = os.path.join(temp_loc, 'output_info.xlsx')
    loan_file = os.path.join(temp_loc, 'loanstatement.xlsx')
    bill_file = os.path.join(temp_loc, 'billstatement.xlsx')
    print(files)
    for i in files:
        if 'bank' in i.lower():
            print('bank')
            file_path = i
            img_path = obs.citi_split_pdf(file_path)
            #
            ocr_info, cap_dic = obs.ocr_recognize(img_path)
            df, df_date = obs.tabulate_citibank(ocr_info)
            df.to_excel(data_file)
            df_date.to_excel(coor_file)

            if threading.currentThread().getName() != 'MainThread':
                pythoncom.CoInitialize()
            excel.excel_coor()
            files.remove(i)

    for i in files:
        if 'facility' in i.lower():
            file_path = i
            img_path = ls.dbs_loan_split_pdf(file_path)
            seg_dir = ls.segment_img(img_path)
            info_df = ls.read_loan(seg_dir)
            info_df.to_excel('loanstatement.xlsx')
            excel.adder('Loan Repayment', loan_file)
            files.remove(i)

    for i in files:
        if 'bill' in i.lower():
            ob.dbs_split_pdf(i)
            cutoff_date, expiry_date, outstanding_balance = ob.ocr_dbs_bill_statement("billsample/")
            df = ob.tabulate_dbs_bill_statement(cutoff_date, expiry_date, outstanding_balance)
            df.to_excel("billstatement.xlsx")
            excel.adder('Bill Statement', bill_file)
            files.remove(i)

    for i in files:
        if 'cat' in i.lower():
            excel.CAT_processor(i)
            files.remove(i)

    return render_template('result_1.html')



if __name__ == "__main__":
    app.run(debug=True, use_reloader=False)