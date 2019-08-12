from flask import Flask, redirect, render_template, request, session, url_for, send_from_directory
from flask_socketio import SocketIO, emit
from flask_dropzone import Dropzone
from time import sleep
from flask_uploads import UploadSet, configure_uploads, IMAGES, patch_request_class
from werkzeug.utils import secure_filename
from excel import creator
from processor import processor
from threading import Thread, Event
import win32com.client as win32
import os
import random, threading, webbrowser
from datetime import datetime



app = Flask(__name__)
dropzone = Dropzone(app)

socketio = SocketIO(app)

thread = Thread()
thread_stop_event = Event()


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
temp_loc = APP_ROOT + "\\static\\template"

xl = win32.gencache.EnsureDispatch('Excel.Application')
xl.Visible = False  # change to False after development
xl.DisplayAlerts = False
wb = xl.Workbooks.Open(os.path.join(temp_loc, 'template.xlsm'))
xl.Quit()
try:
    os.remove(os.path.join(output_loc, 'output.xlsm'))
except:
    pass

class sock():
    def __init__(self, socket):
        self.socket = socket
        self.progress = 0
        self.eta = 96

    def sprint(self, message, progress, eta):
        self.progress = min(100, self.progress + progress)
        self.eta = max(0, self.eta - eta)
        self.socket.emit('sysmessage', {'message': message, 'progress': self.progress, 'eta': self.eta}, namespace='/test')

    def direct(self):
        self.socket.emit('redirect', {'url': '/process'}, namespace='/test')

printer = sock(socketio)

def gen():
    printer.sprint("Starting...", 0, 0)
    creator(files, printer)
    printer.sprint("Finalizing Output...", 5, 0)
    #processor()
    printer.sprint("Completed!", 100, 0)
    # for i in range(1):
    #     printer.sprint('hello' + str(i), i, i+30)
    #     #printer.sprint('hello' + str(i+2), i+90, i+30)
    #     sleep(5)
    #     printer.sprint('hello' + str(i), i+90, i + 30)
    #
    printer.direct()
    # #     socketio.emit('sysmessage', {'message': 'hello' + str(i), 'progress': i, 'eta': i+30}, namespace='/test')
    # #     sleep(5)
    # #     socketio.emit('sysmessage', {'message': 'hello1234' + str(i), 'progress': i+95, 'eta': i+33}, namespace='/test')
    # #socketio.emit('redirect', {'url': '/process'}, namespace='/test')

class Creator(Thread):
    def __init__(self):
        super(Creator, self).__init__()

    def function(self):
        gen()

    def run(self):
        gen()

@app.after_request
def add_header(response):
    """
    Add headers to both force latest IE rendering engine or Chrome Frame,
    and also to cache the rendered page for 10 minutes.
    """
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate, public, max-age=0"
    response.headers["Expires"] = 0
    response.headers["Pragma"] = "no-cache"
    return response

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
    return send_from_directory(save_path, 'output.xlsm', as_attachment=True,
                               attachment_filename='output{}.xlsm'.format(str(datetime.now())[-6:]))
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
    # creator(files)
    # processor()
    return render_template('result_1.html')

@socketio.on('connect', namespace='/test')
def test_connect():
    # need visibility of the global thread object
    #global thread
    global thread
    print('Client connected')

    # Start the random number generator thread only if the thread has not been started before.
    if not thread.isAlive():
        print("Starting Thread")
        thread = Creator()
        thread.start()



@socketio.on('disconnect', namespace='/test')
def test_disconnect():
    print('Client disconnected')



if __name__ == "__main__":
    port = 5000 #+ random.randint(0, 999)
    url = "http://127.0.0.1:{0}".format(port)
    threading.Timer(1.25, lambda: webbrowser.open(url)).start()
    socketio.run(app)