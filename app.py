import os
from flask import Flask, flash, request, redirect, render_template, send_file, send_from_directory, url_for
from werkzeug.utils import secure_filename
from Helpers import read_directory_files_unsignalized
from datetime import datetime

app=Flask(__name__)

app.secret_key = "secret key"
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

# Get current path
path = os.getcwd()
# file Upload
UPLOAD_FOLDER = os.path.join(path, 'uploads')
PEOPLE_FOLDER = os.path.join('static', 'images')
app.config['IMAGE_FOLDER'] = PEOPLE_FOLDER

# Make directory if uploads is not exists
if not os.path.isdir(UPLOAD_FOLDER):
    os.mkdir(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Allowed extension you can set your own
ALLOWED_EXTENSIONS = set(['txt', 'pdf', 'png', 'jpg', 'jpeg', 'gif'])


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def upload_form():
    full_filename = os.path.join(app.config['IMAGE_FOLDER'], 'upload.svg')
    return render_template("upload.html", user_image=full_filename)
   # return render_template('upload.html')


@app.route('/', methods=['POST'])
def upload_file():
    if request.method == 'POST':

        if 'files[]' not in request.files:
            flash('No file part')
            return redirect(request.url)

        files = request.files.getlist('files[]')

        if 'files[]' not in request.files:
            flash('No file part')
            return redirect(request.url)

        files = request.files.getlist('files[]')
        read_directory_files_unsignalized(files, 'Report_Unsignalized_Summary')
        flash('File(s) successfully uploaded')
        #return render_template('download.html')
        now = datetime.now()
        dt_string = now.strftime("%d-%m-%Y_%H%M")
        return send_file(f'Report_Unsignalized_Summary_{dt_string}.docx', as_attachment=True)

@app.route('/download')
def download():
    now = datetime.now()
    dt_string = now.strftime("%d-%m-%Y_%H%M")
    return send_file(f'Report_Signalized_Summary_{dt_string}.docx', as_attachment=True)

# main method where flow starts.
if __name__ == "__main__":
    app.run(debug=True,host='0.0.0.0')
