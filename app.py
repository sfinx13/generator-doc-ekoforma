import zipfile
import os
import datetime
from flask import Flask , render_template, request, redirect, flash, send_from_directory, send_file
import services_handler as service_handler

app = Flask(__name__)
app.config['SECRET_KEY'] = 'secret_key'
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # Limite de 16 Mo
app.config['GENERATED_FILES_FOLDER'] = 'downloads/'

ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

# Activer la toolbar de débogage

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def clean_upload_dir():
    for filename in os.listdir(app.config['UPLOAD_FOLDER']):
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
            os.remove(file_path)

@app.before_request
def log_request_info():
    app.logger.info('Headers: %s', request.headers)
    app.logger.info('Body: %s', request.get_data())

@app.route("/")
def home():
    # Clean upload file
    if os.path.exists(app.config['UPLOAD_FOLDER']):
        clean_upload_dir()
    # Lister les fichiers dans le répertoire
    files = os.listdir(app.config['GENERATED_FILES_FOLDER'])
    valid_files = [f for f in files if f.endswith('.xlsx') or f.endswith('.docx') or f.endswith('.pdf')]
    sorted_files = sorted(valid_files)
    return render_template('index.html', title='Génération de documents administratifs', files=sorted_files)

@app.route("/upload_file", methods=['POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('Pas de fichier selectionné', 'error')

        file = request.files.get('file')

        if file.filename == '':
            flash('Pas de fichier selectionné', 'error')
    
        if file and allowed_file(file.filename):
            clean_upload_dir()
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)
            service_handler.generate_timesheet_zoom()
            service_handler.generate_attendance_certificates()

            flash(f'Fichier {file.filename} uploadé avec succès', 'success')
            os.remove(filepath)

        return redirect("/")
        

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['GENERATED_FILES_FOLDER'], filename, as_attachment=True)

@app.route('/download_all_files')
def download_all_files():
    try:
        folder_path = app.config['GENERATED_FILES_FOLDER']
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_filename = f"tous_les_documents_{timestamp}.zip"
        zip_path = os.path.join(folder_path, zip_filename)

        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for root, dirs, files in os.walk(folder_path):
                for file in files:
                    print(file)
                    if 'xlsx' in file or 'docx' in file or 'pdf' in file:
                        file_path = os.path.join(root, file)
                        zipf.write(file_path, os.path.relpath(file_path, folder_path))

        return send_file(zip_path, as_attachment=True)

    except Exception as e:
        flash(f"Une erreur s'est produite lors de la génération de l'archive ZIP : {e}", 'error')
        return redirect('/')


@app.route('/delete/<filename>')
def delete_file(filename):
    try:
        file_path = os.path.join(app.config['GENERATED_FILES_FOLDER'], filename)

        if os.path.exists(file_path):
            os.remove(file_path)
            flash(f"Le fichier {filename} a été supprimé avec succès.", 'success')
        else:
            flash(f"Le fichier {filename} n'existe pas.", 'error')

    except Exception as e:
        flash(f"Une erreur s'est produite lors de la suppression du fichier: {e}", 'error')
    
    return redirect('/download')

@app.route('/delete_all')
def delete_all_files():
    try:
        folder_path = app.config['GENERATED_FILES_FOLDER']

        if os.path.exists(folder_path):
            for filename in os.listdir(folder_path):
                file_path = os.path.join(folder_path, filename)
                
                if file_path.endswith('.pdf') or file_path.endswith('.docx') or file_path.endswith('xlsx'):
                    os.remove(file_path)

            flash("Le contenu du répertoire a été supprimé avec succès.", 'success')
        else:
            flash("Le répertoire n'existe pas.", 'error')

    except Exception as e:
        flash(f"Une erreur s'est produite lors de la suppression du contenu: {e}", 'error')

    return redirect('/')



if __name__ == '__main__':
    app.run()