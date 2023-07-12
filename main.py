from flask import Flask, request, render_template, flash, redirect, send_file, send_from_directory, url_for
from flask_cors import CORS
import json
import os 
import traceback
import sys
import functions as fn
import shutil
from jinja2 import Environment, FileSystemLoader

CURRENT_DIR = os.path.dirname(__file__)
TMP_DIR = '/tmp'

UPLOAD_FOLDER = f'{TMP_DIR}/uploads'
ALLOWED_EXTENTIONS = {'json','xlsx'}
OUTPUT_FOLDER = f'{TMP_DIR}/output'
ZIP_FOLDER = f'{TMP_DIR}/downloadables'

app = Flask(__name__)
app.secret_key = 'super secret key'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
CORS(app)

@app.route('/')
def main():
    return render_template('index.html')

@app.route('/scenario_viewer/<path:filename>')
def serve_static(filename):
    return send_from_directory(f'{OUTPUT_FOLDER}/website', filename)

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        # clear the temp folders child folders but maintain the structure
        fn.clear_dir(UPLOAD_FOLDER)
        fn.clear_dir(OUTPUT_FOLDER)
        fn.clear_dir(ZIP_FOLDER)

        if request.method == 'POST':
            # if file not uploaded
            if 'file' not in request.files:
                flash('No file part')
                return redirect(request.url)
            
            # get files from request 
            files = request.files.getlist('file')

            # save files
            personae = []
            json_files = False
            xlsx_files = False
            for file in files:
                if file.filename == '':
                    print('No selected file')
                    flash('No file selected')
                    return redirect(url_for('main'))
                file.save(os.path.join(UPLOAD_FOLDER, file.filename))
                if file.filename.split('.')[-1] != 'json':
                    personae.append(file.filename)
                    xlsx_files = True
                elif file.filename.split('.')[-1] == 'json':
                    json_files = True

            # copy website contents only if needed
            if xlsx_files:
                shutil.copytree(f'{CURRENT_DIR}/data view website contents', f'{OUTPUT_FOLDER}/website', dirs_exist_ok=True)

                # create dropdown gen
                # render log 
                env = Environment(loader=FileSystemLoader('templates'))
                template = env.get_template('index_template.jinja')
                output = template.render(personae=personae)

                with open(f'{OUTPUT_FOLDER}/website/index.html', 'w') as f:
                    f.write(output)

            # process files
            summaries = []
            for file in os.listdir(UPLOAD_FOLDER):
                if file.split('.')[-1] != 'json':
                    summary = fn.process_file(file, OUTPUT_FOLDER, UPLOAD_FOLDER, f'{CURRENT_DIR}/standards')
                    with open(f'{OUTPUT_FOLDER}/{file[:-5]}/summary.json', 'w') as f:
                        json.dump(summary, f, indent=4)
                    summaries.append(summary)
                else:
                    _ = fn.process_file(file, OUTPUT_FOLDER, UPLOAD_FOLDER, f'{CURRENT_DIR}/standards')

            if xlsx_files:
                # create excel summary work book
                fn.create_false_path_excel(OUTPUT_FOLDER, summaries)

            # zip archive output folder
            shutil.make_archive(f'{ZIP_FOLDER}/output', 'zip', OUTPUT_FOLDER)


        os.system(f'gcloud app deploy {OUTPUT_FOLDER}/website/app.yaml -q --project scenario-viewer')

        return render_template('file_download.html', summaries = summaries, xlsx_files = xlsx_files, json_files=json_files)
    except Exception as e:
        print(type(e).__name__)
        print(e.__traceback__.tb_lineno)
        print(traceback.format_exc())
        return render_template('error.html', line=e.__traceback__.tb_lineno, error=e)


@app.route('/download', methods=['GET', 'POST'])
def download():
 #   result = send_from_directory(directory=ZIP_FOLDER, path=f'{ZIP_FOLDER}/output.zip', filename='output.zip')
    try:
        return  send_file(f'{ZIP_FOLDER}/output.zip', as_attachment=True)
    except Exception as e:
        error = traceback.format_exc()
        return render_template('error.html', error=error)   


if __name__ == '__main__':
    app.run(debug=True, port=8000)
 