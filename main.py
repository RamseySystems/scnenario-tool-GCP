from flask import Flask, request, render_template, flash, redirect, send_file, send_from_directory, url_for, session
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
ALLOWED_EXTENTIONS = {'json','xlsx'}

app = Flask(__name__, static_url_path='/static')
app.secret_key = 'super secret key'
app.config['UPLOAD_FOLDER'] = TMP_DIR
CORS(app)

@app.route('/', methods=['GET', 'POST'])
def get_user():
    if request.method == 'POST':
        session['user'] = request.form['user_name']
        print(session['user'])
        return redirect(url_for('scenario_tool'))
    
    return """
            <form method="post">
                <label for="user_name">Enter your name:</label>
                <input type="text" id="user_name" name="user_name" required />
                <button type="submit">Submit</button>
            </form>
            """

@app.route('/scenario_tool')
def scenario_tool():
    return render_template('index.jinja')

@app.route('/scenario_viewer/<path:filename>')
def serve_static(filename):
    return send_from_directory(f'{session["user"]}/output/website', filename)

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        # clear the temp folders child folders but maintain the structure
        fn.clear_dir(f'{TMP_DIR}/{session["user"]}/upload')
        fn.clear_dir(f'{TMP_DIR}/{session["user"]}/output')
        fn.clear_dir(f'{TMP_DIR}/{session["user"]}/downloadables')

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
                file.save(os.path.join(f'{TMP_DIR}/{session["user"]}/upload', file.filename))
                if file.filename.split('.')[-1] != 'json':
                    personae.append(file.filename)
                    xlsx_files = True
                elif file.filename.split('.')[-1] == 'json':
                    json_files = True

            # copy website contents only if needed
            if xlsx_files:
                shutil.copytree(f'{CURRENT_DIR}/data view website contents', f'{TMP_DIR}/{session["user"]}/output/website', dirs_exist_ok=True)

                # create dropdown gen
                # render log 
                env = Environment(loader=FileSystemLoader('templates'))
                template = env.get_template('index_template.jinja')
                output = template.render(personae=personae)

                with open(f'{TMP_DIR}/{session["user"]}/output/website/index.html', 'w') as f:
                    f.write(output)

                # get provenance paths
                if 'Provenance.xlsx' in os.listdir(f'{TMP_DIR}/{session["user"]}/upload'):
                    provenance_paths = fn.get_standard_paths(f'{TMP_DIR}/{session["user"]}/upload/Provenance.xlsx', [])

                # make a list of standard paths
                main_path_list = []
                for file in os.listdir(f'{TMP_DIR}/{session["user"]}/upload'):
                    if file[0] == '_':
                        current_standard_path = f'{TMP_DIR}/{session["user"]}/upload/{file}'
                        standard_paths = fn.get_standard_paths(current_standard_path, provenance_paths)
                        main_path_list.append(standard_paths.copy())


            # process scenario files
            summaries = []
            for file in os.listdir(f'{TMP_DIR}/{session["user"]}/upload'):
                if file != 'Provenance.xlsx' and file[0] != '_':
                    if file.split('.')[-1] != 'json':
                        summary = fn.process_file(file, f'{TMP_DIR}/{session["user"]}/output', f'{TMP_DIR}/{session["user"]}/upload', main_path_list)
                        with open(f'{TMP_DIR}/{session["user"]}/output/{file[:-5]}/summary.json', 'w') as f:
                            json.dump(summary, f, indent=4)
                        summaries.append(summary)
                    else:
                        _ = fn.process_file(file, f'{TMP_DIR}/{session["user"]}/output', f'{TMP_DIR}/{session["user"]}/upload', main_path_list)

            if xlsx_files:
                # create excel summary work book
                fn.create_false_path_excel(f'{TMP_DIR}/{session["user"]}/output', summaries)

            # zip archive output folder
            shutil.make_archive(f'{TMP_DIR}/{session["user"]}/downloadables/output', 'zip', f'{TMP_DIR}/{session["user"]}/output')


        os.system(f'gcloud app deploy {TMP_DIR}/{session["user"]}/output/website/app.yaml -q --project scenario-viewer')

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
        return  send_file(f'{TMP_DIR}/{session["user"]}/downloadables/output.zip', as_attachment=True)
    except Exception as e:
        error = traceback.format_exc()
        return render_template('error.html', error=error)   


if __name__ == '__main__':
    app.run(debug=True, port=8000)
 