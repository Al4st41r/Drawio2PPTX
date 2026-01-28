import os
import uuid
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
import sys

# Add parent directory to path to allow importing converter
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from converter import convert

app = Flask(__name__)
app.secret_key = 'supersecretkey' # Change this for production
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'uploads')
app.config['OUTPUT_FOLDER'] = os.path.join(os.path.dirname(__file__), 'outputs')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024 # 16MB limit

ALLOWED_EXTENSIONS = {'drawio', 'xml'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        
        file = request.files['file']
        
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
            
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            unique_id = str(uuid.uuid4())
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{unique_id}_{filename}")
            output_filename = f"{os.path.splitext(filename)[0]}.pptx"
            output_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{unique_id}_{output_filename}")
            
            file.save(input_path)
            
            try:
                convert(input_path, output_path)
                return send_file(output_path, as_attachment=True, download_name=output_filename)
            except Exception as e:
                flash(f"Error converting file: {str(e)}")
                return redirect(request.url)
            finally:
                # Cleanup could happen here or via a cron job
                pass
                
    return render_template('index.html')

if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
    app.run(host='0.0.0.0', port=5003)
