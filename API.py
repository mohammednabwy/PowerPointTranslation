#Libraries
import warnings
warnings.filterwarnings("ignore")
from flask import Flask,send_file,send_from_directory
from flask import  request
from werkzeug.utils import secure_filename
from Service import PPTTranslationService,WordTranslationService
import os
#--------------------------------------------------------
#Public Variables
UPLOAD_FOLDER = '/Files/'
UPLOAD_FOLDER_LOs=UPLOAD_FOLDER
ALLOWED_EXTENSIONS = {'ppt', 'pptx','doc','docx','rtf', 'pdf'}
FILES_URL_Download='http://127.0.0.1:5000/LOs/'
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
APP = Flask(__name__)

#--------------------------------------------------------
#Functions
def allowed_file(filename):
    return '.' in filename and \
        filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
#-------------------------------------------------------- 

@app.route('/favicon.ico')
def favicon():
    return send_from_directory(os.path.join(app.root_path, 'static'),
                               'favicon.ico', mimetype='image/favicon.png')

@app.route('/')
@app.route('/home')
def home():
    return "Welcome to File Translator" 

@app.route('/api/v1/TranslateFile', methods=["POST"])
def TranslatePowerPointFile():    
    file,target=None,''
    try:        
        file = request.files['file']       
        target=request.args.get('target')
    except:
         return "No file " ,406   
    if not allowed_file(file.filename): 
        return "Not allowed file extension",406
    filename = secure_filename(file.filename)    
    filePath=filename
    file.save(filePath)      
    filename, file_extension = os.path.splitext(filePath)   
    translated_file_path=None
    print("Translating File Started..............")
    if file_extension in ('.ppt','.pptx'): 
        translated_file_path=PPTTranslationService.translatePPTFile(filePath,target=target)  
    elif file_extension in ('.doc','.docx','.rtf'): 
        translated_file_path=WordTranslationService.translateWordFile(filePath,target=target) 
    print("Translating File Finished")
    output_file_path=translated_file_path
    output_file_path=output_file_path.replace('\\','/')            
    return send_file(output_file_path, as_attachment=True)  
#-----------------------------------------------------

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000)