#Libraries
import warnings
warnings.filterwarnings("ignore")
from flask import Flask,send_from_directory
from flask import  request
from werkzeug.utils import secure_filename
from Service import PPTTranslationService,WordTranslationService,PDFTranslationService
import os
import json 
#--------------------------------------------------------
#Public Variables
UPLOAD_FOLDER = 'Files/'
ALLOWED_EXTENSIONS = {'ppt', 'pptx','doc','docx','rtf', 'pdf'}
#FILES_URL_Download='http://127.0.0.1:5000/'
FILES_URL_Download='https://filestranslation.herokuapp.com/'
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
APP = Flask(__name__)

#--------------------------------------------------------
#Functions
def allowed_file(filename):
    return '.' in filename and \
        filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
#-------------------------------------------------------- 

@app.route("/Files/<path:path>")
def get_file(path):
    """Download a file."""
    return send_from_directory(UPLOAD_FOLDER, path, as_attachment=True)

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
    filePath=os.path.join(UPLOAD_FOLDER,filename)
    file.save(filePath)      
    filename, file_extension = os.path.splitext(filePath)   
    translated_file_path=''
    print("Translating File Started..............")
    if file_extension in ('.ppt','.pptx'): 
        translated_file_path=PPTTranslationService.translatePPTFile(filePath,target=target)  
    elif file_extension in ('.doc','.docx','.rtf'): 
        translated_file_path=WordTranslationService.translateWordFile(filePath,target=target) 
    elif file_extension == '.pdf':
        translated_file_path=PDFTranslationService.translatePDFFile(filePath)
    print("Translating File Finished")
    output_file_path=translated_file_path
    output_file_path=output_file_path.replace('\\','/')    
    dictionary ={ 
      "output_file_download_url":  FILES_URL_Download + output_file_path   
    }    
    json_object = json.dumps(dictionary, indent = 4)   
    return json_object
    #return send_file(output_file_path, as_attachment=True)  
#-----------------------------------------------------

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000)