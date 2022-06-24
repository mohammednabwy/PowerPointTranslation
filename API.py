#Libraries
import warnings
warnings.filterwarnings("ignore")
from flask import Flask,send_file
from flask import  request
from werkzeug.utils import secure_filename
from Service import PPTTranslationService
import os
#--------------------------------------------------------
#Public Variables
UPLOAD_FOLDER = 'Files/'
UPLOAD_FOLDER_LOs=UPLOAD_FOLDER
ALLOWED_EXTENSIONS = {'ppt', 'pptx'}
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
@app.route('/api/v1/TranslatePowerPointFile', methods=["POST"])
def generateLOsFromFiles():    
    pptFile,source,target=None,'','' 
    try:        
        pptFile = request.files['pptFile']
        source=request.args.get('source') 
        target=request.args.get('target')
    except:
         return "No file " ,406   
    if not allowed_file(pptFile.filename): 
        return "Not allowed file extension",406
    pptfilename = secure_filename(pptFile.filename)
    pptfilePath=os.path.join(app.config['UPLOAD_FOLDER'], pptfilename)
    pptFile.save(pptfilePath)    
    print("Translating File Started..............")
    translated_file_path=PPTTranslationService.translatePPTFile(pptfilePath,src=source,target=target)  
    print("Translating File Finished")
    output_file_path=translated_file_path
    output_file_path=output_file_path.replace('\\','/')            
    return send_file(output_file_path, as_attachment=True)  
#-----------------------------------------------------
if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000)