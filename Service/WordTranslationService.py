import os
from docx import Document

from deep_translator import GoogleTranslator #https://github.com/nidhaloff/deep-translator
#---------------------------------------------
def translate(source_txt,target='fr'):
   translated = GoogleTranslator(source='auto', target=target).translate(source_txt) 
   return  translated
#---------------------------------------------
def save(filename,document):
    document.save(filename)
    return

def translateWordFile(wordfilePath,target='fr'): 
    document=Document(wordfilePath)
    for paragraph in document.paragraphs:     
        for run in paragraph.runs:     
            try:
                cur_text = run.text
                if len(cur_text) < 3 : continue
                new_text = translate(cur_text,target) 
                run.text = new_text 
            except:
                continue
    
    for table in document.tables:
        try:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:                  
                            paragraph.text = translate(paragraph.text,target) 
        except:
                continue
    
    #save 
    file_name=os.path.basename(wordfilePath)
    file_directory=os.path.dirname(wordfilePath) 
    translated_file_name='translated_' + target +"_" +  file_name.split('.')[0] + '.'  +file_name.split('.')[1]
    translated_file = os.path.join(file_directory, translated_file_name)
    save(translated_file,document)
    return translated_file

def main():
    wordfilePath='Files\\test2.docx'  
    translateWordFile(wordfilePath,target='fr')   
    return