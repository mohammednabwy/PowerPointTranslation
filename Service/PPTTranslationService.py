import os
from pptx import Presentation
from deep_translator import GoogleTranslator #https://github.com/nidhaloff/deep-translator
#---------------------------------------------
def translate(source_txt,source='auto', target='fr'):
   translated = GoogleTranslator(source=source, target=target).translate(source_txt) 
   return  translated
#---------------------------------------------
def savePowerPoint(filename,presentation):
    presentation.save(filename)
    return
#---------------------------------------------
def translatePPTFile(pptfilePath,src='auto',target='fr'):     
    presentation = Presentation(pptfilePath)
    for slideIndex in range(0,len(presentation.slides)):       
        slide=presentation.slides[slideIndex]    
        if(slide.shapes!=None and slide.shapes.title!=None and len(slide.shapes.title.text)>0):
            slideTitle = slide.shapes.title.text  
            if len(slideTitle) < 3 : continue
            slide.shapes.title.text =translate(slideTitle)
        for i in range(0,len(slide.shapes)):        
           if not slide.shapes[i].has_text_frame: continue
           paragraphs=slide.shapes[i].text_frame.paragraphs
           for j in range(0,len(paragraphs)):               
               content=slide.shapes[i].text_frame.paragraphs[j].text    
               if len(content) < 3 : continue
               slide.shapes[i].text_frame.paragraphs[j].text=translate(content) 
    #save 
    file_name=os.path.basename(pptfilePath)
    file_directory=os.path.dirname(pptfilePath) 
    translated_file_name='translated_' +  file_name.split('.')[0] + '.'  +file_name.split('.')[1]
    translated_file = os.path.join(file_directory, translated_file_name)
    savePowerPoint(translated_file,presentation)
    return translated_file
#---------------------------------------------
def main():
    pptfilename='Files\\DataStructureAndAlgorthimDesign.pptx'  
    translatePPTFile(pptfilename,target='fr')   
    return