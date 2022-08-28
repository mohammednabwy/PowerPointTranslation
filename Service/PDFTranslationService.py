from pdf2docx import Converter 
import os
from Service import WordTranslationService


def convertEditablePDF2Word(pdf_file_path,out_docx_file_path): 
  cv = Converter(pdf_file_path)
  cv.convert(out_docx_file_path)
  cv.close() 
  return 


def translatePDFFile(pdffilePath,target='ar'):  
    #convert pdf to word
    pdf_file_directory=os.path.dirname(pdffilePath) 
    word_file_name=os.path.splitext(pdffilePath)[0]  
    word_file_path = os.path.join(pdf_file_directory,word_file_name)
    convertEditablePDF2Word(pdffilePath,word_file_path)
    #translate word file
    WordTranslationService.translateWordFile(word_file_path,target)
    return


def main(): 
     pdffilepath='Files\\en_2.pdf'   
     translatePDFFile(pdffilepath,target='ar')  
     return