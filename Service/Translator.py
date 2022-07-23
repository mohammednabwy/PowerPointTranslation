from deep_translator import GoogleTranslator #https://github.com/nidhaloff/deep-translator
#---------------------------------------------

def translate(source_txt,target='fr'):
   translated_text = GoogleTranslator(source='auto', target=target).translate(source_txt) 
   return  translated_text