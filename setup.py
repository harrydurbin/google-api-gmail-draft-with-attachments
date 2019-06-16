
from docx import Document
import pandas as pd
import sys
import re
import subprocess
from sys import argv
from string import Template
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import sys
import warnings

if not sys.warnoptions:
    warnings.simplefilter("ignore")

df = pd.read_excel('jobs.xlsx')

def docx_replace_regex(doc_obj, regex , replace):
    for p in doc_obj.paragraphs:
        if regex.search(p.text):
            inline = p.runs
            for i in range(len(inline)):
                if regex.search(inline[i].text):
                    text = regex.sub(replace, inline[i].text)
                    inline[i].text = text

if __name__ == '__main__':
    df = pd.read_excel('jobs.xlsx')
    n = len(df)-1
    dic = df.loc[n].to_dict()
    dictionary = dic
    no = str(dic['NO'])
    company1 = str(dic['COMPANY']).replace(' ','')
    foldername = no+"_"+company1

    if argv[1] == 'en':
        filename = 'HarryDurbin_Letter_en_TEMPLATE.docx'
        bashCommand1 = 'cp HarryDurbin_Resume_en.pdf '+ foldername+'/HarryDurbin_Resume.pdf'

    elif argv[1] == 'gis':
      filename = 'HarryDurbin_Letter_en_TEMPLATE.docx'
      bashCommand1 = 'cp HarryDurbin_Resume_gis.pdf '+ foldername+'/HarryDurbin_Resume.pdf'

    elif argv[1] == 'data':
       filename = 'HarryDurbin_Letter_en_TEMPLATE.docx'
       bashCommand1 = 'cp HarryDurbin_Resume_data.pdf '+ foldername+'/HarryDurbin_Resume.pdf'

    elif argv[1] == 'w':
       filename = 'HarryDurbin_Letter_w_TEMPLATE.docx'
       bashCommand1 = 'cp HarryDurbin_Resume_w.pdf '+ foldername+'/HarryDurbin_Resume.pdf'

    elif argv[1] == 'c':
        filename = 'HarryDurbin_Letter_w_TEMPLATE.docx'
        bashCommand1 = 'cp HarryDurbin_Resume_c.pdf '+ foldername+'/HarryDurbin_Resume.pdf'

 #   elif argv[1] == 'es':
 #       filename = 'HarryDurbin_Carta_es_TEMPLATE.docx'
 #       bashCommand1 = 'cp HarryDurbin_Curriculum.pdf '+ foldername+'/HarryDurbin_CV.pdf'
   #elif argv[1] == 'w1':
   #    filename = 'HarryDurbin_Letter_w_TEMPLATE.docx'
   #    bashCommand1 = 'cp HarryDurbin_Resume_w1.pdf '+ foldername+'/HarryDurbin_Resume.pdf'
   #elif argv[1] == 'c1':
   #    filename = 'HarryDurbin_Letter_c_TEMPLATE.docx'
   #    bashCommand1 = 'cp HarryDurbin_Resume_w1.pdf '+ foldername+'/HarryDurbin_Resume.pdf'


    doc = Document(filename)
    for word, replacement in dictionary.items():
        word_re=re.compile(word)
        docx_replace_regex(doc, word_re , replacement)
    # newfilename = filename.replace('_TEMPLATE','')
    newfilename = 'HarryDurbin_Letter.docx'

    bashCommand = 'mkdir ' + foldername
    process = subprocess.Popen(bashCommand.split(), stdout=subprocess.PIPE)
    output, error = process.communicate()
    doc.save(foldername+'/'+newfilename)

    # bash command 1

    process = subprocess.Popen(bashCommand1.split(), stdout=subprocess.PIPE)
    output, error = process.communicate()

    docpath = '/home/hd/Documents/jobs/'+ foldername + '/' + 'HarryDurbin_Letter.docx'
    destpath = '/home/hd/Documents/jobs/'+ foldername
    bashCommand = 'sudo lowriter --headless --convert-to pdf --outdir ' + destpath + ' ' + docpath
    process = subprocess.Popen(bashCommand.split(), stdout=subprocess.PIPE)
    output, error = process.communicate()

    print ('Bingo bango -- all set up for {}'.format(company1))
