from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import base64
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
import mimetypes
import os
from docx import Document
import pandas as pd
import sys
import re
import subprocess
from sys import argv
from string import Template


# Email variables. Modify this!
# EMAIL_FROM = 'harry.j.durbin@gmail.com'
# EMAIL_TO = 'harry.durbin@gmail.com'
# EMAIL_SUBJECT = 'Hellkklkkkko'
# EMAIL_CONTENT = 'Hello, this is a testalkjdflkjadlkfj'


df = pd.read_excel('jobs.xlsx')

def read_template(filename):
    with open(filename, 'r') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)

# If modifying these scopes, delete the file token.pickle.

SCOPES = ["https://www.googleapis.com/auth/gmail.compose","https://mail.google.com/", "https://www.googleapis.com/auth/gmail.modify"]
#SCOPES = ['https://www.googleapis.com/auth/gmail.compose','https://mail.google.com','https://www.googleapis.com/auth/gmail.modify']
#SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']'https://www.googleapis.com/auth/gmail.modify',

def service_account_login():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)
    return service

def create_draft(service, user_id, message_body):
    message = {'message': message_body}
    draft = service.users().drafts().create(userId=user_id, body=message).execute()
    return draft

def create_content(name,email,position,foldername,company,int,skills,traits,role,focus,location,place,ln1,ln2,ln3,addressee,site,filename):
    message_template = read_template(filename)
    content = message_template.substitute(PERSON_NAME=name, POSITION=position,COMPANY=company,
    INT=int, SKILLS=skills,TRAITS=traits,ROLE=role,FOCUS=focus,LOCATION=location,PLACE=place,
    LN1=ln1,LN2=ln2,LN3=ln3,ADDRESSEE=addressee,SITE=site)

    return content

def create_message_with_attachments(sender, to, subject, message_text):

    df = pd.read_excel('jobs.xlsx')
    n = len(df)-1
    dic = df.loc[n].to_dict()
    dictionary = dic
    no = str(dic['NO'])
    company1 = str(dic['COMPANY']).replace(' ','')
    foldername = no+"_"+company1

    msg = MIMEMultipart()
    msg['to'] = to
    msg['from'] = sender
    msg['subject'] = subject

    msg.attach(MIMEText(message_text, 'plain'))

    fns = ["HarryDurbin_Letter.pdf","HarryDurbin_Resume.pdf" ]
    for fn in fns:
        attachment = open(foldername+"/"+fn, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % fn)
        msg.attach(part)

    return {'raw': base64.urlsafe_b64encode(msg.as_string().encode()).decode()}

if __name__ == '__main__':


    if argv[1] == 'en':
        filename = 'message.txt'
    elif argv[1] == 'w':
         filename = 'messagew.txt'
    elif argv[1] == 'c':
        filename = 'messagew.txt'
    else:
        print ('You need to give an argument!')

    n = len(df)-1
    dic = df.loc[n].to_dict()
    dictionary = dic

    no = str(dic['NO'])
    company = str(dic['COMPANY']) #.replace(' ','')
    foldername = no+"_"+company
    email = dic['CONTACT']
    name = dic['PERSON']
    #person = dic['PERSON']
    position = dic['POSITION']
    int = dic['INT']
    skills = dic['SKILLS']
    traits = dic['TRAITS']
    role = dic['ROLE']
    focus = dic['FOCUS']
    location = dic['LOCATION']
    place = ''
    ln1 = dic['LN1']
    ln2 = dic['LN2']
    ln3 = dic['LN3']
    addressee = dic['ADDRESSEE']
    site = dic['SITE']

    if email==str(email):

        service = service_account_login()

        EMAIL_FROM = 'harry.j.durbin@gmail.com'
        EMAIL_TO = email.lower()
        EMAIL_SUBJECT = position

        print (foldername)

        EMAIL_CONTENT = create_content(name,email,position,foldername,company,int,skills,traits,role,focus,location,place,ln1,ln2,ln3,addressee,site,filename)

        message = create_message_with_attachments(EMAIL_FROM, EMAIL_TO, EMAIL_SUBJECT, EMAIL_CONTENT)

        create_draft(service,'me', message)

        print ('Bingo bango--created draft email for ' + name + ' at ' + email)
    else:
        print ('No email provided for {}!'.format(company))






    # setup the parameters of the message
    # msg['From']='Harrison Durbin <' + MY_ADDRESS + '>'
    # msg['To']=email.lower()
    # msg['Subject']= " "+position

    # add in the message body
    # msg.attach(MIMEText(message, 'plain'))
    #
    # fns = ["HarryDurbin_Letter.pdf","HarryDurbin_Resume.pdf" ]
    # for fn in fns:
    #     attachment = open(foldername+"/"+fn, "rb")
    #
    #     part = MIMEBase('application', 'octet-stream')
    #     part.set_payload((attachment).read())
    #     encoders.encode_base64(part)
    #     part.add_header('Content-Disposition', "attachment; filename= %s" % fn)
    #
    #     msg.attach(part)
