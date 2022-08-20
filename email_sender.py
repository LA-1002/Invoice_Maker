from __future__ import print_function

import base64
from copyreg import pickle
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import mimetypes

from googleapiclient.discovery import build

import os
import pickle
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import base64;



def authenticate():
    SCOPES = ['https://mail.google.com/']
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle','rb') as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
    
        else:
            flow = InstalledAppFlow.from_client_secrets_file('oAuthKey.json',SCOPES)
            creds = flow.run_local_server(port=0)

    with open('token.pickle','wb') as token:
        pickle.dump(creds,token)
    return build('gmail','v1',credentials=creds);
    



def contents(destination,subject,body,files):
    message = MIMEMultipart(body);

    #Main Body of the Email
    message["to"] = destination
    message["from"] = "traveleco.developer@gmail.com"
    message["subject"] = subject   
    for file in files:
        attachment(message,file)

    return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}


def sendEmail(service,destination,subject,body,files):
    return service.users().messages().send(
        userId='me',
        body=contents(destination,subject,body,files)).execute()


        

def attachment(message,file):
    content_type, encoding = mimetypes.guess_type(file)

    if content_type is None or encoding is not None:
        content_type = 'application/octet-stream'
    main_type, sub_type = content_type.split('/',1)
    if main_type == 'text':
        with open(file,'rb') as fp:
            msg = MIMEText(fp.read().decode(), _subtype=sub_type)
    elif main_type == 'image':
        with open(file,'rb') as fp:
            msg = MIMEImage(fp.read(),_subtype=sub_type)
    elif main_type == 'audio':
        with open(file,'rb') as fp:
            msg = MIMEAudio(fp.read(),_subtype=sub_type)
    elif main_type == 'application':
        with open(file,'rb') as fp:
            msg = MIMEApplication(fp.read(),_subtype=sub_type)
    else:
        with open(file,'rb') as fp:
            msg = MIMEBase(main_type, sub_type)
            msg.set_payload(fp.read())
    filename = os.path.basename(file);
    msg.add_header('Content-Disposition','attachment',filename=filename)
    message.attach(msg);


def Email(destination, subject, body, files):
    service = authenticate()
    sendEmail(service, destination, subject, body, files)

    
    

