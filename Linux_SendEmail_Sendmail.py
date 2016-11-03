#!/apps/python/Python-2.7.8/bin

#bring in all the required modules to be able to send emails using sendmail on Linux
from smtplib import SMTP
from itertools import chain
from errno import ECONNREFUSED
from mimetypes import guess_type
from subprocess import Popen, PIPE
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from socket import error as SocketError
from email.encoders import encode_base64
from email.mime.multipart import MIMEMultipart
from os.path import abspath, basename, expanduser
import os

#function to determine the MIME type of a passed file; input a valid file path; returns the files MIME type in a tuple
def get_mimetype(filename):
    content_type, encoding = guess_type(filename)
    if content_type is None or encoding is not None:
        content_type = "application/octet-stream"
    return content_type.split("/",1)

#function to convert a passed file to proper MIME format for attachment to an e-mail; input a valid file path; returns a MIME object for the provided file
def mimify_file(filename):
    filename = abspath(expanduser(filename))
    basefilename = basename(filename)

    msg = MIMEBase(*get_mimetype(filename))
    msg.set_payload(open(filename,"rb").read())
    msg.add_header("Content-Disposition","attachment",filename=basefilename)

    encode_base64(msg)

    return msg

#function to send an email using the passed to, subject, and body text (optional params are cc, bcc, files, and sender)
def send_email(to, subject, text, **params):
    #parameters
    cc = params.get("cc", [])
    bcc = params.get("bcc", [])
    files = params.get("files", [])
    sender = params.get("sender","xxxx")

    #build the recipient string based on passed inputs
    recipients = list(chain(to, cc, bcc))

    #build the email
    msg = MIMEMultipart()
    msg.preamble = subject
    msg.add_header("From", sender)
    msg.add_header("Subject",subject)
    msg.add_header("To", ", ".join(to))
    cc and msg.add_header("Cc",", ".join(cc))

    msg.attach(MIMEText(text))

    [msg.attach(mimify_file(filename)) for filename in files]

    p = os.popen("/usr/sbin/sendmail -t", "w")
    p.write(msg.as_string())
    status = p.close()
    if status != 0:
        print "Sendmail exit status", status

send_email(["xxxx"], "Test", "Hers is the link you wanted:\nhttp://www.python.org", sender="xxxx", files=["/export/home/curtisc/devPythonScripts/Test.py"])

    

    
