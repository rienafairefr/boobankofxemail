import smtplib
import shutil
import subprocess
import time
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders
import os
import urllib
from time import strftime

from contextlib import contextmanager
from weboob.applications import boobank
import configparser
import argparse

parser = argparse.ArgumentParser(description='Automatic Download ofx from winancial.')

parser.add_argument('--nosend', dest='nosend', action='store_const',
                    const=True, default=False,
                    help='activate so that email is not sent out')
parser.add_argument('-numberofdays', metavar='days', type=int, default=28,
                    help='Number of days to look for data in the past')
parser.add_argument('--account', metavar='AccountName', type=str, default='',
                    help='Account to treat')

inputargs = parser.parse_args()

cp = configparser.ConfigParser()
cp.read('config.conf')

smtp_server = cp.get('CONFIG', 'SMTP_SERVER')
smtp_port = cp.getint('CONFIG', 'SMTP_PORT')
smtp_username = cp.get('CONFIG', 'SMTP_USERNAME')
smtp_password = cp.get('CONFIG', 'SMTP_PASSWORD')
smtp_emailfrom = cp.get('CONFIG', 'SMTP_EMAILFROM')
smtp_emailto = cp.get('CONFIG', 'SMTP_EMAILTO').split(',')

import re
from unicodedata import normalize

_punct_re = re.compile(r'[\t !"#$%&\'()*\-/<=>?@\[\\\]^_`{|},.:]+')

def slugify(text, delim=u'-'):
    """Generates an slightly worse ASCII-only slug."""
    result = []
    for word in _punct_re.split(text.lower()):
        word = normalize('NFKD', word).encode('ascii', 'ignore')
        if word:
            result.append(word)
    return unicode(delim.join(result))

def send_mail(send_from, send_to, subject, text, files=[], server="localhost", port=587, username='', password='',
              isTls=True):
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(text))

    for f in files:
        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(f, 'rb').read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="{0}"'.format(f))
        msg.attach(part)

    smtp = smtplib.SMTP(server, port)
    smtp.set_debuglevel(True)
    if isTls: smtp.starttls()
    smtp.login(username, password)
    if isinstance(send_to,list):
        smtp.sendmail(send_from, send_to, msg.as_string())
    else:
        smtp.sendmail(send_from, [send_to], msg.as_string())
    smtp.quit()


from datetime import timedelta, datetime

enddate = datetime.today() + timedelta(days=7)
startdate = datetime.today() - timedelta(days=inputargs.numberofdays)


Accounts=dict()
for (each_key, each_val) in cp.items('ACCOUNTS'):
    Accounts[each_key]=each_val

retrieveds = []

datafile='data.csv'
if os.path.exists(datafile):
    os.remove(datafile)
cmd='boobank list -f csv > %s'%datafile
BBank=boobank.Boobank()
args=['','list','-f','csv','outfile','data.csv']
BBank.parse_args(['','--debug'])
BBank.load_config()
BBank.load_backends()

data={}
for account in BBank.do('iter_accounts'):
    data[account.fullid] = account

def is_non_zero_file(fpath):
    return os.path.isfile(fpath) and os.path.getsize(fpath) > 0

for id, v in Accounts.items():
    if (inputargs.account!=''):
        if id!=inputargs.account:
            continue

    name= id + '.ofx'

    cmd=['boobank','history',v,startdate.strftime('%Y-%m-%d'),'-f','ofx']

    if os.path.exists(name):
        os.remove(name)

    p = subprocess.Popen(cmd, stdout=subprocess.PIPE,stderr=subprocess.PIPE)
    result = p.stdout.read()
    err = p.stderr.read()
    time.sleep(0.5)
    with open(name,'w') as outputfile:
        outputfile.write(result)
    with open('errfile', 'w') as errfile:
        errfile.write(err)
    if os.path.exists(name) and is_non_zero_file(name):
        newname=name+'_'+str(data[v].balance)+'.ofx'
        shutil.move(name,newname)
        retrieveds.append(newname)
        print('OK')
    else:

        send_mail(subject="Boobank to OFX Email Problem fetching %s" % datetime.today().strftime('%Y-%m-%d'), text='Error fetching the account %s %s' % (id, v),
                  server=smtp_server, port=smtp_port,
                  files=[os.path.abspath('resultfile')],
                  send_from=smtp_emailfrom, send_to=smtp_emailto,
                  username=smtp_username, password=smtp_password
                  )
        exit(-1)


if not inputargs.nosend:
    print("sending email...")
    send_mail(subject="Budget OFX %s" % datetime.today().strftime('%Y-%m-%d'), text='Downloaded files',
              files=retrieveds,
              server=smtp_server, port=smtp_port,
              send_from=smtp_emailfrom, send_to=smtp_emailto,
              username=smtp_username, password=smtp_password
    )
    print("done sending email")

for f in retrieveds:
    os.remove(f)

