#!/usr/bin/env python3
# coding=utf-8


import getpass
import imaplib
import email
from email.header import decode_header, make_header
import mimetypes
import datetime
from dateutil.relativedelta import relativedelta
import re
import os

import config


FILENAME_RE = '[^\w \-@\[\]]'


def expand(thedirectory, msgs):
    global FILENAME_RE

    for part in msgs:
        if part.get_content_maintype() == 'multipart':
            continue

        if part.is_multipart():
            expand(thedirectory, part.get_payload(decode=False))
        else:
            filename = part.get_filename()
            if filename:
                filename = re.sub('\s+', ' ', re.sub(FILENAME_RE, ' ', filename))
            else:
                filename = 'content.txt'

            filenameorig = filename
            partfilename = os.path.join(thedirectory, filename)
            counter = 0
            while os.path.exists(partfilename):
                filename = ('%03d-' + filenameorig) % (counter)
                counter = counter + 1
                partfilename = os.path.join(thedirectory, filename)

            with open(partfilename, 'wb') as fp:
                payload = part.get_payload(decode=True)
                fp.write(payload)


def headerToString(theheader):
    myheader = email.header.Header(errors='ignore')
    decheader = email.header.decode_header(theheader)

    for d in decheader:
        myheader.append(d[0], charset=d[1], errors='ignore')

    return str(myheader)


def getMetadata(msg):
    global FILENAME_RE

    timestamp = 'unknown date'
    if 'Date' in msg:
        timestamp = str(datetime.datetime.fromtimestamp(email.utils.mktime_tz(email.utils.parsedate_tz(msg['Date'])))).replace(":", "-")
    thedirectory = timestamp;
    metadata = ' Date: ' + timestamp;

    if 'From' in msg:
        headerstr = headerToString(msg['From'])
        thedirectory = thedirectory + " " + re.sub(r'<.+?>', '', headerstr)
        metadata = metadata + "\n From: " + headerstr
    if 'To' in msg:
        headerstr = headerToString(msg['To'])
        thedirectory = thedirectory + " " + re.sub(r'<.+?>', '', headerstr)
        metadata = metadata + "\n To: " + headerstr
    if 'CC' in msg:
        headerstr = headerToString(msg['CC'])
        metadata = metadata + "\n CC: " + headerstr
    if 'BCC' in msg:
        headerstr = headerToString(msg['BCC'])
        metadata = metadata + "\n BCC: " + headerstr
    if 'Subject' in msg:
        headerstr = headerToString(msg['Subject'])
        thedirectory = thedirectory + " " + re.sub(r'<.+?>', '', headerstr)
        metadata = metadata + "\n Subject: " + headerstr

    thedirectory = thedirectory.replace('.', '-')
    thedirectory = re.sub('\s+', ' ', re.sub(FILENAME_RE, ' ', thedirectory))
    return {'directory': thedirectory, 'metadata': metadata}


def doStuff(dateString, mode):
    global FILENAME_RE

    for account in config.accounts:
        print("\naccount: " + account["name"])
        if not os.path.exists(account["name"]):
            os.mkdir(account["name"])

        metadatafile = None
        if mode == "download" or mode == "downloadexpand":
            metadatafile = open(os.path.join(account["name"], "Metadata.txt"), 'a')

        if 'ssl' in account and account['ssl']:
            mail = imaplib.IMAP4_SSL(account['server'], account['port'])
        else:
            mail = imaplib.IMAP4(account['server'], account['port'])
        #mail.enable("UTF8=ACCEPT")
        mail.login(account['user'], account['pwd'])

        typ,folders = mail.list()
        for folderb in folders:
            origfolder = folderb.decode('latin-1').replace(" \"/\" ", " \".\" ").split(" \".\" ")[1]
            folder = re.sub('\s+', ' ', re.sub(FILENAME_RE, ' ', origfolder))
            if 'ignorefolders' in account and folder in account['ignorefolders']:
                print("  (ignoring folder " + folder + ")")
                continue
            print("  processing folder " + folder)
            folderwithoutaccountname = folder
            folder = os.path.join(account["name"], folder)

            if not os.path.isdir(folder):
                os.mkdir(folder)

            mail.select(origfolder)
            typ, ids = mail.uid('search', None, dateString)
            if typ != "OK":
                raise "Could not search for messages"

            for id in ids[0].split():

                if mode == "download" or mode == "downloadexpand":
                    themailtyp, themail = mail.uid('fetch', id, '(RFC822)')
                    if themailtyp != "OK":
                        print("ERROR: " + id.decode('latin1'))

                    msg = email.message_from_string(themail[0][1].decode('latin-1'))

                    themetadata = getMetadata(msg)
                    thefilename = themetadata['directory'][0:config.MAX_FILENAME_LENGTH].strip()
                    thedirectory = os.path.join(folder, thefilename)

                    if metadatafile is not None:
                        metadatafile.write(os.path.join(folderwithoutaccountname, thefilename) + "\n" + themetadata['metadata'] + "\n\n")

                    thenumber = 2
                    theorigdirectory = thedirectory

                    if mode == "download":
                        thedirectory = thedirectory + ".eml"
                        while os.path.isdir(thedirectory):
                            thedirectory = theorigdirectory + " (" + str(thenumber) + ").eml"
                            thenumber += 1
                    elif mode == "downloadexpand":
                        while os.path.isdir(thedirectory):
                            thedirectory = theorigdirectory + " (" + str(thenumber) + ")"
                            thenumber += 1

                    if mode == "downloadexpand":
                        os.mkdir(thedirectory)

                        with open(os.path.join(thedirectory, "original-message.eml"), 'wb') as fp:
                            fp.write(themail[0][1])

                        expand(thedirectory, msg.walk())
                    elif mode == "download":
                        with open(thedirectory, 'wb') as fp:
                            fp.write(themail[0][1])

                elif mode == "delete":
                    #print("deleting " + id.decode('latin1'))
                    mail.uid('STORE', id, '+FLAGS', '(\\Deleted)')

        #if mode == "delete":
        #    mail.expunge()

        mail.logout()
        if metadatafile is not None:
            metadatafile.close()


def getDateStringLastMonths(monthdiff=12, numberofmonths=12):
    monthListRfc2822 = ['0', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                        'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    deltafrom = relativedelta(months=monthdiff)
    deltato = relativedelta(months=numberofmonths)

    sinceDate = datetime.datetime.today().replace(day=1) - deltafrom
    beforeDate = sinceDate + deltato
    result = ("(SINCE %s-%s-%s BEFORE %s-%s-%s)" % (sinceDate.strftime('%d'), monthListRfc2822[sinceDate.month], sinceDate.strftime('%Y'), beforeDate.strftime('%d'), monthListRfc2822[beforeDate.month], beforeDate.strftime('%Y')))
    return result

def getDateStringYear(year):
    return "(SINCE 01-Jan-" + year + " BEFORE 01-Jan-" + str(int(year) + 1) + ")"

if __name__ == "__main__":
    year = input("Download messages for year: ")
    dateString = getDateStringYear(year)
    print("Downloading mails between " + dateString)

    theexpand = input("Expand messages (y/N)? ")
    doexpand = theexpand == "y" or theexpand == "Y"

    for account in config.accounts:
        if 'pwd' not in account:
            account['pwd'] = getpass.getpass("password for " + account['name'] + ": ")

    if doexpand:
        doStuff(dateString, "downloadexpand")
    else:
        doStuff(dateString, "download")

    theinput = input("Downloading finished. Delete messages from server (y/N)? ")
    if theinput == "y" or theinput == "Y":
        print("deleting ...")
        doStuff(dateString, "delete")
        print("done.")

