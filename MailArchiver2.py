#!/usr/bin/env python3
# coding=utf-8


import getpass
import imaplib
from imapclient import imap_utf7
import email
from email.header import decode_header, make_header
import mimetypes
import datetime
from dateutil.relativedelta import relativedelta
import re
import os

import config

list_response_pattern = re.compile(
    r'\((?P<flags>.*?)\) "(?P<delimiter>.*)" (?P<name>.*)'
)


def parse_list_response(line, useUTF7=False):
    if useUTF7:
        decodedstring = imap_utf7.decode(line)
    else:
        decodedstring = line.decode('latin-1')

    match = list_response_pattern.match(decodedstring)
    flags, delimiter, mailbox_name = match.groups()
    mailbox_name = mailbox_name.strip('"')
    return (flags, delimiter, mailbox_name)


def getFilenameExtension(thefilename, maxExtensionLength, defaultExtension=""):
    pos = thefilename.rfind(".")
    if pos == -1 or pos == len(thefilename) - 1 or len(thefilename) - pos > maxExtensionLength:
        return (thefilename, defaultExtension)

    return (thefilename[0:pos], thefilename[pos:])


def getFilename(thepath, thefilename, maxFilenameLength, maxExtensionLength, defaultExtension, counterDelimiter="-"):
    if maxExtensionLength > 0:
        filenameSplit = getFilenameExtension(thefilename=thefilename, maxExtensionLength=maxExtensionLength, defaultExtension=defaultExtension)
        rawFilename = filenameSplit[0]
        theextension = filenameSplit[1]
    else:
        rawFilename = thefilename
        theextension = defaultExtension

    rawFilename = re.sub('\s+', ' ', re.sub(config.FILENAME_RE, '', rawFilename)).strip()
    theextension = re.sub('\s+', ' ', re.sub(config.FILENAME_RE, '', theextension)).rstrip()

    maxLength = max(0, maxFilenameLength - len(theextension))
    newFilename = rawFilename[0:maxLength].rstrip() + theextension

    counter = 2
    while os.path.exists(os.path.join(thepath, newFilename)):
        fullpostfix = str(counter) + theextension
        maxLength = maxFilenameLength - len(counterDelimiter + fullpostfix)
        if maxLength > 0:
            newFilename = rawFilename[0:maxLength].rstrip() + counterDelimiter + fullpostfix
        else:
            newFilename = fullpostfix
        counter = counter + 1

    return os.path.join(thepath, newFilename)


def getFoldername(thepath, thefoldername, maxFoldernameLength, counterDelimiter="-"):
    return getFilename(thepath=thepath, thefilename=thefoldername, maxFilenameLength=maxFoldernameLength, maxExtensionLength=0, defaultExtension="", counterDelimiter=counterDelimiter)


def expand(thedirectory, msgs):
    
    for part in msgs:
        if part.get_content_maintype() == 'multipart':
            continue

        if part.is_multipart():
            expand(thedirectory, part.get_payload(decode=False))
        else:
            thefilename = part.get_filename()
            if not thefilename:
                thefilename = 'content' + config.DEFAULT_EXTENSION

            partfilename = getFilename(thepath=thedirectory, thefilename=thefilename, maxFilenameLength=config.MAX_FILENAME_LENGTH, maxExtensionLength=config.MAX_FILENAME_EXTENSION_LENGTH, defaultExtension=config.DEFAULT_EXTENSION)

            try:
                with open(partfilename, 'wb') as fp:
                    payload = part.get_payload(decode=True)
                    fp.write(payload)
            except FileNotFoundError:
                print("ERROR: could not open file " + partfilename)


def headerToString(theheader):
    myheader = email.header.Header(errors='ignore')
    decheader = email.header.decode_header(theheader)

    for d in decheader:
        myheader.append(d[0], charset=d[1], errors='ignore')

    return str(myheader).strip()


def getMetadata(msg):
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
        headerstr = headerToString(msg['To']).replace("\n", " ")
        thedirectory = thedirectory + " " + re.sub(r'<.+?>', '', headerstr)
        metadata = metadata + "\n To: " + headerstr
    if 'CC' in msg:
        headerstr = headerToString(msg['CC']).replace("\n", " ")
        metadata = metadata + "\n CC: " + headerstr
    if 'BCC' in msg:
        headerstr = headerToString(msg['BCC']).replace("\n", " ")
        metadata = metadata + "\n BCC: " + headerstr
    if 'Subject' in msg:
        headerstr = headerToString(msg['Subject'])
        thedirectory = thedirectory + " " + re.sub(r'<.+?>', '', headerstr)
        metadata = metadata + "\n Subject: " + headerstr

    thedirectory = thedirectory.strip()
    return {'directory': thedirectory, 'metadata': metadata}


def doStuff(mainFolder, dateString, mode):
    if not os.path.exists(mainFolder):
        os.mkdir(mainFolder)

    for account in config.accounts:
        accountdir = getFoldername(thepath=mainFolder, thefoldername=account["name"], maxFoldernameLength=config.MAX_FILENAME_LENGTH)
        os.mkdir(accountdir)
        print("\naccount: " + account["name"])

        metadatafile = None
        if mode == "download" or mode == "downloadexpand":
            metadatafile = open(os.path.join(accountdir, "Metadata.txt"), 'a', encoding='utf-8')

        if 'ssl' in account and account['ssl']:
            mail = imaplib.IMAP4_SSL(account['server'], account['port'])
        else:
            mail = imaplib.IMAP4(account['server'], account['port'])
            if 'starttls' in account and account['starttls']:
                mail.starttls()
        #mail.enable("UTF8=ACCEPT")
        mail.login(account['user'], account['pwd'])

        typ,folders = mail.list()
        for folderb in folders:
            flags, delimiter, originalfolder = parse_list_response(folderb, True)
            _, _, originalfolderunencoded = parse_list_response(folderb, False)

            if 'ignorefolders' in account and originalfolder in account['ignorefolders']:
                print("  (ignoring folder " + originalfolder + ")")
                continue
            print("  processing folder " + originalfolder)
            #continue

            folder = getFoldername(thepath=accountdir, thefoldername=originalfolder, maxFoldernameLength=config.MAX_FILENAME_LENGTH)
            os.mkdir(folder)

            mail.select('"{}"'.format(originalfolderunencoded))
            typ, ids = mail.uid('search', None, dateString)
            if typ != "OK":
                raise "Could not search for messages"

            for id in ids[0].split():
                if mode == "download" or mode == "downloadexpand":
                    themailtyp, themail = mail.uid('fetch', id, '(RFC822)')
                    if themailtyp != "OK":
                        print("ERROR: " + id.decode('latin1'))

                    if themail is None or themail[0] is None:
                        print("WARNING: ignoring empty mail " + str(id))
                        continue
                    msg = email.message_from_string(themail[0][1].decode('latin-1'))

                    themetadata = getMetadata(msg)
                    thefilename = themetadata['directory']

                    if mode == "download":
                        thedirectory = getFilename(thepath=folder, thefilename=thefilename, maxFilenameLength=config.MAX_FILENAME_LENGTH, maxExtensionLength=config.MAX_FILENAME_EXTENSION_LENGTH, defaultExtension=".eml")
                        
                        with open(thedirectory, 'wb') as fp:
                            fp.write(themail[0][1])
                    elif mode == "downloadexpand":
                        thedirectory = getFoldername(thepath=folder, thefoldername=thefilename, maxFoldernameLength=config.MAX_FILENAME_LENGTH)
                        os.mkdir(thedirectory)

                        msgfilename = getFilename(thepath=thedirectory, thefilename="original-message.eml", maxFilenameLength=config.MAX_FILENAME_LENGTH, maxExtensionLength=config.MAX_FILENAME_EXTENSION_LENGTH, defaultExtension=".eml")
                        with open(msgfilename, 'wb') as fp:
                            fp.write(themail[0][1])

                        expand(thedirectory, msg.walk())

                    if metadatafile is not None:
                        metadatafile.write(thedirectory[len(accountdir + os.sep):] + "\n" + themetadata['metadata'] + "\n\n")


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

    theexpand = input("Expand messages (Y/n)? ")
    doexpand = True
    if theexpand == "n" or theexpand == "N":
        doexpand = False

    for account in config.accounts:
        if 'pwd' not in account:
            account['pwd'] = getpass.getpass("password for " + account['name'] + ": ")

    yearfolder = os.path.join(config.WORKING_DIR, year)

    if doexpand:
        doStuff(yearfolder, dateString, "downloadexpand")
    else:
        doStuff(yearfolder, dateString, "download")

    theinput = input("Downloading finished. Delete messages from server (y/N)? ")
    if theinput == "y" or theinput == "Y":
        print("deleting ...")
        doStuff(yearfolder, dateString, "delete")
        print("done.")

