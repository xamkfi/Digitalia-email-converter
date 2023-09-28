#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Apr 14 13:27:03 2016

@author: digitalia-aj


SLS
-Only emails and calendars, 
-Attachments, keep the original also
-Runtime files --> delete
-Reportin once a month

"""
#from convertfiles_tester import pdfA
#from statsmodels.tsa.vector_ar.tests.test_var import basepath
#from binstar_client.utils.tables import TEMPLATE

"""Python imports"""
import folderSeeker
import os
#import chardet
#import tarfile
#import zipfile
import subprocess
import shutil
import re
import string
import multiprocessing
from multiprocessing import Pool
import time
import sys
#import uuid
import random
import operator
import atexit
import mimetypes
import mailbox
import email
import quopri
#import base64
#import codecs
import convertMetaFile
import tarfile
#from msg_parser import MsOxMessage #Used to convert .msg to .eml file
#from selenium import webdriver
#from selenium.webdriver.firefox.options import Options
from bs4 import BeautifulSoup

from pdfrw import PdfReader, PdfWriter, PdfDict

from pathvalidate import sanitize_filename

from dateutil.parser import parse
#import unicodedata
#import chardet
#import eml_parser
from email.header import decode_header, make_header

#development = True #For api use change to false
"""
*************************************************************
Variables and paths, select according to environment
*************************************************************
"""
#if development:
"""own .py imports"""
import logwriter
import cleaner
import convertfiles
droid_path ='/home/software/droid/droid.sh'
droid_signature_path='/home/software/droid/DROID_SignatureFile_V112.xml'
droid_container_path='/home/software/droid/container-signature-20230510.xml'
pdfa_def_path = '/home/digitalia-aj/extraspace/eclipse-workspace/Email-converter/pdfa_def.ps' #content needs to be modified if version changes
#ghostscript_path = '/snap/bin/gs' #this version causes error
ghostscript_path = 'gs' #Uses the ubuntu default ghostscript 10.0.1 causes unrecoverable error

verapdf_path = '/home/software/verapdf/verapdf'

# else: #This is used for dockerized env   
#     from . import logwriter
#     from . import cleaner
#     from . import convertfiles
#     home_path='/app'
#     droid_path = os.path.join(home_path, 'droid/droid.sh')
#     #droid_signature_path=os.path.join(home_path, 'droid/DROID_SignatureFile_V93.xml')
#     #droid_container_path=os.path.join(home_path, 'droid/container-signature-20171130.xml')
#     ghostscript_path = 'gs'
#     verapdf_path = os.path.join(home_path, 'verapdf/verapdf')


leaveAttachmentFolder = True #One of these must be true if attachments are kept. Demomode sets both to false
#embedfiles = False #Embed attachmets into PDF/A-3 file or not

#Some default values
convertAttachments = True # If the attachments are converted or not
attachmentToPDFA = True #If true, pdf conversion  for attachments will create PDF/A compliant files
validate = True #Should the results be validated, mainly for testing purposes

"""By default handles both below == both True"""
handleMeetings = False # If true, process the meetings
handleMessages = True # If true, process the messages

#File naming convention
useTopicsAsFilenames = False # It se to true, pdf files will be named by topic of email
addTimetoFilename = False # If set to true, received time will be added to the filename

#Commandline booleans
delflag = True #deletes all unneeded runtime files
deleteOriginalAttachment = False
dirflag = False #Get dirname from the cli
fileflag = False #Gets filename from the cli
zipflag = False #Tar.gz produced content if true
attachmentflag = False #Attachments are attached to pdf/a files if true

calculateMails = False # Used in independent website testing environment, not required in cli environment but can be kept True
finalCleanupDone = False #Just to make sure tar.gz won't happen before cleanup

onGoingPoolCounter = 0
completedConversion = 0
totalNumberofEmails = 0
countedPaths =[]
cleanTheseDirs = []
firstCall = True
exportPath = ""

htmlBegin = """<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
div.gallery {
  margin: 5px;
  border: 1px solid #ccc;  
  width: 180px;
}

div.gallery:hover {
  border: 1px solid #777;
}

div.gallery img {
  width: 100%;
  height: auto;
}

div.desc {
  padding: 15px;
  text-align: center;
}
</style>
</head>
"""

htmlEnd = """
</body>
</html>"""

    
#Not to be checked filetypes list
extensions = set(['.py','.jar','.log','.csv', '.txt'])

alphanum = string.ascii_lowercase + string.digits

def doCmd(cmd):
    cmdProcess = subprocess.Popen(cmd, stdout=subprocess.PIPE)
    result = cmdProcess.communicate()[0] #0 is the actual result, 1 would be the error    
    return result


"""
Writes mbox and eml mail headers into a file
"""
def printEmailHeaders(header, headerfilewriter):
    for s in header:
        key = str(s[0])
        value = str(s[1])                
        if '=?' in value:
            value = make_header(decode_header(value))
            #decodedvalue = decode_header(value)
        headerfilewriter.write(key+": "+str(value)+"\n")
        #logwriter.printmessage(key+":"+str(value))
        #Headers must be written into CombinedHeaders.txt
        headerfilewriter.close    
    return

"""
Used in case of eml file
"""
def emlManipulate(exportPath):
    dircontent = os.listdir(exportPath)
    #logwriter.printmessage(dircontent)
    for onefile in dircontent:
        #logwriter.printmessage("{}".format(dircontent))
        newfile = os.path.join(exportPath,str(onefile))
        #logwriter.printmessage("Now the path is {}".format(newfile))
        if os.path.isfile(newfile):        
            #rawdata = open(newfile, 'rb').read()
            #encoding = chardet.detect(rawdata)['encoding']   
            #logwriter.printmessage("Encoding = {}".format(encoding))
            #filedata = filereader.read()
            #logwriter.printmessage("{}".format(filedata))
            #message = email.message_from_file(rawdata, 'rb', encoding)
            message = email.message_from_binary_file(open(newfile, "rb"))
            #message = email.message_from_binary_file(codecs.open(newfile, "rb"))
            headerfilewriter = open(os.path.join(exportPath, "InternetHeaders.txt"), 'w')
            #logwriter.printmessage("MMM: "+str(message))
            header = message._headers            
            printEmailHeaders(header, headerfilewriter)
            #logwriter.printmessage("test {}".format(onefile))
            walkEmailContent(message, exportPath, onefile)            
    return

"""
Used in case of mbox, creates folder structure and moves files
"""
def mboxManipulate(exportPath):
    logwriter.printmessage("Mboxmanipulate")
    dircontent = os.listdir(exportPath)
        #logwriter.printmessage(files)
    for onefile in dircontent:
        logwriter.printmessage("PATH! - {}".format((os.path.join(exportPath,str(onefile)))))
        os.makedirs(os.path.join(exportPath,"Message"+str(onefile)))
        newfile = exportPath+"/Message"+str(onefile)+"/Message"
        cleaner.renameFile(os.path.join(exportPath,onefile), newfile ) 
        print(sys.stdin.encoding)
        if os.path.isfile(newfile):
            curpath = os.path.dirname(os.path.abspath(newfile))
            logwriter.printmessage("Newly create file found in %s"%curpath)
            #Lets write InternetHeaders.txt file
            headerfilewriter = open(os.path.join(curpath, "InternetHeaders.txt"), 'w')
           
            email = mailbox.mbox(newfile)
            #logwriter.printmessage("Email content = {}".format(email))
            message = email[0] 
            #logwriter.printmessage("Message = {}".format(message))
            header = message._headers
            printEmailHeaders(header, headerfilewriter)
            #logwriter.printmessage("Headers for file {} are {}".format(curpath, header))
            walkEmailContent(message, curpath, newfile)
    
    return

"""
Tries to fix badly formatted filenames
"""
def fixInvalidFileNames(filename, attachpath=""):
    """This part tries identify filetypes for files without endings"""
    if attachpath !="":
        try:           
            fileend = filename.rsplit('.',1)[1]
        except:        
            fullpath = os.path.join(attachpath,filename)
            logwriter.printmessage("{} no ending, trying to resolve".format(fullpath))
            findcmd = ['file','--extension', fullpath]
            result = subprocess.run(findcmd, capture_output=True, text=True, universal_newlines=True)
            returntype = result.stdout.split(":")[1].strip()
            logwriter.printmessage(returntype)
            if '/' in returntype:
                #logwriter.printmessage("splitting")
                returntype = returntype.split('/')[0]
                #logwriter.printmessage("after split {}".format(returntype))
                
            if '?' in returntype:
                returntype = "unknown"
            filename = filename+"."+returntype
            logwriter.printmessage("Filename now {}".format(os.path.join(attachpath,filename)))
    #filename = str(filename).lower()   
    #filename = re.sub(r'(\\x.{2})', '-', filename) #Reqular expressions replaces all after and with . eg. .doc, .docx etc
    
    fixedFilename = sanitize_filename(filename)
    #logwriter.printmessage("{} vs. {}".format(filename, fixedFilename))

    """if '%20' in filename:        
        #logwriter.printmessage("BADBAD väli {}".format(filename))
        filename = filename.replace('%20','-')
    if ':' in filename:
        #logwriter.printmessage("BADBAD- {}".format(filename))
        filename = filename.replace(':', '-')
    if '*' in filename:
        #logwriter.printmessage("BADBAD- {}".format(filename))
        filename = filename.replace('*', '-')
    """
    """        
    lastchars = filename[-5:]
    if '.' not in lastchars:
        print("Not a correct file: {}".format(filename))
        filename = filename+".unknown"             
    """
    
    return fixedFilename

"""
Actual email content separator used by mbox and eml manipulate
"""
def walkEmailContent(message, curpath, openedFile):
    partcounter = 1
    #logwriter.printmessage("path = {}, file = {}".format(curpath, openedFile))
    for part in message.walk():              
    
        filename=""
        #logwriter.printmessage(part.get_filename())
        payload = part.get_payload(decode=True)        
        #payload = part.get_payload()
        partcontenttype = part.get_content_type()            
        partcharset = part.get_content_charset('iso-8859-1')        
        #logwriter.printmessage("Part content type = {}, part charset = {}".format(partcontenttype, partcharset))
        try:
            partencoding = part['Content-Transfer-Encoding']
            #logwriter.printmessage("part encoding is {}".format(partencoding))
            if partencoding is None:
                partencoding = "None"
                #logwriter.printmessage("part encoding changed to text '{}'".format(partencoding))
        except Exception as e:
            #logwriter.printmessage("Content transfer encoding not found: {}, using none as a value".format(e.__cause__))
            partencoding = "none"
        #decoded_payload = ""
        #logwriter.printmessage("content part = "+str(partcounter)+" "+curpath+" - "+str(partcharset)+" : - "+str(partcontenttype))
        #logwriter.printmessage("File {}, part {} encoding = {}".format(openedFile, partcounter, partencoding))
        #if not 'multipart' in str(partcontenttype):     
                        
        #payload = part.get_payload()                      
        
        #defects = part.defects
        #logwriter.printmessage(str(defects))
        
        if partcontenttype =='text/plain':                        
            filename = "Message.txt"            
            contentwrite = open(os.path.join(curpath,filename), 'wb')
            contentwrite.write(payload)          
            contentwrite.close         
                                                       
        if partcontenttype == 'application/rtf':
            #payload = part.get_payload(decode=True)
            filename = "Message.rtf"
            #decoded_payload = payload.decode(partcharset)                
            contentwrite = open(os.path.join(curpath,filename), 'wb')
            #logwriter.printmessage(str(partcounter)+" TESTI: %s"%decoded_payload)
        
            contentwrite.write(payload)
            contentwrite.close        
            
        if partcontenttype == 'text/html':
            filename = "Message.html"            
            if 'quoted' in partencoding or '8bit' in partencoding:              
                #logwriter.printmessage("Quoted printable or 8bit lets decode")
                #payload = quopri.decodestring(str(payload))
                payload = quopri.decodestring(payload)
                #payload = payload.replace(b'iso-8859-1', b'utf8')                
                payload = payload.replace(b'utf-8', b'utf8')                
                #logwriter.printmessage(str(payload))
                badchars =["x83"]
                if any (bad in str(payload) for bad in badchars):
                    logwriter.printmessage("Unwanted character found")
                    #logwriter.printmessage("Bad html characters in {}, using rtf".format(os.path.join(curpath, openedFile)))
                    #html contains bad encoding, do not write it
                else:
                    #payload = payload.replace('iso-8859-1', 'utf8')
                    
                    #logwriter.printmessage("html message = {}".format(payload))
                    #logwriter.printmessage("Writing message {}".format(filename))
                    contentwrite = open(os.path.join(curpath,filename), 'wb')
                    contentwrite.write(payload)
            
            
            else:    
                if 'None' in str(partcharset):               
                    logwriter.printmessage("No charset found in html --> lets force use of rtf or txt")
                    #So no actions needed
                else:
                    #payload = payload.decode(partcharset)
                    contentwrite = open(os.path.join(curpath,filename), 'wb')
                    contentwrite.write(payload)
            
            contentwrite.close
        
        
        if part.get_filename(): #If part has a filename it is an attachment, unnamed attachments will be ignored                            
            """
            Content-ID: <16161327803.B5ac8F4.77767@email-manipulator>
            Content-Type: application/octet-stream; name="Kesärekryt 2018.pptx"
            Content-Transfer-Encoding: base64
            Content-Disposition: attachment; filename="Kesärekryt 2018.pptx"            
            
            Content-ID: <16166629195.2C3274.72507@email-manipulator>
            Content-Type: message/rfc822
            Content-Transfer-Encoding: 8bit
            Content-Disposition: attachment; filename="VS: Mustan laatikon jatkosta -
            toimittajabrändäystä ja supertimantteja?"            
            """
            
            
            temppath = os.path.join(curpath, "Attachments")
            if not os.path.isdir(temppath):
                os.makedirs(os.path.join(curpath, temppath))
            
            if partcontenttype == 'message/rfc822':                
                templist = email.header.decode_header(part['Content-Disposition'])
                #logwriter.printmessage(templist)
                try:
                    tempname = templist[0][0].decode('utf-8').split('=')[1].replace('"','') #Seems to have only one element, tuple. first value is the name
                except (AttributeError):
                    tempname = templist[0][0].split('=')[1].replace('"','')
                tempname = fixInvalidFileNames(tempname.replace("\n\r", "").replace("\r", "").replace('\n', '').replace('/',''))
                #logwriter.printmessage(tempname)
                part = quopri.decodestring(part.as_bytes())   
                
                contentwrite = open(os.path.join(temppath, tempname+".msg"), 'wb')
                contentwrite.write(part)
            else:
                #cid = part['Content-ID']
                #logwriter.printmessage(part['Content-Type'])
                #logwriter.printmessage(part['Content-Type'].decodedvalue)

                """
                filename = part.get_filename()              
                logwriter.printmessage("Filename before {}".format(filename))
                filename = fixInvalidFileNames(filename.replace("\n", "").replace("\r", "")).replace('/','').replace(':','-')
                #filename = fixInvalidFileNames(part.get_filename())
                #filename = quopri.encodestring(bytes(filename),'utf-8').decode(encoding='utf-8')
                logwriter.printmessage("Attachment name after decode {}".format(repr(filename)))
                """
                
                decct = decode_header(part['Content-Type'])[0][0] #first of list and first of tuple and split the results
                #logwriter.printmessage("now {} - {}".format(decct, type(decct)))
                if(type(decct)==str):
                    filename = part.get_filename() #Gets the filename but the encoding is wrong, usully
                    filename = fixInvalidFileNames(filename.replace("\n\r", "").replace("\r", "").replace("\n", "").replace("/","").replace('%0A', ''))
                #application/octet-stream; name="Kesärekryt 2018.pptx"
                else:
                    decct = decct.decode('utf-8')
                    filename = decct.split('"')[1]                    
                    filename = fixInvalidFileNames(filename.replace("\n\r", "").replace("\r", "").replace("\n", "").replace("/","").replace('%0A', ''))
                #logwriter.printmessage("after {}".format(name))
                
                           
                #logwriter.printmessage("Attachment name : {}".format(filename))
                
                binarywriter = open(os.path.join(temppath,str(filename)), 'wb')
                #logwriter.printmessage("After: %s"%filename)
                #logwriter.printmessage("Payload lenght = {}".format(len(payload)))
                binarywriter.write(payload)
                binarywriter.close
                #decoded_payload = payload
            
        
        partcounter+=1
        #contentwrite.close
    return


def calculateExtractedFolders(modified_walk_dir):
    #global calculateMails
    global totalNumberofEmails
    global countedPaths
        
    #count1 = 0     
    try:
        for root, dirs, files in os.walk(modified_walk_dir):            
            for dirname in dirs:                               
                #part1, part2 = os.path.split(root)
                #logwriter.printmessage(str(root)+"-->"+str(part1)+" -vs.- "+str(countedPaths))
                #logwriter.printmessage(root+ " "+filename +str(os.path.getsize(os.path.join(root, filename))))                    
                #if os.path.getsize(os.path.join(root, filename))>0: #Additional check that 0 sized files are rejected
                if "Message" in dirname and root not in countedPaths and "Attachments/Attachment" not in root and "Poistetut" not in root and "Deleted" not in root:                                        
                    countedPaths.append(root) #Adds walked path so it won't be accessed again                   
                    #logwriter.printmessage(os.listdir(root))
                    for tempdirs in os.listdir(root):
                        if "Message" in tempdirs:
                            totalNumberofEmails += 1                                      
                    #logwriter.printmessage(os.listdir(part1))
                    #logwriter.printmessage("email-count:"+str(count1)) 
                    #calculateMails = False             
                #logwriter.printmessage(count1)
        logwriter.printmessage("email-count:"+str(totalNumberofEmails))
               
    except OSError:
        pass
    
    return

def convertAttachmentNoEmbed(root):
    #logwriter.printmessage("convert, don't embed {}".format(root))
    attachpath = os.path.join(root, "Attachments")
    dircontent = os.listdir(attachpath)
    #logwriter.printmessage(dircontent)
    for x in range(len(dircontent)):
        fileInDir = dircontent[x]        
        if 'Attachment00' in fileInDir:
            #logwriter.printmessage("ATTACHMENT in ATTACHMENT {}".format(root))                           
            zdir = os.path.join(attachpath, fileInDir)
            tfile = zdir+".zip"
            #logwriter.printmessage(zdir)            
            shutil.make_archive(zdir, 'zip', zdir)
            """
            zipf = zipfile.ZipFile(tfile, mode='w')
            #tarf = tarfile.open(tfile, "w")
            zipf.add(zdir, arcname="")
            zipf.close()
            """
            if os.path.isfile(tfile):
                cleaner.removeDir(zdir)
        else:     
            #filepathBeforeFix = os.path.join(attachpath, fileInDir)
            fileAfterFix= fixInvalidFileNames(fileInDir, attachpath)
            orgFile = os.path.join(attachpath,fileInDir)
            renFile = os.path.join(attachpath,fileAfterFix)
            if str(orgFile) != str(renFile):
                #logwriter.printmessage("File name changed during the fix, renaming file : {}".format(filepathBeforeFix))
                cleaner.renameFile(orgFile, renFile)
                
        try:
            if os.path.isfile(renFile):
                #if convertAttachments:
                convertfiles.outsideLandingPlace(renFile, attachmentToPDFA, deleteOriginalAttachment)
                    
        except:
            logwriter.printmessage("Something strange in {}".format(root))
            pass
                        
    return


#AJ mod 22.9 --> attachments converter added


"""this is the actual
 function that does the conversion from
message to pdf
"""
def msgPDFConversion(root, filename):
        
    try:                              
        file_path = os.path.join(root, filename) #Full path to the received file
        #logwriter.printmessage(file_path)
        #tempFile_path = os.path.join(root, filename)+".htm"
                
        if os.path.isdir(os.path.join(root, "Attachments")):            
            #logwriter.printmessage("Attachments folder found {}".format(root))
            convertAttachmentNoEmbed(root)
           
        
        head, temptail = os.path.split(root)
        uniqueID = ''.join(random.choices(alphanum, k=4))
        #logwriter.printmessage("UUID = {}".format(uniqueID))
        """
        If set at true, topic of email will become the name of the pdf file
        """        
        if useTopicsAsFilenames:        
            #/Subject (Kutsu: 'Tulevaisuuspolkuja' -Etelä-Savon tulevaisuusfoorumi 15.10.2015 Mikkelissä)
            if os.path.isfile(os.path.join(root, "ModifiedCombinedHeaders.txt")):
                logwriter.printmessage("Use topic as filename, extract from modifiedcombined headers")
                topicreader = open(os.path.join(root, "ModifiedCombinedHeaders.txt"), 'r')
                for line in topicreader:
                    if line.startswith("/Subject ("):
                        #logwriter.printmessage(line)
                        temptail = line.lstrip("/Subject (")
                        temptail = str(temptail).rstrip().rstrip(")")
                        logwriter.printmessage(temptail)
                        #In case of topics as filenames, add only first 8 chars of uuid to the filename --> shorter filename
                        stripedUniqueID = uniqueID.split('-')[0]
                        tail = temptail+"_"+stripedUniqueID+"_"
                topicreader.close()
        else:       
            if 'eml.export' in temptail:
                tail = uniqueID+"_"    
            else:
                tail = temptail+"_"+uniqueID+"_"    #Lets separate the uuid with _ _ marks
        #logwriter.printmessage ("MIDDLE2 "+str(filename))
        if addTimetoFilename:
            logwriter.printmessage("Adding time")
            datereader = open(os.path.join(root, "ModifiedCombinedHeaders.txt"), 'r')
            for line in datereader:
                if line.startswith("/Date ("):
                    #logwriter.printmessage(line)
                    tempdate = line.lstrip("/Date (")
                    tempdate = tempdate.split('+')[0].strip("\r\n")
                    
                    logwriter.printmessage(tempdate)
                    if str(tempdate).endswith(")"):
                        #logwriter.printmessage("Bad time ending")
                        tempdate = tempdate.rstrip(")")
                        #logwriter.printmessage(tempdate)
                    senddate = parse(tempdate)
                    
                    #logwriter.printmessage ("MIDDLE3 "+str(filename))
                    senddate = str(senddate).replace(' ','_').replace(':','-')
                  
                    tail = str(senddate)+"_"+tail 
                    
            datereader.close()
        #htmlfile = os.path.join(root, 'Message.html')        
        temppdfpointer = tail+".pdf"
        #print (temppdfpointer)
        pdf_file = os.path.join(root, temppdfpointer)
        #origfile = os.path.join(root, filename)
        #logwriter.printmessage("2 ::"+str(filename))
        if str(filename).endswith('html'):                  
            #logwriter.printmessage("Html file found in {}".format(file_path))            
            
            #cmd_wkhtmltopdf = ["wkhtmltopdf", "--image-quality", "100", "-n", "--no-outline", "--enable-smart-shrinking", file_path, pdf_file] # -q is quiet            
            cmd_wkhtmltopdf = ["wkhtmltopdf", "-n", file_path, pdf_file] # -q is quiet            
            
            #logwriter.printmessage(cmd_wkhtmltopdf)
            doCmd(cmd_wkhtmltopdf)
            
        elif str(filename).endswith('rtf') or str(filename).endswith('txt'):
            #logwriter.printmessage ("RTF file in %s"%root)
            """ABIWORD rtf & txt --> pdf"""
            cmd_abi = ['abiword', '--plugin=AbiCommand', '-t', pdf_file, file_path]
            #logwriter.printmessage(cmd_abi)
            doCmd(cmd_abi)                      


        #Final check that correctly named pdf file exists, if not create it
        if not os.path.isfile(pdf_file):
            tempf = open(os.path.join(root, pdf_file), 'w')
            tempf.close()

        temppdfApointer = tail+"A.pdf"
        #temppointer = tail+"A-test.pdf"
        pdfAFile = os.path.join(root, temppdfApointer)
        #testPdfFile = os.path.join(root, temppointer)
        #logwriter.printmessage(testPdfFile)
        outputF = "-o"+os.path.join(root, temppdfApointer)
        #logwriter.printmessage("Final output file == {}".format(outputF))       
        #Added -dNOTRANSPARENCY into below cmd. Lowers quality but ensures valid PDF/A-3b file
        #Removed -dUseCIEColor -dColorConversionStrategy=RGG -sColorConversionStrategy
        #Changed /prepress --> /printer (gets device independent color profile)             
        
        
        cmd_gs = [ghostscript_path, '-dPDFA=3',
        '-dBATCH', '-dNOPAUSE', '-dNOOUTERSAVE', '-dNOSAFER', '-dPDFSETTINGS=/prepress', '-dQUIET',
        '-dPDFACompatibilityPolicy=1', '-dAutoFilterColorImages=false', '-dColorImageFilter=/FlateEncode',
        '-dAutoFilterGrayImages=false', '-dGrayImageFilter=/FlateEncode', '-dMonoImageFilter=/FlateEncode',
        '-sColorConversionStrategy=UseDeviceIndependentColor','-dEmbedAllFonts=true',
        '-sDEVICE=pdfwrite', outputF, pdfa_def_path, pdf_file]
        """
        cmd_gs = [ghostscript_path, '-dPDFA=3',
        '-dBATCH', '-dNOPAUSE', '-dNOOUTERSAVE', '-dNOSAFER', '-dPDFSETTINGS=/prepress', '-dQUIET',
        '-dPDFACompatibilityPolicy=1', '-dAutoFilterColorImages=false', '-dColorImageFilter=/FlateEncode',
        '-dAutoFilterGrayImages=false', '-dGrayImageFilter=/FlateEncode', '-dMonoImageFilter=/FlateEncode',
        '-sColorConversionStrategy=UseDeviceIndependentColor','-dEmbedAllFonts=true',
        '-sDEVICE=pdfwrite', outputF, pdfa_def_path, pdf_file]
        """
         
        if os.path.isfile(os.path.join(root, "attachments.info")):
            infofilelocation = os.path.join(root, "attachments.info")
            #print ("KOKO ====",os.path.getsize(infofilelocation))
            cmd_gs.append (infofilelocation)
        #logwriter.printmessage(cmd_gs)
        doCmd(cmd_gs)
       
        embedMetadataToPDFA(pdfAFile)
        
        #VERAPDF validation process
        """
        --format Chooses output format. Default: mrr
        Possible Values: [xml, mrr, html, text, batch]
        """
        if validate==True:        
            validationfile = os.path.join(root, temppdfApointer)
            cmd_vera = [verapdf_path, '--format', 'xml', validationfile]
            results = doCmd(cmd_vera)          
            if 'isCompliant="true"' in str(results):
                logwriter.printmessage ("VALIDATION SUCCESS {}".format(temptail))
                pass
            else:
                logwriter.printmessage ("VALIDATION ERROR {}".format(validationfile))
    
    except OSError:
        logwriter.printmessage("ERROR"+str(OSError.message()))
        #pass
        
    return 1 #Tells to counter that one process has ended

def embedMetadataToPDFA(pdfAFile):
    """
        Below part will read all metadata from modifiedcombinedheaders.txt 
        and embeds those into pdf
    """
    #logwriter.printmessage("Final pdf a = {}".format(pdfAFile))
    headersFile = os.path.join(os.path.dirname(pdfAFile), "ModifiedCombinedHeaders.txt")
    #logwriter.printmessage("Header file = {}".format(headersFile))
    if os.path.isfile(headersFile) and os.path.isfile(pdfAFile):
        newMetaData = PdfDict()
        pdfReader = PdfReader(pdfAFile) #PdfReader object is a subclass of PdfDict, which allows easy access to an entire document:                
        #pdfWriter = PdfWriter(testPdfFile)
        headerReader = open(headersFile, 'r')
        for line in headerReader.readlines():
            if line.startswith('/'):
                try:
                    lineB, lineE = line.split(' ', 1)
                    lineB = lineB.lstrip('/')
                    lineE = lineE.lstrip('(').strip().rstrip(')')
                    if 'not found' in lineE:
                        #logwriter.printmessage("WARNING invalid date format in {}".format(root)) # to find misformed dateformats
                        pass
                    
                    newMetaData.__setattr__(lineB, lineE)
                    #metadataWriter.Info.LineB=str(lineE)
                    #logwriter.printmessage(str(lineB)+" -- "+str(lineE))
                except ValueError:
                    pass
        #logwriter.printmessage(newMetaData)
        pdfReader.Info.update(newMetaData)
        #metadataWriter.Info.update(newMetaData)
        PdfWriter(pdfAFile, trailer=pdfReader).write()
    return
               
    
"""
Html page fetcher, used to collect html formatted email elements

def fetchHtmlForThePage(url):
    #logwriter.printmessage("Fetching file : {}".format(url))
    url = "file://"+url
  
    try:
        options = Options()
        options.add_argument("--headless")
        #options.set_headless(Headless=True)
        browser = webdriver.Firefox(firefox_options=options)    
        browser.set_page_load_timeout(10)
        browser.get(url)        
        #logwriter.printmessage("-------")
        
        #logwriter.printmessage("type {}".format(type(browser.page_source.decode)))
        html = browser.page_source
        soupHtml = BeautifulSoup(html, 'lxml')        
        browser.quit()
    except TimeoutError:
        soupHtml = '0'
    return soupHtml
"""




"""
Def that will get the return value from multiprocess def 

def multiprocess_results(result):
    global onGoingPoolCounter
    global completedConversion
    #logwriter.printmessage(result)
    if '-1' in str(result):
        completedConversion+=abs(result)
    onGoingPoolCounter+=result
    logwriter.printmessage ("Completed :{}".format(onGoingPoolCounter))
    #logwriter.printmessage("Completed:"+str(completedConversion))
    return
"""


"""
This def multiprocess all found calender paths
"""
def multiprocessCalender(basepath): #15, 50, 1398    
    try:
        convertMetaFile.convertMetaData(basepath, handleMessages)
        value = True
    except:
        #logwriter.printmessage("Metadata conversion problem {}".format(basepath))
        value = False  
    #location, title/subject, sender, starttime, duration             
    if value != False:
        #logwriter.printmessage("Stars to process Calanders..")
        filename =''.join(random.choices(alphanum, k=6)) #Creates a random uuid name for the calendar event
        if os.path.isdir(os.path.join(basepath, "Attachments")):            
            #logwriter.printmessage("Meeting with attachments {}".format(basepath))
            convertAttachmentNoEmbed(basepath)
        
        try:
            reader = open(os.path.join(basepath,"ModifiedCombinedHeaders.txt"),'r')
            allList = [] #List containing tuplses of key-value
           
            for line in reader:                
                line = line[1:-2]                    
                tempsplit = line.split('(',1)                    
                try:
                    tempkey = tempsplit[0].strip()
                    tempvalue = tempsplit[1].strip()                    
                    tempTuple = (tempkey,tempvalue)
                    allList.append(tempTuple)
                except IndexError:
                    pass
            
            #finalDict = {} # for collecting the needed ones, but only the ones with data
            #keys = ["Subject", "Title", "Sendername", "Senderemailaddress", "Starttime", "Endtime", "Location"]
            subjectKeys = ["Subject","Title","Conversationtopic"]
            senderKeys = ["Sendername","Sentrepresentingname"]            
            startKeys = ["Starttime"]
            endKeys = ["Endtime"]
            locationKeys = ["Location"]
            participantKeys = ["Displayname","Recipientdisplayname", "Sendername", "Sentrepresentingname"]
            #recurrenceKeys = ["Recurrencepattern"]
            
            subjectValues = ["Aihe: "]
            senderValues = ["Lähettäjä: "]
            startValues = ["Alkamisaika: "]
            endValues = ["Päättymisaika: "]
            locationValues = ["Paikka: "]
            participantValues = ["Osallistujat: "]
            #recurrenceValues = ["Toistuvuus: "]
            #recurrenceStartValues = ["Ensimmäinen: "]
            #recurrenceEndValues = ["Viimeinen: "]
            
            for tempTuple in allList:
                
                #if "Appointment00181" in basepath:
                    #logwriter.printmessage("XXX-{},{}".format(key, tempDict[key]))
                if tempTuple[0] in subjectKeys:
                    if tempTuple[1] not in subjectValues:
                        subjectValues.append(tempTuple[1])
                if tempTuple[0] in senderKeys:
                    if tempTuple[1] not in senderValues:
                        senderValues.append(tempTuple[1])
                if tempTuple[0] in startKeys:
                    if tempTuple[1] not in startValues:
                        startValues.append(tempTuple[1])
                if tempTuple[0] in endKeys:
                    if tempTuple[1] not in endValues:
                        endValues.append(tempTuple[1])
                if tempTuple[0] in locationKeys:
                    if tempTuple[1] not in locationValues:
                        locationValues.append(tempTuple[1])
                if tempTuple[0] in participantKeys:
                    #logwriter.printmessage("Found participant : {}".format(tempDict[key]))
                    if tempTuple[1] not in participantValues:
                        participantValues.append(tempTuple[1])                
                             
            
            finalList = []
            if len(subjectValues)>1: finalList.append(subjectValues)
            if len(senderValues)>1: finalList.append(senderValues)
            if len(startValues)>1: finalList.append(startValues)
            if len(endValues)>1: finalList.append(endValues)
            if len(locationValues)>1: finalList.append(locationValues)
            if len(participantValues)>1: finalList.append(participantValues)            
            
            
            #Now finalList contains all with values, write those to htmlfile
            calenderInfos = ""
            for onelist in finalList:
                header = onelist[0] #Topic always the first
                onelist.pop(0) #Removes the header                
                calenderInfos+="<b>{}</b>{}<br>".format(header, onelist)
            #logwriter.printmessage("Subjects: {}".format(subjectValues))
            #logwriter.printmessage("HTML {}".format(os.path.dirname(path)))
            
            calenderInfos = calenderInfos.replace("[","").replace("]","").replace("'","")
            filenamestart = os.path.join(basepath,filename)
            htmlFile = filenamestart+".htm"
            
            pdf_file = filenamestart+".pdf"
            pdfA_file = filenamestart+"_A.pdf"
            #logwriter.printmessage("HTML path = {}".format(htmlFile))
            htmlWriter = open(htmlFile, 'w')
            htmlWriter.write("{}{}{}".format(htmlBegin, calenderInfos, htmlEnd))
            htmlWriter.close()
            if (os.path.isfile(htmlFile)):        
                #cmd_wkhtmltopdf = ["wkhtmltopdf", "--image-quality", "100", "-n", "--no-outline", "--enable-smart-shrinking", htmlFile, pdf_file] # -q is quiet            
                cmd_wkhtmltopdf = ["wkhtmltopdf","-n", htmlFile, pdf_file] # -q is quiet            
                
                #logwriter.printmessage(cmd_wkhtmltopdf)
                doCmd(cmd_wkhtmltopdf)
            
            if(os.path.isfile(pdf_file)):
                outputF = "-o"+pdfA_file
                """
                cmd_gs = [ghostscript_path, '-dPDFA=3',
            '-dBATCH', '-dNOPAUSE', '-dNOOUTERSAVE', '-dNOSAFER', '-dPDFSETTINGS=/prepress',
            '-dPDFACompatibilityPolicy=1', '-dAutoFilterColorImages=false', '-dColorImageFilter=/FlateEncode',
            '-dAutoFilterGrayImages=false', '-dGrayImageFilter=/FlateEncode', '-dMonoImageFilter=/FlateEncode',
            '-sColorConversionStrategy=UseDeviceIndependentColor','-dEmbedAllFonts=true',
            '-sDEVICE=pdfwrite', outputF, 'pdfa_def.ps', pdf_file]
                """
                
                cmd_gs = [ghostscript_path, '-dPDFA=3',
            '-dBATCH', '-dNOPAUSE', '-dNOOUTERSAVE', '-dNOSAFER', '-dPDFSETTINGS=/prepress',
            '-dPDFACompatibilityPolicy=1', '-dAutoFilterColorImages=false', '-dColorImageFilter=/FlateEncode',
            '-dAutoFilterGrayImages=false', '-dGrayImageFilter=/FlateEncode', '-dMonoImageFilter=/FlateEncode',
            '-sColorConversionStrategy=UseDeviceIndependentColor','-dEmbedAllFonts=true',
            '-sDEVICE=pdfwrite', outputF, pdfa_def_path, pdf_file]         
                doCmd(cmd_gs)
            #logwriter.printmessage("{}{}{}{}{}".format(htmlSubject, htmlTitle, htmlSender, htmlDate, htmlDuration))
            if os.path.isfile(pdfA_file):
                #Adds metadata to calendar file
                embedMetadataToPDFA(pdfA_file)
                if validate==True:        
                    #validationfile = os.path.join(root, temppdfApointer)
                    cmd_vera = [verapdf_path, '--format', 'xml', pdfA_file]
                    results = doCmd(cmd_vera)          
                    if 'isCompliant="true"' in str(results):
                        logwriter.printmessage ("VALIDATION SUCCESS {}".format(pdfA_file))
                        pass
                    else:
                        logwriter.printmessage ("VALIDATION ERROR {}".format(pdfA_file))
                
        except:
            logwriter.printmessage("Something strange in {}".format(basepath))
            pass    
        
    return

"""
This def handles calender markings 
"""

"""
def handleCalender(calendarDir):
    calCount = 0
    maxCpuCount = multiprocessing.cpu_count()
    processCount = maxCpuCount-2       
    metpool = Pool(processCount) #Uses all available or required cores 
    
    logwriter.printmessage("Handling calenders in dir {}".format(calendarDir))
    for root,dirs,files in os.walk(calendarDir):
        for file in files:
            if("Attachments" not in root): #We do not process attachments paths for calenders                
                #logwriter.printmessage(os.path.join(root,file))
                metpool.apply_async(convertMetaFile.convertMetaData, args=(root,))
    metpool.close()
    metpool.join()    
    calpool = Pool(processCount) #Uses all available or required cores
    for root,dirs,files in os.walk(calendarDir):
        for file in files:            
            if(file=="ModifiedCombinedHeaders.txt"):
                calCount+=1
                calpath = os.path.join(root,file)
                logwriter.printmessage("CAL path = {}".format(calpath))
                calpool.apply_async(multiprocessCalender, args=(calpath,))                
    calpool.close()
    calpool.join()
    
    
    pdfcount = 0
    for root,dirs,files in os.walk(calendarDir):
        #logwriter.printmessage("{}-{}".format(root, len(files)))
        pdffound = False
        for file in files: 
            if(file.endswith('_A.pdf')):
                pdffound=True
                pdfcount+=1
        if pdffound==False:
            logwriter.printmessage("PDF not found in {}".format(root))
    logwriter.printmessage("Found {} calender items and {} pdf files".format(calCount, pdfcount))
    return
"""


def walkthrough_topdf_ended():
    global onGoingPoolCounter
    if onGoingPoolCounter==0:
        logwriter.printmessage("Ei varmaan koskaan tulostu..")
        return True
        #que.put("END")
    else:
        return False   


def attachmenttypes (walk_dir):
    #Creates an empty dictionary that holds attachments types
    allendingtypes = {}
    EndingsInfolocation = os.path.join(walk_dir, "AttachmentTypes.html")
    #print("From attachmenttypes:")
    logwriter.printmessage("Beginning attachment type analysis")
    for root, dirs, files in os.walk(walk_dir):
        for dirname in dirs:
            if str(dirname).endswith("Attachments"):
                completedir = os.path.join(root,dirname)
   
                for filenames in os.listdir(completedir):
                    ending = os.path.splitext(filenames)[-1].lower()
                    if ending not in allendingtypes:
                        allendingtypes[ending] = 1
                        logwriter.printmessage("ending "+str(ending)+" in folder "+str(completedir))

                    else:
                        i = allendingtypes.get(ending)
                        i = i+1
                        allendingtypes[ending] = i

    sortedtypes = sorted(allendingtypes.items(), key=operator.itemgetter(1), reverse=True)
    endingsfile = open(EndingsInfolocation, 'w')
    #html = '<table><tr><th>' + '</th><th>'.join(allendingtypes.keys()) + '</th></tr>'
    html = '<!DOCTYPE html><html><body><table style="width:30%" border="1">'
    for keys, values in sortedtypes:
    #for keys, values in allendingtypes.items():
        html += '<tr><th>'+str(keys)+'</th><td>'+str(values)+'</td></tr>'


    html+= '</table></body></html>'
    endingsfile.write(html)


def endofscript():
    global ConversionStartTime
    runtime = display_time(time.time()-ConversionStartTime, 3)
    logwriter.printmessage("COMPLETE PROSESSING TIME = {}".format(runtime))
    logwriter.printmessage("Closing log file!")
    logwriter.printmessage("FINISHED")
    logwriter.closeLogFile()
    #logfile.close
    return

def ensureFileType(fullFilePath):    
    #logwriter.printmessage(extensions)
    #Microsoft Outlook 2003-2007 PUID = x-fmt/249        
    #/home/user/droid/droid.sh -Nr konferenssit -Ns "/home/user/droid/DROID_SignatureFile_V92.xml" -q
    
    fileending = os.path.splitext(fullFilePath)[1]
    if fileending in extensions:
    #if str(fullFilePath).endswith('.py') or str(fullFilePath).endswith('.jar') or str(fullFilePath).endswith('.log'):
        result='not checked'
    else:
        droidcmd = [droid_path, '-Nr', fullFilePath, '-Ns', droid_signature_path, '-Nc', droid_container_path, '--quiet']
        #droidcmd = [droid_path, '-Nr', fullFilePath, '--quiet']        
        droidprocess= subprocess.Popen(droidcmd, stdout=subprocess.PIPE, universal_newlines=True)
        result = str(droidprocess.communicate()[0])
        logwriter.printmessage("Droid raw results {}".format(result))
        try:
            result = result.split(',')[-1]
            
        except IndexError:
            pass
    logwriter.printmessage("Droid final result {}".format(result))
    return result

intervals = (
    ('weeks', 604800),  # 60 * 60 * 24 * 7
    ('days', 86400),    # 60 * 60 * 24
    ('hours', 3600),    # 60 * 60
    ('minutes', 60),
    ('seconds', 1),
    )
def display_time(seconds, granularity=2): #Converts seconds to weeks, days, hours, minutes
    result = []
    for name, count in intervals:
        value = seconds // count
        if value:
            seconds -= value * count
            if value == 1:
                name = name.rstrip('s')
            result.append("{} {}".format(value, name))
    return ', '.join(result[:granularity])


def zipDemoFolders(includedPaths, rootPath):
    logwriter.printmessage("Tar.gzipping demo paths")    
    tar = tarfile.open(os.path.join(rootPath,"demofiles.tar.gz"), 'w:gz')
    for path in includedPaths:        
        tar.add(path, arcname=os.path.basename(path))
    
    return

def multiprocessFoundDir(path):    
    convertMetaFile.convertMetaData(path, handleMessages) #Converts all metadatafiles to accepted format
    msgFiles = []
    for file in os.listdir(path):
        if file.startswith("Message"):
            msgFiles.append(file)
    msgFile = sorted(msgFiles)[0]
    msgPDFConversion(path, msgFile)
    
    
    return
   
"""
New found folder handler
"""
def handleFoundDirs(messageDirs, appointmentDirs, meetingDirs):
    maxCpuCount = multiprocessing.cpu_count()    
    #Counts the number of messages&appointments
    allDirs = messageDirs+appointmentDirs+meetingDirs    
    logwriter.printmessage("{} folders with content".format(len(allDirs)))
        
    processCount = max(maxCpuCount-1, len(allDirs))
    if processCount>maxCpuCount-1:
        processCount = maxCpuCount-1
    
    pool = Pool(processCount) #Uses all available or required cores
    count = 0
  
        
    if handleMessages: #Handles messagedirs
        logwriter.printmessage("##### Handling {} messages #####".format(len(messageDirs)))
        for oneDir in messageDirs:
            """Split to multiprosessing already in here"""
            #multiprocess_results(1) #Adds one to ongoing pool counter
            
            pool.apply_async(multiprocessFoundDir, args=(oneDir,))
            count+=1                         
        
    meets = len(appointmentDirs)+len(meetingDirs)
    if handleMeetings and meets>0: #handles meetingdirs only if exists        
        logwriter.printmessage("##### Handling {} meetings #####".format(len(appointmentDirs)+len(meetingDirs)))
        allMeetDirs = appointmentDirs+meetingDirs
        for oneDir in allMeetDirs:
            pool.apply_async(multiprocessCalender, args=(oneDir,))
            count+=1

    pool.close()
    pool.join()
    return


"""
This is the actual main process that is launched
delflag, fileflag, zipflag and attachmentflag = boolean --> by default all are set to False
dataDir = path to directory where the uploaded file is
dataFile = filename of the uploaded file
"""

def mainActor(dataDir, dataFile, fileflag, attachmentflag, delflag, zipflag, apirun=None):                  
    #appointmentFound = False #Used to identify if an appointment path is found
  
    
    #global beginfilenamestriped
    #original_walk_dir = os.path.dirname(os.path.abspath(__file__))    
    #walk_dir = dataDir # Changed to dir where the uploaded or otherwise found datafile is
    
    #logwriter.printmessage ("PYTHON VERSION: {}".format(sys.version_info))
    #logwriter.printmessage ("Run dir: {}".format(dataDir))
    ConversionStartTime = time.time()
    if (dataDir==""):
        dataDir = os.path.dirname(os.path.abspath(dataFile))
        print("Directory is now {}".format(dataDir))  
    fullpath = os.path.join(dataDir, dataFile)
    logwriter.printmessage("Begin datafile = {} in dir {}".format(dataFile, dataDir))   
    logwriter.printmessage("Full path to item = {}".format(fullpath))
    df = ensureFileType(fullpath) #Full path to uploaded file
    logwriter.printmessage("Filetype check : {}".format(df))
    individualEmlFile = False           
    
    if "Attachment" not in dataDir:
    
        if 'fmt/950' in df or 'x-fmt/430' in df or 'x-fmt/249' in df or 'x-fmt/248' in df:
            if 'fmt/950' in df or 'x-fmt/430' in df: #this is either .eml or .mbox or mbox without ending or .msg                      
                logwriter.printmessage(dataFile)    
                beginfilenamestriped = cleaner.stripFileName(dataFile, dataDir)
                if 'x-fmt/430' in df: #.msg file
                    maxCpuCount = 1 #Individual file, handle with one core
                    #logwriter.printmessage(beginfilenamestriped.split('.')[0])
                    emlfile = str(beginfilenamestriped.split('.')[0])+".eml"
                    #msgObj = MsOxMessage(beginfilenamestriped)
                    #jsonString = msgObj.get_message_as_json()
                    #msg_properties_dict = msgObj.get_
                    
                    emlconvertcmd = ['msgconvert', os.path.join(dataDir, beginfilenamestriped)] #converts msg to mbox format
                    #logwriter.printmessage(emlconvertcmd)
                    emlprocess= subprocess.Popen(emlconvertcmd)
                    emlprocess.communicate() 
                    #msgconvert does the eml file into base dir, move it to correct one
                    emlpath = os.path.join(dataDir, emlfile)
                    logwriter.printmessage("Assumed eml path = {}".format(emlpath))
                    if (os.path.isfile(os.path.join(os.path.dirname(os.path.abspath(__file__)), emlfile))):
                        logwriter.printmessage("File found in a wrong dir, move to correct one")                    
                        cleaner.renameFile(os.path.join(os.path.dirname(os.path.abspath(__file__)), emlfile), emlpath)
                    
                                    
                    beginfilenamestriped = emlfile
                    individualEmlFile = True
                
                #Thus fmt/950 may mean either .eml or .mbox lets check it with file
                filecmd = ['file', dataDir+'/'+beginfilenamestriped]
                fileprocess= subprocess.Popen(filecmd, stdout=subprocess.PIPE, universal_newlines=True)
                filecmdresult = str(fileprocess.communicate()[0])
                logwriter.printmessage(filecmdresult)
                maxCpuCount = multiprocessing.cpu_count()
                if "SMTP" in filecmdresult or individualEmlFile:
                    exportPath = dataDir+'/eml-emails/'+beginfilenamestriped+'.export'     
                else:
                    exportPath = dataDir+'/'+beginfilenamestriped+'.export'
                
                if os.path.isdir(exportPath):
                    cleaner.removeDir(exportPath)
                if not os.path.isdir(exportPath):
                    logwriter.printmessage("Path does not exists, creating {}".format(exportPath))
                    os.makedirs(exportPath)
                    
                if "SMTP" in filecmdresult or individualEmlFile:
                    #Individual file, handle with one core
                    #maxCpuCount = 1
                    logwriter.printmessage("Appers to be individual eml file: {}".format(beginfilenamestriped))
                    #cleaner.renameFile(os.path.join(walk_dir,item), os.path.join(exportPath, beginfilenamestriped))
                    cmdmove = ['mv', os.path.join(dataDir, beginfilenamestriped), os.path.join(exportPath, beginfilenamestriped)]                
                    logwriter.printmessage(cmdmove)
                    cpprocess = subprocess.Popen(cmdmove)
                    cpprocess.communicate()
                    emlManipulate(exportPath)
                    modified_walk_dir = exportPath
                    walkthrough_topdf(modified_walk_dir) 
                else:
                    logwriter.printmessage("mbox file found, extracting")           
                    splitpath = "-o"+exportPath
                    #logwriter.printmessage("Mbox export path = {}".format(splitpath))
                    cmdgitsplit = ['git', 'mailsplit', '-b', '-d5', splitpath, dataDir+'/'+beginfilenamestriped]
                    logwriter.printmessage("{}".format(cmdgitsplit))
                    gitsplitprocess = subprocess.Popen(cmdgitsplit)
                    gitsplitprocess.communicate()
                    #Below function will browse the newly create files and moves those to correct folders                  
                    mboxManipulate(exportPath)
                    modified_walk_dir = exportPath
                    #Counts the suitable amount of cores to be used, max is defined elsewhere
                    finalDir = modified_walk_dir
                    for root, dirs, files in os.walk(modified_walk_dir):
                        if len(dirs)<maxCpuCount:
                            maxCpuCount = len(dirs)
                        for dirname in dirs:
                            if "Message" in dirname: #If Message is found we know that we are in a right path                            
                                logwriter.printmessage("Detected correct directory {}".format(dirname))                            
                                finalDir = modified_walk_dir
                                break
                        break
                    walkthrough_topdf(finalDir)
                                
               
                if delflag:
                    cleanTheseDirs.append(exportPath) 
                
            #logwriter.printmessage(droidfiletype+"-x-fmt/249")    
            if 'x-fmt/249' in df or 'x-fmt/248' in df: #pst or ost file
            #if (str(item)).lower().endswith(".pst") or (str(item)).lower().endswith(".ost"):
                #for beginfile in glob.glob("*.pst"):
                firstCall = True            
                beginfilenamestriped=cleaner.stripFileName(dataFile, dataDir)
                originalPath = os.path.join(dataDir, beginfilenamestriped)
                #beginfilenamestriped = os.path.join(dataDir, beginfilenamestriped)
                exportPath = fullpath+'.export/'
               
                if(os.path.isdir(exportPath)):
                    cleaner.removeDir(exportPath) #calls to remove the extracted folder if it exists
            
                #cmdpffexport = ['pffexport', '-f', 'all', originalPath] #Extracts .pst or .ost file              
                cmdpffexport = ['pffexport', '-f', 'all', '-q', '-t', fullpath, originalPath] #Extracts .pst or .ost file              
                logwriter.printmessage("##### Extraction of Outlook datafile starts #####")
                exportStart = time.time()
                pffprocess = subprocess.Popen(cmdpffexport)
                pffprocess.communicate()                    
                exportEnd = time.time()
                logwriter.printmessage("Extraction took {}s".format(round(exportEnd-exportStart, 3)))
                    
                #logwriter.printmessage ("beginfile stripped = "+beginfilenamestriped)            
                #logwriter.printmessage("Export path = {}".format(exportPath))
                
                #finds all exported empty directories and deletes those
                emptyDirs = folderSeeker.findAllEmptyDirs(exportPath)
                if len(emptyDirs)>0:
                    logwriter.printmessage("##### Removing empty dirs #####")
                    cleaner.handleListofEmptyDirs(emptyDirs)
                
                """Finds all folders containing messages and appointments"""
                logwriter.printmessage("##### Content folder seeking starts #####")
                contentSeekStart = time.time()
                messageDirs, appointmentDirs, meetingDirs = folderSeeker.findAllContentDirs(exportPath)
                #logwriter.printmessage("Found {} message, {} appointments and {} meetings"
                #      .format(len(messageDirs), len(appointmentDirs), len(meetingDirs)))
                contentSeekEnd = time.time()
                logwriter.printmessage("Seek took {}".format(round(contentSeekEnd-contentSeekStart,3)))
                
                """Now handling for the found dirs"""
                #handleCalender((calanderDir))
                #walkthrough_topdf(modified_walk_dir)
                handleFoundDirs(messageDirs, appointmentDirs, meetingDirs)
                            
                    
            if delflag:
                logwriter.printmessage("DEL flag ending {}".format(cleanTheseDirs))
                
                deletedThings = cleaner.handleFinalCleanup(exportPath)
                while deletedThings>0:
                    logwriter.printmessage("Looping final cleanup")
                    deletedThings = cleaner.handleFinalCleanup(exportPath)
                
                
            if zipflag:
                logwriter.printmessage ("ZIPFLAG found, Zipping folder structure...")
                
                shutil.make_archive(fullpath, 'zip', exportPath)
                expectedZipFile = fullpath+".zip"
                if os.path.isfile(expectedZipFile):           
                    logwriter.printmessage("Inside zipping -{}".format(expectedZipFile))
                    return expectedZipFile
            
        else:#Exits if exported language is currently not supported
            logwriter.printmessage("At this phase this sofware only supports Outlook .ost and .pst files. Others are ignored..")
            #return "Invalid filetype"
        
        return "This should never return"

#csvwriter = open('metainfo.csv', 'w')

if __name__ == "__main__":
    beginfilenamestriped = ""
    sep = ";"
    emails={}
        
    """CLI run"""
    dataDir = ""
    dataFile = ""
    ConversionStartTime = time.time()
    original_walk_dir = os.path.dirname(os.path.abspath(__file__)) 
    
    endings = (".pst", ".ost", ".eml", ".msg", ".mbox")
    
    for x in range (1, len(sys.argv)):
        if sys.argv[x]=="-del":
            delflag = True
            logwriter.printmessage ("Delete action initiated")
        elif sys.argv[x]=="-d":            
            dataDir = sys.argv[x+1]
            logwriter.printmessage ("Dir action initiated with value "+str(dataDir))
        elif sys.argv[x]=="-zip":
            zipflag = True
            logwriter.printmessage ("Zip action initiated")
        elif sys.argv[x]=="-a":
            attachmentflag = True
            logwriter.printmessage ("Count attachments action initiated")
    
    walk_dir = os.getcwd() # should be the directory where program was run
    
    
        
    if dataFile == "" or dataDir=="":   
        #So running without special paths
        """
        First seeks all files and add suitable ones to a list           
        """
        emailPaths = []
        for root, dirs, items in os.walk(walk_dir):    
            for item in items:
                if (item.endswith(endings)): 
                    #logwriter.printmessage("Found {}".format(os.path.join(root,item)))
                    emailPaths.append(os.path.join(root,item))
        
        for item in emailPaths:        
            if os.path.isfile(item):
                logwriter.printmessage("Cli run found {}".format(str(item)))            
                dataDir = os.path.dirname(os.path.abspath(item))
                dataFile = os.path.basename(item)                    
                finalZipFilePath = mainActor(dataDir, dataFile, fileflag, attachmentflag, delflag, zipflag)
             
              
    #logwriter.printmessage("Final Zip file = {}".format(finalZipFilePath))        
    
    atexit.register(endofscript)




"""****************************************************************************
*******************CURRENTLY UNUSED BUT CONSIDERED MODULES*********************
****************************************************************************"""


"""Working attachment infos
            infofile.write("/F ("+filepath+") (r) file def\n")
            infofile.write("[/_objdef {mstream"+str(i)+"} /type /stream /OBJ pdfmark\n")            
            infofile.write("[{mstream"+str(i)+"} F /PUT pdfmark\n")
            infofile.write("[{mstream"+str(i)+"} << /Type /EmbeddedFile /Subtype (application/octet-stream) >> /PUT pdfmark\n")
            infofile.write("[ /Name ("+fn+")\n")
            infofile.write("/FS << /Type /Filespec\n")
            infofile.write("/F ("+fn+")\n")
            infofile.write("/UF ("+fn+")\n")
            infofile.write("/Desc ("+fn+")\n")
            #infofile.write("/Subtype (application/octet-stream)")
            infofile.write("/AFRelationship/Supplement")
            #infofile.write("/Subtype/application/octet-stream")
            infofile.write("/EF << /F {mstream"+str(i)+"}")
            infofile.write(">>")
            #infofile.write("/Subtype(application/octet-stream)\n")
            infofile.write(">>\n")

            #infofile.write("/Params<</Size "+os.path.getsize(files)+">>")
            infofile.write("/EMBED pdfmark\n")
            infofile.write("[{mstream"+str(i)+"} /CLOSE pdfmark\n")
            i=i+1
"""
