#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Aug 22 11:44:09 2017
Modified on Tue Jun 16

@author: digitalia-aj
"""
import os
import subprocess
import shutil

#Own created codes
import logwriter
import cleaner    
from IPython.utils.capture import capture_output
#from nltk.corpus.reader import wordlist

gs = 'gs' #Path to ghostscript
pdfa_def_path = '/home/digitalia-aj/extraspace/eclipse-workspace/Email-converter/pdfa_def.ps' #if gs updates, change content
openof = '/usr/bin/soffice' #Path to libre / open office

"""Lists defining convertible filetypes"""
pictureList = ['jpg', 'bmp', 'gif']
officeList = ['doc','docx','rtf']
wordList = ['wpt','wcm','wri','wpd', 'wps']
pptList = ['ppt','pptx', 'pptm']
xlsList = ['xls','xlsx']
pdfList = ['pdf']
allList = pictureList+officeList+wordList+pptList+xlsList+pdfList
#endFix = "_converted"

def doconvert(cmd):    
    try:
        #logwriter.printmessage("Convert command: {}".format(cmd))
        convertprocess = subprocess.Popen(cmd, stdout=subprocess.PIPE)                              
        convertprocess.communicate()                        
    except subprocess.CalledProcessError as e:
        logwriter.printmessage("Conversion timeout error > {}".format(e.message))
        pass
    return

def unzip(attachpath:str, filename:str):
    try:
        """If a zip file, it's content is extracted and an extraction path is returned"""
        unzipPath= os.path.join(attachpath, filename+'-unzipped')
        shutil.unpack_archive(os.path.join(attachpath, filename),unzipPath)
        #with zipfile.ZipFile(os.path.join(attachpath, filename), 'r') as myZip:        
        #    myZip.extractall(unzipPath)
        return unzipPath
    except Exception as e:
        logwriter.printmessage("Cannot extract zip file {} --> {}".format(os.path.join(attachpath, filename), e))
        pass
    return ""
    

def convertfile(fullPath:str, convertToPDFA:bool, deleteOriginalAttachment:bool):
    """Converts item (str path or filename, both works) into archival format, returns ??]"""            
    attachpath, filename = fullPath.rsplit(os.path.sep,1)    
    filebegin, fileend = filename.rsplit('.',1)
    #endinglow = fileend.lower()
    
    PNGfile = os.path.join(attachpath, filebegin)+".png"
    ODSfile = os.path.join(attachpath, filebegin)+".ods"
    #ODPfile = os.path.join(attachpath, filebegin)+".odp"
    PDFfile = os.path.join(attachpath, filebegin)+".pdf"
    word_dockcmd = [openof, '--writer', '--headless', '--nologo', '--invisible', '--norestore', '--convert-to', 'pdf', '--outdir', attachpath, fullPath]
    word_pptcmd = [openof, '--impress','--headless', '--nologo', '--invisible', '--norestore', '--convert-to', 'pdf', '--outdir', attachpath, fullPath]
    word_xlscmd = [openof, '--calc','--headless', '--nologo', '--invisible','--norestore', '--convert-to', 'ods', '--outdir', attachpath, fullPath]    
    #word_xlscmd = ['/usr/lib/libreoffice/program/soffice.bin', '--calc','--headless', '--nologo', '--invisible','--norestore', '--convert-to', 'ods', '--outdir', attachpath, fullPath]
    
    abicmd = ['abiword', '-t', os.path.join(attachpath, filebegin)+".pdf", fullPath]
    piccmd = ['convert', fullPath, PNGfile]    
    
    endinglow = fileend.lower()
    
    deleteTempFile = False
    isPDFfile = False
    finalFile = ""
    #Here checks if a conversion needs to be done
    if endinglow in allList:
    
        if endinglow in officeList: 
            doconvert(abicmd)
            if os.path.isfile((PDFfile)):            
                isPDFfile = True
                if convertToPDFA:
                    deleteTempFile = True            
                finalFile = PDFfile
        elif endinglow in wordList:
            doconvert(word_dockcmd)
            if os.path.isfile(PDFfile): 
                isPDFfile = True 
                if convertToPDFA:
                    deleteTempFile = True              
                finalFile = PDFfile
        elif endinglow.endswith("pdf"): #incoming file is a pdf document        
            isPDFfile = True
            finalFile = PDFfile
        elif endinglow in pptList:
            doconvert(word_pptcmd)
            if os.path.isfile(PDFfile):
                isPDFfile = True
                if convertToPDFA:
                    deleteTempFile = True           
            finalFile = PDFfile
        elif endinglow in xlsList:
            doconvert(word_xlscmd)        
            finalFile = ODSfile
        elif endinglow in pictureList:            
            doconvert(piccmd)        
            finalFile = PNGfile
            
        #Removes the original attachment if requested #SLS case keeps
        if deleteOriginalAttachment:        
            cleaner.removeFile(fullPath)
            
    else: #File that will not be handled
        finalFile = "Not supported"           
   
    
    """
    If requested convertion to PDF/A format, this part will do it
    Also removes the orginal PDF file if PDF/A file is found after conversion
    """
    if convertToPDFA and isPDFfile:
        #logwriter.printmessage("Requested conversion of {} to PDF/A".format(fullPath))        
        finalAttachmentFile = "-o"+os.path.join(attachpath, filebegin)+"_A.pdf"
        finalAttachmentFileClean = os.path.join(attachpath, filebegin)+"_A.pdf"        
       
                        
        pdfAcmd = [gs, '-dPDFA=3',
                   '-dBATCH', '-dNOPAUSE', '-dNOOUTERSAVE', '-dNOSAFER', '-dPDFSETTINGS=/prepress',
                   '-dPDFACompatibilityPolicy=1', '-dAutoFilterColorImages=false', '-dColorImageFilter=/FlateEncode',
                   '-dAutoFilterGrayImages=false', '-dGrayImageFilter=/FlateEncode', '-dMonoImageFilter=/FlateEncode',
                   '-dEmbedAllFonts=true', '-r1200', '-dNOTRANSPARENCY',
                   '-sDEVICE=pdfwrite', finalAttachmentFile, pdfa_def_path, PDFfile]      
        """
        pdfAcmd = [gs, '-dPDFA=3',
                   '-dBATCH', '-dNOPAUSE', '-dNOOUTERSAVE', '-dNOSAFER', '-dPDFSETTINGS=/prepress',
                   '-dPDFACompatibilityPolicy=1', '-dAutoFilterColorImages=false', '-dColorImageFilter=/FlateEncode',
                   '-dAutoFilterGrayImages=false', '-dGrayImageFilter=/FlateEncode', '-dMonoImageFilter=/FlateEncode',
                   '-dEmbedAllFonts=true', '-r1200', '-dNOTRANSPARENCY',
                   '-sDEVICE=pdfwrite', finalAttachmentFile, 'pdfa_def.ps', PDFfile]      
        
        """  
        
        attachmentProcess = subprocess.Popen(pdfAcmd)
        attachmentProcess.communicate()
        if os.path.isfile(finalAttachmentFileClean) and deleteTempFile:
            #logwriter.printmessage("PDF/A file found, lets delete plain PDF file {}".format(PDFfile))
            cleaner.removeFile(PDFfile) #Removes the temp PDF file if PDF/A is found 
            finalFile = finalAttachmentFileClean
    
    return finalFile


def outsideLandingPlace(item:str, convertToPDFA:bool=False, deleteOriginalAttachment:bool=False):
    """The start point of converter, called from outside function, item is either item name or full path to item"""
    #logwriter.printmessage("Called convert with: {}".format(item))
    attachpath = os.path.dirname(os.path.abspath(item))
    fullPath = os.path.join(attachpath, item)
    returnValue = ""
    #Checks if this truly is a file, if not returns an error
    if not os.path.isfile(fullPath):    
        returnValue = ("File {} not found..".format(fullPath))        
    else:
        returnValue = convertfile(fullPath, convertToPDFA, deleteOriginalAttachment)
    return returnValue
    """
    else:
        #print(fullPath.rsplit(os.path.sep,1))
        try:
            attachpath, filename = fullPath.rsplit(os.path.sep,1)    
            fileend = filename.rsplit('.',1)[1]
        except:
            logwriter.printmessage("File {} has no ending, trying to resolve it".format(fullPath))
            findcmd = ['file','--extension', fullPath]
            result = subprocess.run(findcmd, capture_output=True, text=True)
            logwriter.printmessage(result)
            
        
            
            
       
        if endinglow =='zip':        
            unzipPath = unzip(attachpath, fullPath)
            if unzipPath !="":
                #returnValue = zipfile.ZipFile(os.path.join(attachpath, filename+'-converted.zip'), 'a')        
                for root, dirs, files in os.walk(unzipPath):
                    for file in files:
                        fullPath = os.path.join(root, file)                
                        convertfile(fullPath, convertToPDFA, deleteOriginalAttachment)
                returnValue = shutil.make_archive(filebegin+"-converted",'zip',unzipPath)                
            else:
                returnValue = "Probably not a valid zip file..? Please check the file and try again"
            
        else:
        """
        
        #logwriter.printmessage("Outsidelandingplace:%s"%filetype)
    #logwriter.printmessage("FINAL return value = {}".format(returnValue))
    


