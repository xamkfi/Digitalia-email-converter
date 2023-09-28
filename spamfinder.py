#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Dec 13 09:11:58 2018

@author: digitalia-aj
"""
import os
#import multiprocessing
#from multiprocessing import Pool

from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument

walkDir = os.getcwd()
#cpuCount = multiprocessing.cpu_count()
#pool = Pool(cpuCount)

text3 = "date format for"
#text1 = "GRAM-teemaseminaari"
#text2 = ""

def printmessage(message):    
    print(message)
    return

def multiprocess_results(retvalue):
    printmessage(retvalue)

def checkPDFMetadata(root, file, path):
    if os.path.isfile(os.path.join(root, "ModifiedCombinedHeaders.txt")):
        freader = open(os.path.join(root, "ModifiedCombinedHeaders.txt"), "r")
        data = freader.read()
        if text3 in data:            
            printmessage("File: {}\n".format(path))
        
    else:
        fp = open(path, 'rb')
        parser = PDFParser(fp)
        doc = PDFDocument(parser)
        metadata = str(doc.info)  # The "Info" metadata
        if text3 in metadata:
            printmessage("File: {}\n".format(path))
            printmessage(metadata)
        
    
def multiProcessCheckFile(root, file):
    #print("nuuh")
    #printmessage(file)
    if os.path.isfile(os.path.join(root, "ModifiedCombinedHeaders.txt")):
        freader = open(os.path.join(root, "ModifiedCombinedHeaders.txt"), "r")
        data = freader.read()
        if text3 in data:            
            return root
        else:
            return "NA"
    

notFoundCount = 0
for root, dirs, files in os.walk(walkDir):
    pdfAFound = False
    for file in files:        
        if str(file).endswith('_A.pdf'):                        
            #printmessage("{}".format(root))
            pdfAFound = True
            checkPDFMetadata(root, file, os.path.join(root, file))
            #pool.apply_async(checkPDFMetadata, args=(os.path.join(root,file)), callback=multiprocess_results)
            #pool.apply_async(multiProcessCheckFile, args=(root, file), callback=multiprocess_results)
            #pool.apply_async(multipath, args=(root, metaname, beginfilename), callback=multipathResults)
            #multiProcessCheckFile(pdffile)
    if pdfAFound == False:
        if 'Message' in str(root) and 'Attachment' not in str(root):
            printmessage("No _A file in {}".format(root))
            notFoundCount+=1
            
printmessage("No pdfA file in {} folder".format(notFoundCount))
#pool.close()
#pool.join()            
