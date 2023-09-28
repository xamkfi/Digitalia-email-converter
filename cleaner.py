#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Sep 20 09:21:50 2017

Cleaner module, moves files, renames, deletes folders etc.

@author: digitalia-aj
"""
development = True #Api use requires from . import
if development:
    import logwriter
else:
    from . import logwriter

import shutil
import os


def stripFileName(beginfilename, datadir):
    beginfilenamestriped = beginfilename.replace(" ","_")
    if beginfilename != beginfilenamestriped:
        os.rename(os.path.join(datadir,beginfilename), os.path.join(datadir,beginfilenamestriped))
    return beginfilenamestriped

"""Called at the beginning to ensure that
the there aren't directory that we are going to form
"""
def removeDir(dirname): #actual method that removes the folder
    try:        
        shutil.rmtree(dirname)
        #logwriter.printmessage ("Removed --> {}".format(dirname))
    except OSError:
        logwriter.printmessage ("ERROR in removing {}".format(dirname))
        pass
    return

"""Called if the delflag is set to True with cli param
Removes given fullpath file
"""
def removeFile(fullPath): #actual method that remove files
    try:        
        #logwriter.printmessage ("DEL: {}".format(fullPath))
        os.remove(fullPath)

    except OSError:
        pass
    return

def handleFinalCleanup(exportPath):
    logwriter.printmessage("Final cleanup for = {}".format(exportPath))        
    deletedThings = 0            
    try:
        for root, dirs, files in os.walk(exportPath):
            """Seeks attachment folders with content but empty parent --> delete"""
            for dir in dirs:
                if dir == "Attachments":
                    dirFiles = []
                    for path in os.scandir(os.path.join(root)):
                        if path.is_file():
                            dirFiles.append(path.name)
                    
                    Afound = False;
                    for onefile in dirFiles:
                        if "_A.pdf" in onefile:
                            Afound = True
                            
                    if Afound == False:
                        removeDir(os.path.join(root,dir))
                        #removeDir(os.path(root))                    
                            
            """Deletes all other files except the final _A.pdf file"""
            for file in files:                    
                if not str(root).endswith("Attachments"): #Folder but not attachment folder
                    if not str(file).endswith("_A.pdf"): #delete all except final _A.pdf file
                        #orgfile = os.path.join(completedir, file)
                        #moveOneUp(orgfile, file)
                        #if zipflag:
                        #    moveFileToUpperFolder(orgfile, file, root)        
                        #logwriter.printmessage("Delete: {}".format(os.path.join(root, file)))
                        removeFile(os.path.join(root, file))
                        deletedThings+=1                    
                elif str(root).endswith("Attachments"):#This happens inside Attachments folder                    
                    if "_Attachment.txt" in str(file):
                        removeFile(os.path.join(root, file))
                        deletedThings+=1
                            
            
        for root, dirs, files in os.walk(exportPath):
            for dirname in dirs:
                completedir = os.path.join(root,dirname)                                                           
                #logwriter.printmessage("dirname ="+completedir)
                       
                if not os.listdir(completedir):
                    #logwriter.printmessage("Removing empty dir {}".format(completedir))
                    removeDir(completedir)
                    deletedThings+=1
        logwriter.printmessage("Directory clean-up completed, deleted {} items".format(deletedThings))            
                           
        
    except OSError:
        logwriter.printmessage("OSError "+str(OSError.strerror))
        pass
    
    return deletedThings

"""
def moveOneUp(filepath, fname):
    logwriter.printmessage("MoveOneUp")
    destPath = filepath.split("Message")[0]
    logwriter.printmessage("destination path "+str(destPath))
    os.rename(filepath, destPath+"/"+fname)
    return
"""

def moveFileToUpperFolder(orgfile, fname, root):
    # move file to upper folder    
    destPath = root.rsplit('/', 1)[0]
    #logwriter.printmessage("POLUT "+str(destPath)+" "+str(orgfile))
    os.rename(orgfile, destPath+"/"+fname)
    return

def handleEmptyDir(path):
    if not os.listdir(path):           
        removeDir(path)    
    return

def renameFile(origfile, renamedFile):
    logwriter.printmessage("Renaming:"+str(origfile)+" into "+str(renamedFile))
    os.replace(origfile, renamedFile)
    if os.path.isfile(renamedFile):
        return
   

def handleListofEmptyDirs(pathList):
    for onepath in pathList:
        handleEmptyDir(onepath)
    return
