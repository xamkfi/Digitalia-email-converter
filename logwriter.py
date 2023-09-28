#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Sep 20 08:13:38 2017

Log writer module

@author: digitalia-aj
"""
import time


"""
This function prints the messages to console as well as
into predefined logfile
"""

#firstcall = True
logfile = open('conversion.log', 'w', 1) #1 should tell to write line by line without buffer

def printmessage(message):
    #global firstcall
    global logfile
    #if firstcall:        
    #    firstcall = False
    #printmessage(time.strftime("%H:%M:%S", time.localtime()))    
    try:
        print(message)
        logfile.write(time.strftime("%H:%M:%S", time.localtime())+">"+str(message)+"\n")
    except OSError:
        pass
    return

def closeLogFile():
    global logfile
    logfile.close