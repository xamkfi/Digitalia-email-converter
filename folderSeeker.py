#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Apr 17 11:41:52 2019

@author: digitalia-aj
"""
import os
import logwriter
exclude = ['alarms', 'drafts']

def findAllEmptyDirs(startPath):
    emptyDirs = []
    for (root, dirs, files) in os.walk(startPath):
        if len(dirs) == 0 and len(files) == 0 :
            emptyDirs.append(root)    
    return emptyDirs

def findAllContentDirs(startPath):
    messages = []
    appointments = []
    meetings = []
    for root, dirs, files in os.walk(startPath):
        for fileName in files:
            #logwriter.printmessage("Root {}".format(root))                               
            if any(exc in str(root).lower() for exc in exclude):
                logwriter.printmessage("Excluding {}".format(root))
            else:
                fileName = str(fileName).lower()
                dirname = os.path.dirname(os.path.join(root, fileName))
                if "message" in fileName:                
                    if dirname not in messages:
                        messages.append(dirname)                
                if "appointment" in fileName:
                    if dirname not in appointments:
                        appointments.append(dirname)                
                if "meeting" in fileName:
                    if dirname not in meetings:
                        meetings.append(dirname)   
    return messages, appointments, meetings