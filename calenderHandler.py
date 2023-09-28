#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Feb 10 14:44:03 2020

@author: digitalia-aj
"""
import os
import convertfiles

def handleCalender(path):
    print(path)
    return


if __name__ == "__main__":
    walk_dir = os.path.dirname(os.path.abspath(__file__))
    calPaths = []
    for root, dirs, files in os.walk(walk_dir):
        for file in files:
            if file =="Meeting.txt" or file == "Appointment.txt":
                calPaths.append(root)
        
    for onePath in calPaths:
        handleCalender(onePath)
        pathContents = os.listdir(onePath)        
        if "Message.rtf" in pathContents:
            print ("rtf file found in {}".format(onePath))
            convertfiles.outsideLandingPlace(os.path.join(onePath, "Message.rtf"), True)


        