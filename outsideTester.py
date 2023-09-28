#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Mar 25 08:25:09 2019

@author: digitalia-aj
"""

import archivepst

if __name__ == "__main__":
    print("Starting from outside tester")
    fileflag = attachmentflag = False
    delflag  = zipflag = True
    dataDir = '/home/digitalia-aj/spostit/testi'
    dataFile = 'konferenssit.pst'
    
    finalZipFile = archivepst.mainActor(delflag, dataDir, fileflag, dataFile, zipflag, attachmentflag)
    
    print("Printing result outside caller, {}".format(finalZipFile))