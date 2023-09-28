#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Jun 15 08:52:00 2020

@author: digitalia-aj
"""

import convertfiles
import argparse

if __name__ == "__main__":               
    ap = argparse.ArgumentParser()
    ap.add_argument("-f", required=True, help="item name")
    ap.add_argument("-del", action='store_true', help="If given, destroys the original file after archival version is created")
    ap.add_argument("-a", action='store_true', help="If given, converts to PDF/A instead of PDF ")
    args = vars(ap.parse_args())
    fullPath  = args['f']
    pdfA = args['a']
    delete = args['del']
    
    convertfiles.outsideLandingPlace(fullPath, pdfA, delete)