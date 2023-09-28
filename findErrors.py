#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Feb 19 08:32:51 2019
Reads the content of pla-mail.exp.txt which contains errors
and prints the content of attachments folder
Some or few files in that folder has names incompatible with Windows

@author: digitalia-aj
"""

"""Loads error file"""
def loadStopWords():
    print("Loading error file")
    #result = ""
    #i = 0
    #reader = open('/home/user/stopwords/stopwords.txt', 'r')
    reader = open('pla-mail-exp.txt', 'r')
    #reader = open("stopwords.txt", "r")
    print("Loading done")
    stopwordsread = reader.read().splitlines()
    print("Reading words done")
    print(stopwordsread)
    return stopwordsread
        


"""Main"""
print("Staring execution")
loadStopWords()