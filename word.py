#!/usr/bin/env python  
# -*- coding: utf-8 -*-  
from win32com.client import Dispatch  
import win32com.client

class word(object):
    '''
    classdocs
    '''

    def __init__(self, filename=None):
        self.wordApp = win32com.client.Dispatch('Word.Application')       
        if filename:
            self.filename = filename       
            self.word = self.wordApp.Workbooks.Open(filename)
        else:        
            self.word = self.wordApp.Workbooks.Add()        
            self.filename = '' 
        