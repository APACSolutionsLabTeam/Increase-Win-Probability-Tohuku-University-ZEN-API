# -*- coding: utf-8 -*-
"""
Created on Tue Dec 15 15:02:54 2020

@author: ZSPANIYA
"""
## Import required library
import win32com.client

## Import the ZEN OAD Scripting into Python
Zen = win32com.client.GetActiveObject("Zeiss.Micro.Scripting.ZenWrapperLM")

## Remove all open ZEN documents
Zen.Application.Documents.RemoveAll(False)

## Execute an Z-stack experiment in ZEN
experimentDemo = Zen.Acquisition.Experiments.GetByName("ExperimentDemo.czexp")
imageDemo = Zen.Acquisition.Execute(experimentDemo)




