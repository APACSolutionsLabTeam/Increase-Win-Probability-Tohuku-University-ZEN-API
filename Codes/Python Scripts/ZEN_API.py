# -*- coding: utf-8 -*-
"""
Created on Mon Dec 14 09:23:11 2020

@author: ZSPANIYA
"""
import win32com.client
import tifffile as tf
import matplotlib.pyplot as plt

## Import the ZEN OAD Scripting into Python
Zen = win32com.client.GetActiveObject("Zeiss.Micro.Scripting.ZenWrapperLM")

## Remove all open ZEN documents
Zen.Application.Documents.RemoveAll(False)

## Execute an Z-stack experiment in ZEN
experimentDemo = Zen.Acquisition.Experiments.GetByName("ExperimentDemo.czexp")
imageDemo = Zen.Acquisition.Execute(experimentDemo)

## Control microscope hardware using ZEN API
initialFocus = Zen.Acquisition.FindSurface()# initial Find Focus
Zen.Devices.Stage.TargetPositionY = 59620
Zen.Devices.Stage.TargetPositionX = 93240
Zen.Devices.Stage.Apply()
Zen.Application.Wait(1000)
softwareAutoFocus = Zen.Acquisition.FindAutofocus(30)
Zen.Devices.Stage.MoveTo(0.0,0.0)

hws=Zen.Devices.ReadHardwareSetting()
hwsid=hws.GetAllComponentIds()
print(hwsid)
for obj in hwsid:
    hwsp=hws.GetAllParameterNames(obj)
    print(obj,hwsp)





























