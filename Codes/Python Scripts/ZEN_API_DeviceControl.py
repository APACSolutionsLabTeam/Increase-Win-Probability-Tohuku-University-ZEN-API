# -*- coding: utf-8 -*-
"""
Created on Tue Dec 15 15:03:56 2020

@author: ZSPANIYA
"""
## Import required libraries
import win32com.client
import tifffile as tf

## Import the ZEN OAD Scripting into Python
Zen = win32com.client.GetActiveObject("Zeiss.Micro.Scripting.ZenWrapperLM")

## Control microscope hardware using ZEN API
initialFocus = Zen.Acquisition.FindSurface()# initial Find Focus
Zen.Devices.Stage.TargetPositionY = 59620
Zen.Devices.Stage.TargetPositionX = 93240
Zen.Devices.Stage.Apply()
Zen.Application.Wait(1000)
softwareAutoFocus = Zen.Acquisition.FindAutofocus(30)
Zen.Devices.Stage.MoveTo(0.0,0.0)
