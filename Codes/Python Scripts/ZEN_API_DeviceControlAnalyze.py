# -*- coding: utf-8 -*-
"""
Created on Tue Dec 15 16:25:24 2020

@author: ZSPANIYA
"""

## Import required libraries
import win32com.client
import cv2
import matplotlib.pyplot as plt
import numpy as np

## Import the ZEN OAD Scripting into Python
Zen = win32com.client.GetActiveObject("Zeiss.Micro.Scripting.ZenWrapperLM")

## Remove all open ZEN documents
Zen.Application.Documents.RemoveAll(False)

#initialFocus = Zen.Acquisition.FindSurface()# initial Find Focus
X_pos = Zen.Devices.Stage.ActualPositionX
Y_pos = Zen.Devices.Stage.ActualPositionY
x1 = 0
y1 = 0
 
for i in range(0, 5):
    
    #softwareAutoFocus = Zen.Acquisition.FindAutofocus(30)
    
    experimentDemo_1 = Zen.Acquisition.Experiments.GetByName("ExperimentDemo_1.czexp")
    imageDemo = Zen.Acquisition.Execute(experimentDemo_1)
    
    ## Save the image in an externl folder for python to access
    imageDemo.Save_2(r'C:/temp/imageDemo.tif')
    imageDemoRead=cv2.imread(r'C:/temp/imageDemo.tif')
    
    fig1 = plt.figure() 
    plt.imshow(imageDemoRead)
    
    ## Image Processing Using Python
    imageSmoothedGray = cv2.cvtColor(imageDemoRead, cv2.COLOR_BGR2GRAY)
    imageSmoothed = cv2.GaussianBlur(imageSmoothedGray,(15,15),0)
    imageSmoothedFloat = imageSmoothed.astype(np.float64)
    imageSmoothedCorrected = 255*(imageSmoothedFloat/np.max(imageSmoothedFloat))
    imageSmoothedCorrected_1 = imageSmoothedCorrected.astype(np.uint8)
    fig2 = plt.figure()
    plt.imshow(imageSmoothedCorrected_1, 'gray')
    
    Zen.Devices.Stage.MoveTo(X_pos+x1, Y_pos+y1 )
    x1 = x1+50
    y1 = y1+50

    
    
