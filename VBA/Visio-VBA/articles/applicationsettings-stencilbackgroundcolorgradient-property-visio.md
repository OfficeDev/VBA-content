---
title: ApplicationSettings.StencilBackgroundColorGradient Property (Visio)
keywords: vis_sdr.chm16251530
f1_keywords:
- vis_sdr.chm16251530
ms.prod: visio
api_name:
- Visio.ApplicationSettings.StencilBackgroundColorGradient
ms.assetid: e73b2f5a-6ddf-0e46-62f1-8409e7e0608c
ms.date: 06/08/2017
---


# ApplicationSettings.StencilBackgroundColorGradient Property (Visio)

Determines the background gradient color of the Microsoft Visio stencil window for the current session. Read/write. 


## Syntax

 _expression_ . **StencilBackgroundColorGradient**

 _expression_ A variable that represents an **ApplicationSettings** object.


### Return Value

OLE_COLOR


## Remarks

The  **StencilBackgroundColorGradient** property setting does not persist from one session of Visio to the next.

Valid values for an  **OLE_COLOR** property within Visio can be one of the following:




- &;H00 _bbggrr,_ where _bb_ is the blue value between 0 and 0xFF (255), _gg_ the green value, and _rr_ the red value.
    
- &;H800000 _xx_ , where _xx_ is a valid **GetSysColor** index.
    


For details about the  **GetSysColor** function, search for " **GetSysColor** " in the Microsoft Platform SDK on MSDN.

The  **OLE_COLOR** data type is used for properties that return colors. When a property is declared as **OLE_COLOR** , the **Properties** window displays a color-picker dialog box that allows the user to select the color for the property visually, rather than having to remember the numeric equivalent.

In addition, you can use the following Microsoft Visual Basic for Applications (VBA) color constants for  **OLE_COLOR** .



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **vbBlack**|0x0 |Black|
| **vbRed**|0xFF |Red|
| ** vbGreen**|0xFF00 |Green|
| **vbYellow**|0xFFFF|Yellow|
| **vbBlue**|0xFF0000 |Blue|
| ** vbMagenta**|0xFF00FF |Magenta|
| ** vbCyan**|0xFFFF00|Cyan|
| ** vbWhite**|0xFFFFFF|White|
You can set the background gradient color for an individual stencil window by using the  **ActiveWindow.Windows(**_n_**).BackgroundColorGradient** property, where _n_ represents the index number of the stencil window in the **Windows** collection. When a stencil window is opened as a floating window, it can be an active window, and you can set its background gradient color by using the **ActiveWindow.BackgroundColorGradient** property. If you have opened a stencil window in this manner, setting the **BackgroundColor** property of the active stencil window to a value other than the default (-1) overrides the **StencilBackgroundColorGradient** setting for that window. To be able to reset the background gradient color of the same active stencil window by setting the **StencilBackgroundColorGradient** property, you must reset **BackgroundColorGradient** to its default value, -1. If multiple stencil windows of this type are open, setting **BackgroundColorGradient** for one window has no effect on the setting for other open windows.




 **Note**  You can specify two colors for the stencil background. If users' screen resolution is adequate, one of the colors will grade into the other from the top to the bottom of the screen. To be able to use this feature, users must set their monitors to display 32-bit color. The ability to set stencil background color programmatically for users running in high-contrast mode is restricted.


