---
title: Window.BackgroundColorGradient Property (Visio)
keywords: vis_sdr.chm11651910
f1_keywords:
- vis_sdr.chm11651910
ms.prod: visio
api_name:
- Visio.Window.BackgroundColorGradient
ms.assetid: a23e1075-9a3f-e04a-c6eb-8e4d983b8970
ms.date: 06/08/2017
---


# Window.BackgroundColorGradient Property (Visio)

Determines the background gradient color of the active Microsoft Visio drawing window and its associated print preview and full-screen view windows. Read/write.


## Syntax

 _expression_ . **BackgroundColorGradient**

 _expression_ A variable that represents a **Window** object.


### Return Value

OLE_COLOR


## Remarks

Valid values for an  **OLE_COLOR** property within Visio can be one of the following:




- &;H00 _bbggrr,_ where _bb_ is the blue value between 0 and 0xFF (255), _gg_ the green value, and _rr_ the red value.
    
- &;H800000 _xx_ , where _xx_ is a valid **GetSysColor** index.
    


For details about the  **GetSysColor** function, search for " **GetSysColor** " in the Microsoft Platform SDK on MSDN.

The  **OLE_COLOR** data type is used for properties that return colors. When a property is declared as **OLE_COLOR** , the **Properties**window in the Visual Basic Editor editor displays a color-picker dialog box that allows the user to select the color for the property visually, rather than having to remember the numeric equivalent.

In addition, you can use the following Microsoft Visual Basic for Applications (VBA) color constants for  **OLE_COLOR** .



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **vbBlack**|0x0|Black|
| **vbRed**|0xFF|Red|
| ** vbGreen**|0xFF00|Green|
| **vbYellow**|0xFFFF|Yellow|
| **vbBlue**|0xFF0000 |Blue|
| ** vbMagenta**|0xFF00FF |Magenta|
| ** vbCyan**|0xFFFF00|Cyan|
| ** vbWhite**|0xFFFFFF|White|
For a drawing window, print preview window, or full-screen view window, setting the  **BackgroundColorGradient** property of the window to a value other than the default (-1) overrides the **ApplicationSettings.DrawingBackgroundColorGradient** setting for that set of windows. To be able to reset this property for the same set of windows, you must reset **BackgroundColorGradient** to its default value, -1. If multiple drawings are open, setting **BackgroundColorGradient** for one set of windows has no effect on the setting for other open window sets.

However, for a ShapeSheet window, setting the  **BackgroundColorGradient** property of the window to a value other than the default has no effect, either on the background color of the ShapeSheet window, or on any of the other window settings.


## Example

The following VBA macro shows how to use the  **BackgroundColorGradient** property to get and set the active window background gradient color. It also shows how to get an **ApplicationSettings** object from the Visio **Application** object, and it demonstrates the relationship between the **BackgroundColorGradient** property and the **ApplicationSettings.DrawingBackgroundColorGradient** property. This example assumes there is an active drawing window open in Visio and that initially all background gradient color properties are set to their default values.


```vb
Public Sub BackgroundColorGradient_Example() 
 
 Dim vsoApplicationSettings As Visio.ApplicationSettings 
 Set vsoApplicationSettings = Visio.Application.Settings 
 
 'Get the current application background gradient color. 
 Debug.Print vsoApplicationSettings.DrawingBackgroundColorGradient 
 
 'Get the current window background color gradient. 
 Debug.Print ActiveWindow.BackgroundColorGradient 
 
 'Change the application background gradient color. 
 'This also changes the current window color as 
 'well as the setting in the Color Settings dialog box. 
 vsoApplicationSettings.DrawingBackgroundColor = vbRed 
 
 'Change the active window background gradient color. 
 ActiveWindow.BackgroundColorGradient = vbMagenta 
 
 'Change the application background gradient color again. 
 'This time, there is no change in the current 
 'window color, but the dialog box setting changes. 
 vsoApplicationSettings.DrawingBackgroundColorGradient = vbYellow 
 
 'Reset Window.BackgroundColorGradient to its default value. 
 ActiveWindow.BackgroundColorGradient = -1 
 
 'Change the application background gradient color again. 
 'Now both the active window color 
 'and the dialog box setting change. 
 vsoApplicationSettings.DrawingBackgroundColorGradient = vbBlue 
 
End Sub
```


