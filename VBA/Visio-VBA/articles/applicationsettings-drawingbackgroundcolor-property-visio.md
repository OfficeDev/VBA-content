---
title: ApplicationSettings.DrawingBackgroundColor Property (Visio)
keywords: vis_sdr.chm16251810
f1_keywords:
- vis_sdr.chm16251810
ms.prod: visio
api_name:
- Visio.ApplicationSettings.DrawingBackgroundColor
ms.assetid: c07d8268-d0f6-afc7-8c6f-da16a3f643a0
ms.date: 06/08/2017
---


# ApplicationSettings.DrawingBackgroundColor Property (Visio)

Determines the background color of the Microsoft Visio drawing window for the current session. Read/write.


## Syntax

 _expression_ . **DrawingBackgroundColor**

 _expression_ A variable that represents an **ApplicationSettings** object.


### Return Value

OLE_COLOR


## Remarks

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
Setting the  **BackgroundColor** property of the active window to a value other than the default (-1) overrides the **DrawingBackgroundColor** setting for that window. To be able to reset the background color of the same active window by setting the **DrawingBackgroundColor** property, you must reset **BackgroundColor** to its default value, -1. If multiple windows are open, setting **BackgroundColor** for one window has no effect on the setting for other open windows.




 **Note**  The ability to set drawing background color programmatically for users running in high-contrast mode is restricted.


## Example

The following VBA macro shows how to use the  **DrawingBackgroundColor** property to get and set the application background color. It also shows how to get an **ApplicationSettings** object from the Visio **Application** object, and it demonstrates the relationship between the **DrawingBackgroundColor** property and the **Window.BackgroundColor** property. This example assumes there is a drawing window open in Visio and that initially all background color properties are set to their default values.


```vb
Public Sub DrawingBackgroundColor_Example() 
 
 Dim vsoApplicationSettings As Visio.ApplicationSettings 
 Set vsoApplicationSettings = Visio.Application.Settings 
 
 'Get the current application background color. 
 Debug.Print vsoApplicationSettings.DrawingBackgroundColor 
 
 'Get the active window background color. 
 Debug.Print ActiveWindow.BackgroundColor 
 
 'Change the application background color. 
 'This will also change the active window color as 
 'well as the setting in the Color Settings dialog box 
 vsoApplicationSettings.DrawingBackgroundColor = vbRed 
 
 'Change the active window background color. 
 ActiveWindow.BackgroundColor = vbMagenta 
 
 'Change the application background color again. 
 'This time, there is no change in the current 
 'window color, but the dialog box setting changes. 
 vsoApplicationSettings.DrawingBackgroundColor = vbYellow 
 
 'Reset Window.BackgroundColor to its default value. 
 ActiveWindow.BackgroundColor = -1 
 
 'Change the application background color again. 
 'Now both the active window color 
 'and the dialog box setting change. 
 vsoApplicationSettings.DrawingBackgroundColor = vbBlue 
 
End Sub
```


