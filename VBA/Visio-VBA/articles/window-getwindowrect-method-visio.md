---
title: Window.GetWindowRect Method (Visio)
keywords: vis_sdr.chm11616330
f1_keywords:
- vis_sdr.chm11616330
ms.prod: visio
api_name:
- Visio.Window.GetWindowRect
ms.assetid: 272714c6-3502-4baa-5006-2dcec8c0dfbd
ms.date: 06/08/2017
---


# Window.GetWindowRect Method (Visio)

Gets the size and position of the client area of a window.


## Syntax

 _expression_ . **GetWindowRect**( **_pnLeft_** , **_pnTop_** , **_pnWidth_** , **_pnHeight_** )

 _expression_ A variable that represents a **Window** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pnLeft_|Required| **Long**|The coordinate of the left side of the window.|
| _pnTop_|Required| **Long**|The coordinate of the top of the window.|
| _pnWidth_|Required| **Long**|The distance in pixels from the left side to the right side of the window.|
| _pnHeight_|Required| **Long**|The distance in pixels from the top to the bottom of the window.|

### Return Value

Nothing


## Remarks

The  **GetWindowRect** method gets the size and position of the client area of the window with respect to the window that owns the **Windows** collection to which it belongs. For the **Windows** collection of an **Application** object, the "with respect to" window is the MDICLIENT window of the Microsoft Visio main window. For the **Windows** collection of a **Window** object, the "with respect to" window is the client area of the drawing window.


## Example

The following example shows how to use the  **GetWindowRect** method to get the size and position of a **Window** object. It opens the **Pan &; Zoom** window and prints the window's coordinates, width, and height in the Immediate window.


```vb
Public Sub GetWindowRect_Example() 
 
 Dim vsoApplication As Visio.Application 
 Dim vsoPZWindow As Visio.Window 
 Dim pinLeft As Long, pinTop As Long, pinWidth As Long, pinHeight As Long 
 
 Set vsoApplication = Visio.Application 
 
 'Display the Pan &; Zoom window 
 Set vsoPZWindow = vsoApplication.ActiveWindow.Windows.ItemFromID(visWinIDPanZoom) 
 vsoPZWindow.Visible = True 
 
 'Get the existing window size and position 
 vsoPZWindow.GetWindowRect pinLeft, pinTop, pinWidth, pinHeight 
 Debug.Print pinLeft, pinTop, pinWidth, pinHeight 
 
 'Change the window height and get the new values 
 vsoPZWindow.SetWindowRect pinLeft, pinTop, pinWidth, pinHeight + 50 
 vsoPZWindow.GetWindowRect pinLeft, pinTop, pinWidth, pinHeight 
 Debug.Print pinLeft, pinTop, pinWidth, pinHeight 
 
End Sub
```


