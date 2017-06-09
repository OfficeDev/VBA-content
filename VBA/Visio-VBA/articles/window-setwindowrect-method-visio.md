---
title: Window.SetWindowRect Method (Visio)
keywords: vis_sdr.chm11616590
f1_keywords:
- vis_sdr.chm11616590
ms.prod: visio
api_name:
- Visio.Window.SetWindowRect
ms.assetid: f9f24c79-9c8f-ec0d-f894-1c10150db75e
ms.date: 06/08/2017
---


# Window.SetWindowRect Method (Visio)

Sets the size and position of the client area of a window.


## Syntax

 _expression_ . **SetWindowRect**( **_nLeft_** , **_nTop_** , **_nWidth_** , **_nHeight_** )

 _expression_ A variable that represents a **Window** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _nLeft_|Required| **Long**|The coordinate of the left side of the window.|
| _nTop_|Required| **Long**|The coordinate of the top of the window.|
| _nWidth_|Required| **Long**|The distance in pixels from the left side to the right side of the window.|
| _nHeight_|Required| **Long**|The distance in pixels from the top to the bottom of the window.|

### Return Value

Nothing


## Remarks

The  **SetWindowRect** method sets the size and position of the client area of the window with respect to the window that owns the **Windows** collection to which it belongs. For the **Windows** collection of an **Application** object, the "with respect to" window is the MDICLIENT window of the Visio main window. For the **Windows** collection of a **Window** object, the "with respect to" window is the client area of the drawing window.

 **SetWindowRect** has no effect when the window is docked.


## Example

The following example shows how to use the  **SetWindowRect** method to set the size and position of a **Window** object. It opens the **Pan &; Zoom** window and prints the window's coordinates, width, and height in the **Immediate** window. Then it uses **SetWindowRect** to change the height of the window, and prints the new values.


```vb
Public Sub SetWindowRect_Example() 
 
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


