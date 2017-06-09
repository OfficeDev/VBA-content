---
title: Application.WindowHandle32 Property (Visio)
keywords: vis_sdr.chm10014660
f1_keywords:
- vis_sdr.chm10014660
ms.prod: visio
api_name:
- Visio.Application.WindowHandle32
ms.assetid: d4c653ae-6582-0d86-75ee-969fe978e754
ms.date: 06/08/2017
---


# Application.WindowHandle32 Property (Visio)

Returns the 32-bit handle of a Microsoft Visio window. Read-only.


## Syntax

 _expression_ . **WindowHandle32**

 _expression_ A variable that represents an **Application** object.


### Return Value

Long


## Remarks

The  **WindowHandle32** property of an **Application** object returns one of the following:




- The  **HWND** for the main Visio (frame) window (most common).
    
- The  **HWND** for the container application's main frame window if Visio is running in-place and active.
    
- The  **HWND** for the window returned by the **GetActiveWindow** () function if either frame window is disabled (for example, if a modal dialog box is running). For details about the **GetActiveWindow** function, see the Microsoft Platform SDK on the Microsoft Developer Network (MSDN) Web site.
    


Use the  **WindowHandle32** property of the **Window** object to obtain the **HWND** for a window in the **Windows** collection of a Visio instance.

You can use the obtained  **HWND** in Windows API calls.


 **Note**  Calls to the  **WindowHandle** property (now hidden) are directed to the **WindowHandle32** property.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVApplication.WindowHandle**
    
-  **Microsoft.Office.Interop.Visio.IVApplication.WindowHandle32**
    

## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to get the 32-bit handle of a window.


```vb
 
Public Sub WindowHandle32_Example()  
 
    Dim vsoWindow As Visio.Window  
    Dim lngWindowHandle32 As Long  
 
    'Get the active window. 
    Set vsoWindow = ActiveWindow  
 
    'Get the 32-bit handle of the active window.  
    lngWindowHandle32 = vsoWindow.WindowHandle32  
 
    'Verify that you got the handle.  
    Debug.Print "The active window handle is"; lngWindowHandle32  
 
End Sub
```


