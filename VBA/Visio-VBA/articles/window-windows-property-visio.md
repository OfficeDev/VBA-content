---
title: Window.Windows Property (Visio)
keywords: vis_sdr.chm11614665
f1_keywords:
- vis_sdr.chm11614665
ms.prod: visio
api_name:
- Visio.Window.Windows
ms.assetid: 6e063a03-71e5-d2e2-d9d0-38fcae604d26
ms.date: 06/08/2017
---


# Window.Windows Property (Visio)

Returns the  **Windows** collection for a Microsoft Visio instance or window. Read-only.


## Syntax

 _expression_ . **Windows**

 _expression_ A variable that represents a **Window** object.


### Return Value

Windows


## Remarks

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVWindow.Windows**
    

## Example

This Microsoft Visual Basic macro gets the  **Windows** collection of the **Application** object and prints the ID of each window in the collection in the Immediate window.


```vb
Public Sub Windows_Example() 
  
    Dim vsoApplication As Visio.Application  
    Dim vsoWindows As Visio.Windows 
    Dim intCounter As Integer  
 
    'Get the Windows collection.  
    Set vsoApplication = Application  
    Set vsoWindows = vsoApplication.Windows 
 
    For intCounter = 1 To vsoWindows.Count 
        Debug.Print vsoWindows.Item(intCounter).ID 
    Next intCounter  
 
End Sub
```


