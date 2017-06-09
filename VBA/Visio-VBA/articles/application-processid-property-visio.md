---
title: Application.ProcessID Property (Visio)
keywords: vis_sdr.chm10014145
f1_keywords:
- vis_sdr.chm10014145
ms.prod: visio
api_name:
- Visio.Application.ProcessID
ms.assetid: d089bfa9-83a4-1b44-80ab-f23c5198801f
ms.date: 06/08/2017
---


# Application.ProcessID Property (Visio)

Returns the unique identity of the current Microsoft Visio process. Read-only.


## Syntax

 _expression_ . **ProcessID**

 _expression_ A variable that represents an **Application** object.


### Return Value

 **Long**


## Remarks

The  **ProcessID** property returns a value unique to the indicated instance. The application doesn't reuse the value until 4294967296 (2^32) more threads have been created on the current workstation.


 **Important**  The value returned by  **ProcessID** is not the same as the Windows Process ID of the current Visio instance.


## Example

This Microsoft Visual Basic for Applications (VBA) program shows how to use the  **ProcessID** property to determine the unique identity of the current Microsoft Visio process.


```vb
Sub ProcessID_Example () 
 
    Dim vsoApplication As Visio.Application 
 
    'Get the current instance of Microsoft Office Visio. 
    Set vsoApplication = Visio.Application 
 
    'Prints the unique ID of the current Visio process.  
    Debug.Print "Visio Process identifier: "; vsoApplication.ProcessID 
 
End Sub
```


