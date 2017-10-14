---
title: InvisibleApp.ProcessID Property (Visio)
keywords: vis_sdr.chm17514145
f1_keywords:
- vis_sdr.chm17514145
ms.prod: visio
api_name:
- Visio.InvisibleApp.ProcessID
ms.assetid: ae20f8f4-3236-68c7-161c-fdc87f68e7ca
ms.date: 06/08/2017
---


# InvisibleApp.ProcessID Property (Visio)

Returns the unique identity of the current Microsoft Visio process. Read-only.


## Syntax

 _expression_ . **ProcessID**

 _expression_ A variable that represents an **InvisibleApp** object.


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


