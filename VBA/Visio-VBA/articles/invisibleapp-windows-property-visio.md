---
title: InvisibleApp.Windows Property (Visio)
keywords: vis_sdr.chm17514665
f1_keywords:
- vis_sdr.chm17514665
ms.prod: visio
api_name:
- Visio.InvisibleApp.Windows
ms.assetid: acabb5fc-b55b-b3d4-032d-db8947eb7a12
ms.date: 06/08/2017
---


# InvisibleApp.Windows Property (Visio)

Returns the  **Windows** collection for a Microsoft Visio instance or window. Read-only.


## Syntax

 _expression_ . **Windows**( **_lpdispRet_** )

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

Windows


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


