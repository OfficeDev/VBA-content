---
title: Master.OpenIconWindow Method (Visio)
keywords: vis_sdr.chm10716410
f1_keywords:
- vis_sdr.chm10716410
ms.prod: visio
api_name:
- Visio.Master.OpenIconWindow
ms.assetid: 5e2b2437-05cc-4855-e0bb-96b097c98d3c
ms.date: 06/08/2017
---


# Master.OpenIconWindow Method (Visio)

Opens an icon window that shows a master's icon.


## Syntax

 _expression_ . **OpenIconWindow**

 _expression_ A variable that represents a **Master** object.


### Return Value

Window


## Remarks

If the master's icon is already displayed in an icon window, the  **OpenIconWindow** method activates that window rather than opening another window.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **OpenIconWindow** method to open an icon editing window.


```vb
 
Public Sub OpenIconWindow_Example() 
 
 Dim vsoMaster As Visio.Master 
 Dim vsoIconWindow As Visio.Window 
 
 'Add a master to the document stencil and open its icon editing window. 
 Set vsoMaster = ThisDocument.Masters.Add 
 Set vsoIconWindow = vsoMaster.OpenIconWindow 
 
End Sub
```


