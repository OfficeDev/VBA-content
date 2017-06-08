---
title: InvisibleApp.ShowStatusBar Property (Visio)
keywords: vis_sdr.chm17514380
f1_keywords:
- vis_sdr.chm17514380
ms.prod: visio
api_name:
- Visio.InvisibleApp.ShowStatusBar
ms.assetid: e58097a2-a5e9-00ca-c3b3-74a3d7717907
ms.date: 06/08/2017
---


# InvisibleApp.ShowStatusBar Property (Visio)

Determines whether the status bar is shown. Read/write. 


## Syntax

 _expression_ . **ShowStatusBar**

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

Integer


## Remarks

The  **ShowStatusBar** property persists each time you run the application. The **ShowToolbar** property is valid for a Microsoft Visio instance only.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to hide and show the status bar.


```vb
 
Public Sub ShowStatusBar_Example() 
 
 'Switch the status bar on or off. 
 Application.ShowStatusBar = Not Application.ShowStatusBar 
 
End Sub
```


