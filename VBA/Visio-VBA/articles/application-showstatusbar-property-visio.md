---
title: Application.ShowStatusBar Property (Visio)
keywords: vis_sdr.chm10014380
f1_keywords:
- vis_sdr.chm10014380
ms.prod: visio
api_name:
- Visio.Application.ShowStatusBar
ms.assetid: a6eade7f-b056-92ef-0a57-acd466f6a99a
ms.date: 06/08/2017
---


# Application.ShowStatusBar Property (Visio)

Determines whether the status bar is shown. Read/write.


## Syntax

 _expression_ . **ShowStatusBar**

 _expression_ A variable that represents an **Application** object.


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


