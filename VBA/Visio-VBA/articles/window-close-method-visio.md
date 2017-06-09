---
title: Window.Close Method (Visio)
keywords: vis_sdr.chm11616125
f1_keywords:
- vis_sdr.chm11616125
ms.prod: visio
api_name:
- Visio.Window.Close
ms.assetid: 43cb221f-ea65-c12a-e664-0f0fb35685e0
ms.date: 06/08/2017
---


# Window.Close Method (Visio)

Closes a window.


## Syntax

 _expression_ . **Close**

 _expression_ A variable that represents a **Window** object.


### Return Value

Nothing


## Remarks

If the indicated window is the only window open for a document and the document contains unsaved changes, an alert appears asking if you want to save the document. You can use the  **AlertResponse** property to prevent the alert from appearing.

If you close a docked stencil window, only that window is closed. However, if you close a drawing window that contains docked stencils, the docked stencil window is also closed.


## Example

This example shows how to close all open ShapeSheet windows. It assumes at least one ShapeSheet window is open in Microsoft Visio.


```vb
 
Public Sub Close_Example() 
 Dim intCounter As Integer 
 intCounter = Windows.Count 
 
 'Close all ShapeSheet windows that are open. 
 While intCounter <> 0 
 If Windows(intCounter).Type = visSheet Then 
 Windows(intCounter).Close 
 intCounter = Windows.Count 
 Else 
 intCounter = intCounter - 1 
 End If 
 Wend 
End Sub
```


