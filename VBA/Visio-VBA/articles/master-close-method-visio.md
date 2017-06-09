---
title: Master.Close Method (Visio)
keywords: vis_sdr.chm10716125
f1_keywords:
- vis_sdr.chm10716125
ms.prod: visio
api_name:
- Visio.Master.Close
ms.assetid: 69607a2c-dc59-d170-733a-3557a996a67e
ms.date: 06/08/2017
---


# Master.Close Method (Visio)

Closes a master.


## Syntax

 _expression_ . **Close**

 _expression_ A variable that represents a **Master** object.


### Return Value

Nothing


## Remarks

Use the  **Close** method for a **Master** object after opening a master for editing using the **Open** method. The **Close** method pushes any changes made to the master while it was open to instances of the master.


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


