---
title: COMAddIn.Connect Property (Office)
keywords: vbaof11.chm219005
f1_keywords:
- vbaof11.chm219005
ms.prod: office
api_name:
- Office.COMAddIn.Connect
ms.assetid: b1392380-c19f-ab3e-c9dc-c62438b16500
ms.date: 06/08/2017
---


# COMAddIn.Connect Property (Office)

Gets or sets the state of the connection for the specified  **COMAddIn** object. Read/write.


## Syntax

 _expression_. **Connect**

 _expression_ A variable that represents a **COMAddIn** object.


## Remarks

The  **Connect** property returns **True** if the add-in is active; it returns **False** if the add-in is inactive. An active add-in is registered and connected; an inactive add-in is registered but not currently connected.


## Example

The following example displays a message box that indicates whether COM add-in one is registered and currently connected.


```
If Application.COMAddIns(1).Connect Then 
 MsgBox "The add-in is connected." 
Else 
 MsgBox "The add-in is not connected." 
End If
```


## See also


#### Concepts


[COMAddIn Object](comaddin-object-office.md)
#### Other resources


[COMAddIn Object Members](comaddin-members-office.md)

