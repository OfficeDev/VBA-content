---
title: Global.ActivePage Property (Visio)
keywords: vis_sdr.chm12813030
f1_keywords:
- vis_sdr.chm12813030
ms.prod: visio
api_name:
- Visio.Global.ActivePage
ms.assetid: ac981df6-5d43-f93d-f10c-41bd8e0ad2de
ms.date: 06/08/2017
---


# Global.ActivePage Property (Visio)

Returns the active  **Page** object. Read-only.


## Syntax

 _expression_ . **ActivePage**

 _expression_ A variable that represents a **Global** object.


### Return Value

Page


## Remarks

The  **ActivePage** property returns a **Page** object only when the active window displays a drawing page; otherwise, it returns **Nothing** . To verify that a page is active, use the **Is** operator to compare the **ActivePage** property with **Nothing** .

It is possible to get the active window without qualification from the Microsoft Visio global object, which is automatically available to VBA code that is part of the VBA project of a Visio document. For example, you can use this code: 




```vb
Set vsoPage = ActivePage
```


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to get the active page without qualification from the Visio global object, which is automatically available to VBA code that is part of the VBA project of a Visio document.


```vb
 
Public Sub ActivePage_Example() 
 
 Dim vsoPage As Page 
 
 'Find out if a page exists, and if it does, get the page. 
 If Not(ActivePage Is Nothing) Then 
 Set vsoPage = ActivePage 
 Debug.Print vsoPage.Name 
 Else 
 Debug.Print "No active page." 
 End If 
 
End Sub
```


