---
title: Selection.ConnectShapes Method (Visio)
keywords: vis_sdr.chm11152010
f1_keywords:
- vis_sdr.chm11152010
ms.prod: visio
api_name:
- Visio.Selection.ConnectShapes
ms.assetid: 40e9c839-69f0-2142-6b9c-249212e373a4
ms.date: 06/08/2017
---


# Selection.ConnectShapes Method (Visio)

Connects two or more selected shapes with a dynamic connector. Returns **Nothing** .


## Syntax

 _expression_ . **ConnectShapes**

 _expression_ A variable that represents a **Selection** object.


### Return Value

Nothing


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **ConnectShapes** method to connect two shapes.


```vb
Public Sub ConnectShapes_Example() 
 
 Dim vsoShape1 As Visio.Shape 
 Dim vsoShape2 As Visio.Shape 
 
 Set vsoShape1 = Application.ActiveWindow.Page.DrawRectangle(2, 9, 4, 7) 
 Set vsoShape2 = Application.ActiveWindow.Page.DrawRectangle(5, 6, 7, 3) 
 
 ActiveWindow.DeselectAll 
 ActiveWindow.Select vsoShape1, visSelect 
 ActiveWindow.Select vsoShape2, visSelect 
 Application.ActiveWindow.Selection.ConnectShapes 
 
End Sub
```


