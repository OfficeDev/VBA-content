---
title: Selection.SendBackward Method (Visio)
keywords: vis_sdr.chm11116540
f1_keywords:
- vis_sdr.chm11116540
ms.prod: visio
api_name:
- Visio.Selection.SendBackward
ms.assetid: 645a5686-6421-f8dd-425f-3cb5b0b7de85
ms.date: 06/08/2017
---


# Selection.SendBackward Method (Visio)

Moves a shape or selected shapes back one position in the z-order.


## Syntax

 _expression_ . **SendBackward**

 _expression_ A variable that represents a **Selection** object.


### Return Value

Nothing


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to move a shape back one position in the z-order on a page.


```vb
 
Public Sub SendBackward_Example() 
 
 Dim vsoShape1 As Visio.Shape 
 Dim vsoShape2 As Visio.Shape 
 Dim vsoShape3 As Visio.Shape 
 
 'Draw three rectangles. 
 Set vsoShape1 = ActivePage.DrawRectangle(1, 1, 5, 5) 
 vsoShape1.Text = "1" 
 Set vsoShape2 = ActivePage.DrawRectangle(2, 2, 6, 6) 
 vsoShape2.Text = "2" 
 Set vsoShape3 = ActivePage.DrawRectangle(3, 3, 7, 7) 
 vsoShape3.Text = "3" 
 
 'Move vsoShape2 back one position in the z-order. 
 vsoShape2.SendBackward 
 
End Sub
```


