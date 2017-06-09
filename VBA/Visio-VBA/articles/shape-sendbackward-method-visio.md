---
title: Shape.SendBackward Method (Visio)
keywords: vis_sdr.chm11216540
f1_keywords:
- vis_sdr.chm11216540
ms.prod: visio
api_name:
- Visio.Shape.SendBackward
ms.assetid: 9e43cfd9-c2d3-9042-46e3-39e209ae54aa
ms.date: 06/08/2017
---


# Shape.SendBackward Method (Visio)

Moves a shape or selected shapes back one position in the z-order.


## Syntax

 _expression_ . **SendBackward**

 _expression_ A variable that represents a **Shape** object.


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


