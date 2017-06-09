---
title: Shape.SendToBack Method (Visio)
keywords: vis_sdr.chm11216545
f1_keywords:
- vis_sdr.chm11216545
ms.prod: visio
api_name:
- Visio.Shape.SendToBack
ms.assetid: faa9cab5-0b2f-8331-e0df-8b4f4be1e69f
ms.date: 06/08/2017
---


# Shape.SendToBack Method (Visio)

Moves the shape or selected shapes to the back of the z-order.


## Syntax

 _expression_ . **SendToBack**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Nothing


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to move a shape to the back of the z-order on a page.


```vb
 
Public Sub SendToBack_Example() 
 
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
 
 'Move vsoShape3 to the back of the z-order. 
 vsoShape3.SendToBack 
 
End Sub
```


