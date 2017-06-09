---
title: Shape.BringForward Method (Visio)
keywords: vis_sdr.chm11216095
f1_keywords:
- vis_sdr.chm11216095
ms.prod: visio
api_name:
- Visio.Shape.BringForward
ms.assetid: 88e5c746-e7f2-eddd-35c9-2d9c09c1a602
ms.date: 06/08/2017
---


# Shape.BringForward Method (Visio)

Brings the shape or selected shapes forward one position in the z-order.


## Syntax

 _expression_ . **BringForward**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Nothing


## Example

The following macro shows how to bring a shape forward in the z-order on a page.


```vb
 
Public Sub BringForward_Example() 
 
 Dim vsoShape1 As Visio.Shape 
 Dim vsoShape2 As Visio.Shape 
 Dim vsoShape3 As Visio.Shape 
 
 'Draw three rectangles. 
 Set vsoShape1 = ActivePage.DrawRectangle(1, 1, 4, 4) 
 vsoShape1.Text = "1" 
 Set vsoShape2 = ActivePage.DrawRectangle(2, 2, 5, 5) 
 vsoShape2.Text = "2" 
 Set vsoShape3 = ActivePage.DrawRectangle(3, 3, 6, 6) 
 vsoShape3.Text = "3" 
 
 'Bring vsoShape1 forward one position in the z-order. 
 vsoShape1.BringForward 
 
End Sub
```


