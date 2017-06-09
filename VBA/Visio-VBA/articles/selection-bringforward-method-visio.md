---
title: Selection.BringForward Method (Visio)
keywords: vis_sdr.chm11116095
f1_keywords:
- vis_sdr.chm11116095
ms.prod: visio
api_name:
- Visio.Selection.BringForward
ms.assetid: d12a81a5-6faa-6828-bdf0-279c27c89571
ms.date: 06/08/2017
---


# Selection.BringForward Method (Visio)

Brings the shape or selected shapes forward one position in the z-order.


## Syntax

 _expression_ . **BringForward**

 _expression_ A variable that represents a **Selection** object.


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


