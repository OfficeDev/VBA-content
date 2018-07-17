---
title: Selection.SelectAll Method (Visio)
keywords: vis_sdr.chm11116535
f1_keywords:
- vis_sdr.chm11116535
ms.prod: visio
api_name:
- Visio.Selection.SelectAll
ms.assetid: e2280c51-84e8-4403-1c9e-f3bc504aff2f
ms.date: 06/08/2017
---


# Selection.SelectAll Method (Visio)

Selects all possible shapes in a window or selection.


## Syntax

 _expression_ . **SelectAll**

 _expression_ A variable that represents a **Selection** object.


### Return Value

Nothing


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to select all the shapes on the page.


```vb
 
Public Sub SelectAll_Example() 
 
 Const MAX_SHAPES = 6 
 Dim vsoShapes(1 To MAX_SHAPES) As Visio.Shape 
 Dim intCounter As Integer 
 
 'Draw six rectangles. 
 For intCounter = 1 To MAX_SHAPES 
 Set vsoShapes(intCounter) = ActivePage.DrawRectangle(intCounter, intCounter + 1, intCounter + 1, intCounter) 
 Next intCounter 
 
 'Deselect all the shapes on the page. 
 ActiveWindow.DeselectAll 
 
 'Select all the shapes on the page. 
 ActiveWindow.SelectAll 
 
End Sub
```


