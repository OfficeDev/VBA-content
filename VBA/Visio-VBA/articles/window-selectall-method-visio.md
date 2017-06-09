---
title: Window.SelectAll Method (Visio)
keywords: vis_sdr.chm11616535
f1_keywords:
- vis_sdr.chm11616535
ms.prod: visio
api_name:
- Visio.Window.SelectAll
ms.assetid: 81c32217-3336-3017-ecdc-cfa0f6048fc2
ms.date: 06/08/2017
---


# Window.SelectAll Method (Visio)

Selects all possible shapes in a window or selection.


## Syntax

 _expression_ . **SelectAll**

 _expression_ A variable that represents a **Window** object.


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


