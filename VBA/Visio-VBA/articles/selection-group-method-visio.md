---
title: Selection.Group Method (Visio)
keywords: vis_sdr.chm11116345
f1_keywords:
- vis_sdr.chm11116345
ms.prod: visio
api_name:
- Visio.Selection.Group
ms.assetid: 79afc3c4-7350-2196-7a07-3b7c5629568a
ms.date: 06/08/2017
---


# Selection.Group Method (Visio)

Groups the objects that are selected in a selection, or it converts a shape into a group.


## Syntax

 _expression_ . **Group**

 _expression_ A variable that represents a **Selection** object.


### Return Value

Shape


## Example

The following example shows how to group  **Shape** objects.


```vb
 
Public Sub Group_Example() 
 
 Dim vsoShape1 As Visio.Shape 
 Dim vsoShape2 As Visio.Shape 
 Dim vsoGroupShape As Visio.Shape 
 Dim vsoSelection As Visio.Selection 
 
 'Draw two rectangles. 
 Set vsoShape1 = ActivePage.DrawRectangle(1, 2, 2, 1) 
 Set vsoShape2 = ActivePage.DrawRectangle(1, 4, 2, 3) 
 
 'Deselect all shapes, and then select the two rectangles. 
 Set vsoSelection = ActiveWindow.Selection 
 vsoSelection.Select vsoShape1, visDeselectAll + visSelect 
 vsoSelection.Select vsoShape2, visSelect 
 
 'Group the rectangles into a group shape. 
 Set vsoGroupShape = vsoSelection.Group 
 
End Sub
```


