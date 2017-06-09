---
title: Shape.Group Method (Visio)
keywords: vis_sdr.chm11216345
f1_keywords:
- vis_sdr.chm11216345
ms.prod: visio
api_name:
- Visio.Shape.Group
ms.assetid: fe19f27f-47ad-93ef-1d82-4010d8cb6e47
ms.date: 06/08/2017
---


# Shape.Group Method (Visio)

Groups the objects that are selected in a selection, or it converts a shape into a group.


## Syntax

 _expression_ . **Group**

 _expression_ A variable that represents a **Shape** object.


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


