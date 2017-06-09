---
title: Selection.Trim Method (Visio)
keywords: vis_sdr.chm11116615
f1_keywords:
- vis_sdr.chm11116615
ms.prod: visio
api_name:
- Visio.Selection.Trim
ms.assetid: 0063d29a-3e47-bb2b-71fd-328c19a0a65b
ms.date: 06/08/2017
---


# Selection.Trim Method (Visio)

Trims selected shapes into smaller shapes.


## Syntax

 _expression_ . **Trim**

 _expression_ A variable that represents a **Selection** object.


### Return Value

Nothing


## Remarks

Calling the  **Trim** method is equivalent to clicking **Trim** in the Microsoft Visio user interface (click **Operations** in the **Shape Design** group on the[Developer](http://msdn.microsoft.com/library/1bdc55f5-8fc7-7257-03d5-c049eceb29ff%28Office.15%29.aspx) tab).

The new shapes inherit the formatting of the first selected shape, have no text, and are the topmost shapes in their container—the  _n_th shape,  _n_th - 1 shape,  _n_th - 2 shape, and so forth in the  **Shapes** collection of their containing shape, where _n_ = count. The original shapes are deleted and no shapes are selected when the operation is complete.

The  **Trim** method is similar to the **Fragment** method but differs in the following ways:




- Shapes produced by the  **Trim** method coincide with the distinct paths of the selected shapes, taking overlap into account.
    
- Shapes produced by the  **Fragment** method coincide with the distinct regions of the selected shapes, also taking overlap into account.
    



## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Trim** method to trim selected shapes into smaller shapes along their intersections.


```vb
Public Sub Trim_Example() 
 
 Dim vsoShape1 As Visio.Shape 
 Dim vsoShape2 As Visio.Shape 
 Dim shapeCount As Integer 
 
 'Draw two shapes that intersect 
 Set vsoShape1 = ActivePage.DrawRectangle(1, 4, 4, 1) 
 Set vsoShape2 = ActivePage.DrawOval(2, 6, 3, 2) 
 
 'Deselect the oval and then select both of the new shapes on the page 
 ActiveWindow.DeselectAll 
 ActiveWindow.SelectAll 
 
 'Create a selection object and assign the selected shapes to it 
 Dim vsoSelection As Visio.Selection 
 Set vsoSelection = ActiveWindow.Selection 
 
 'Trim the selected shapes 
 vsoSelection.Trim 
 
 'Move one of the newly created shapes 
 ActiveWindow.DeselectAll 
 shapeCount = ActivePage.Shapes.Count 
 
 Set vsoShape1 = ActivePage.Shapes(shapeCount - 2) 
 ActiveWindow.Select vsoShape1, visSelect 
 ActiveWindow.Selection.Move 2, 2 
 
End Sub
```


