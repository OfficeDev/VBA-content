---
title: Window.Selection Property (Visio)
keywords: vis_sdr.chm11614310
f1_keywords:
- vis_sdr.chm11614310
ms.prod: visio
api_name:
- Visio.Window.Selection
ms.assetid: 67c3b3d3-9fe4-ff0c-db94-4a2109f29736
ms.date: 06/08/2017
---


# Window.Selection Property (Visio)

Returns a  **Selection** object that represents what is presently selected in the window, or assigns a selection created by the **CreateSelection** method to a **Selection** object. Read/write.


## Syntax

 _expression_ . **Selection**

 _expression_ A variable that represents a **Window** object.


### Return Value

Selection


## Remarks

The  **Selection** object is independent of the selection in the window, which can subsequently change as a result of user actions.

A  **Selection** object is a set of shapes in a common context on which you can perform actions. A **Selection** object is analogous to more than selected shapes in a drawing window. Once you set or retrieve a **Selection** object, you can change the set of shapes the object represents by using the **Select** method.

After you use the  **CreateSelection** method to create a selection, you can then use the **Selection** property to actually display the newly created selection in the Microsoft Visio drawing window. See the second example that follows.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Selection** property to get all the selected shapes in the window.


```vb
Public Sub Selection_Example() 
 
 Const MAX_SHAPES = 6 
 Dim vsoShapes(1 To MAX_SHAPES) As Visio.Shape 
 Dim vsoSelection As Visio.Selection 
 Dim intCounter As Integer 
 
 'Draw six rectangles. 
 For intCounter = 1 To MAX_SHAPES 
 Set vsoShapes(intCounter) = ActivePage.DrawRectangle(intCounter, intCounter + 1, intCounter + 1, intCounter) 
 Next intCounter 
 
 'Deselect all the shapes in the active window. 
 ActiveWindow.DeselectAll 
 
 'Select all the shapes in the active window. 
 ActiveWindow.SelectAll 
 
 'Get the selected shapes and assign them to a Selection object. 
 Set vsoSelection = ActiveWindow.Selection 
 
End Sub
```

This VBA macro shows how to use the  **CreateSelection** method to select all shapes on a particular layer. Then it uses the **Selection** property to display the selection in the Visio drawing window.

Before running this macro, create two layers in your drawing, one named "a" and one named "b", and then add shapes to both layers.




```vb
Public Sub Selection_Example_2() 
 
 Dim vsoLayer As Layer 
 Dim vsoSelection As Visio.Selection 
 
 Set vsoLayer = ActivePage.Layers.ItemU("a") 
 Set vsoSelection = ActivePage.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, VsoLayer) 
 
 Application.ActiveWindow.Selection = vsoSelection 
 
End Sub
```


