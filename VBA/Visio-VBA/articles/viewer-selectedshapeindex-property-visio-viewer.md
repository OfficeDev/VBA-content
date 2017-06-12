---
title: Viewer.SelectedShapeIndex Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.SelectedShapeIndex
ms.assetid: dbf6c737-e8b5-8768-533f-2625d99a1717
ms.date: 06/08/2017
---


# Viewer.SelectedShapeIndex Property (Visio Viewer)

Gets the index in the collection of shapes of the selected shape in the drawing that is open in Microsoft Visio Viewer. Read-only.


## Syntax

 _expression_. **SelectedShapeIndex**

 _expression_An expression that returns a  **Viewer** object.


### Return Value

 **Long**


## Remarks

The collection of shapes is one-based, so the index of the first shape in the collection is 1.

If no shapes are selected in the drawing, the  **SelectedShapeIndex** property returns 0.


## Example

The following code iterates through the collection of shapes in the document that is open in Visio Viewer, selects each shape in turn, and then gets the value of the first shape data item (custom property) assigned to each shape. If it discovers a custom property value of "Computer 100", it displays a message box to that effect.


```vb
Dim intSelectedShapeIndex As Integer

Dim intShapeCounter As Integer

For intShapeCounter = 1 To vsoViewer.ShapeCount

    vsoViewer.SelectShape (intShapeCounter)

    intSelectedShapeIndex = vsoViewer.SelectedShapeIndex

    If vsoViewer.CustomPropertyValue(intSelectedShapeIndex, 1) = "Computer 100" Then

        Interaction.MsgBox ("Selected shape name is " &; vsoViewer.CustomPropertyValue(intSelectedShapeIndex, 1))

    End If

Next
```


