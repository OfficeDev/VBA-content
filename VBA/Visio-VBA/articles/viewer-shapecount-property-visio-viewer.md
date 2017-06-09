---
title: Viewer.ShapeCount Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.ShapeCount
ms.assetid: b1a8a4a8-5140-4586-fc4d-be64b47d0158
ms.date: 06/08/2017
---


# Viewer.ShapeCount Property (Visio Viewer)

Gets the count of shapes in the drawing that is open in Microsoft Visio Viewer. Read-only.


## Syntax

 _expression_. **ShapeCount**

 _expression_An expression that returns a  **Viewer** object.


### Return Value

 **Long**


## Remarks

If there are no shapes in the drawing, the ShapeCount property returns 0.


## Example

The following code gets the count of shapes in the drawing that is open in Visio Viewer. Subshapes and group shapes are both included in the count.


```vb
Debug.Print vsoViewer.ShapeCount
```


