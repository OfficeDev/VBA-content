---
title: Viewer.SubShapeAtPoint Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.SubShapeAtPoint
ms.assetid: ffb35fad-33ee-30d0-680f-008418b58864
ms.date: 06/08/2017
---


# Viewer.SubShapeAtPoint Property (Visio Viewer)

Gets the ID of the subshape in the drawing that is open in Microsoft Visio Viewer, at the specified point in the Visio Viewer window, in the coordinate system of the window, measured in pixels. Read-only.


## Syntax

 _expression_. **SubShapeAtPoint**( **_X_**,  **_Y_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|X|Required| **Long**|The x-coordinate, in pixels, of the point.|
|Y|Required| **Long**|The y-coordinate, in pixels, of the point.|

### Return Value

 **Long**


## Remarks

A  _subshape_ is a shape that is a member of a group shape.

The origin of the coordinate system of the Visio Viewer window is the upper left corner. The positive x-direction is to the right, and the positive y-direction is down.

If there is no subshape at the specified point, the  **SubShapeAtPoint** property returns 0.


## Example

The following code gets the ID of the subshape at point (200, 200) in the drawing that is open in Visio Viewer.


```vb
Debug.Print vsoViewer.SubShapeAtPoint (200, 200)
```


