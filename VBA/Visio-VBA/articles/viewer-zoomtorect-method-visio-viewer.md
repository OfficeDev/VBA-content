---
title: Viewer.ZoomToRect Method (Visio Viewer)
ms.prod: visio
api_name:
- Visio.ZoomToRect
ms.assetid: 80d4da31-55b9-abc8-9727-6ebd8ebe0ddb
ms.date: 06/08/2017
---


# Viewer.ZoomToRect Method (Visio Viewer)

Zooms to display a rectanglular section, specified by the parameters, of the drawing that is open in Microsoft Visio Viewer.


## Syntax

 _expression_. **ZoomToRect**( **_Left_**,  **_Top_**,  **_Right_**,  **_Bottom_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Left|Required| **Long**|The x-coordinate of the top left corner of the rectangle to zoom to.|
|Top|Required| **Long**|The y-coordinate of the top left corner of the rectangle to zoom to.|
|Right|Required| **Long**|The x-coordinate of the bottom right corner of the rectangle to zoom to.|
|Bottom|Required| **Long**|The y-coordinate of the bottom right corner of the rectangle to zoom to.|

### Return Value

Nothing


## Remarks

The coordinate system for the  **ZoomToRect** method has its origin at the top left corner of the Visio Viewer window. Positive directions are to the right (x) and down (y). The units of measurement are pixels.

The  **ZoomToRect** method, in effect, crops a rectangular section of the drawing, specified by the parameters, and then displays that section in the entire Visio Viewer window. The parameters are a set of two pairs of coordinates, the first pair specifying the upper left corner of the section, and the second pair the lower right corner.


## Example

The following code zooms to display a rectangular section of the drawing that is open in Visio Viewer. The upper left corner of the displayed section corresponds to the upper left corner of the Visio Viewer window as originally displayed. The lower right corner corresponds to a point 300 pixels down and 300 pixels to the right of the upper left corner in that original display.


```
vsoViewer.ZoomToRect 0, 0, 300, 300
```


