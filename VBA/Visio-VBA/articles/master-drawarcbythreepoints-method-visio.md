---
title: Master.DrawArcByThreePoints Method (Visio)
keywords: vis_sdr.chm10751415
f1_keywords:
- vis_sdr.chm10751415
ms.prod: visio
api_name:
- Visio.Master.DrawArcByThreePoints
ms.assetid: d2df1c41-8164-d941-21a8-2e1b00de6199
ms.date: 06/08/2017
---


# Master.DrawArcByThreePoints Method (Visio)

Creates a shape whose path consists of an arc defined by the three points passed as parameters.


## Syntax

 _expression_ . **DrawArcByThreePoints**( **_xBegin_** , **_yBegin_** , **_xEnd_** , **_yEnd_** , **_xControl_** , **_yControl_** )

 _expression_ A variable that represents a **Master** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _xBegin_|Required| **Double**|The x-coordinate of the begin point of the arc.|
| _yBegin_|Required| **Double**|The y-coordinate of the begin point of the arc.|
| _xEnd_|Required| **Double**|The x-coordinate of the endpoint of the arc.|
| _yEnd_|Required| **Double**|The y-coordinate of the endpoint of the arc.|
| _xControl_|Required| **Double**|The x-coordinate of the control point of the arc.|
| _yControl_|Required| **Double**|The y-coordinate of the control point of the arc.|

### Return Value

Shape


## Remarks

All points should be in internal drawing units with respect to the coordinate space of the master, page, or group where the shape is being drawn.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **DrawArcByThreePoints** method to draw an arc on the drawing page.


```vb
Public Sub DrawArcByThreePoints_Example 
 
 Dim vsoShape As Visio.Shape 
 Set vsoShape = ActivePage.DrawArcByThreePoints(3, 3, 6, 8, 5, 5) 
 
End Sub
```


