---
title: Shape.HitTest Method (Visio)
keywords: vis_sdr.chm11213645
f1_keywords:
- vis_sdr.chm11213645
ms.prod: visio
api_name:
- Visio.Shape.HitTest
ms.assetid: 1250ac1d-32f8-d078-3a01-6e2ce045d254
ms.date: 06/08/2017
---


# Shape.HitTest Method (Visio)

Determines if a given  _x,y_ position hits outside, inside, or on the boundary of a shape.


## Syntax

 _expression_ . **HitTest**( **_xPos_** , **_yPos_** , **_Tolerance_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _xPos_|Required| **Double**|The x-coordinate to be tested for a hit.|
| _yPos_|Required| **Double**|The y-coordinate to be tested for a hit.|
| _Tolerance_|Required| **Double**|How close  _xPos,yPos_ must be to a shape for a hit to occur.|

### Return Value

Integer


## Remarks

The  **HitTest** method considers only visible geometry and ignores hidden geometry.

Use internal drawing units (inches in the drawing) for the  _xPos_,  _yPos_, and  _Tolerance_ values. These values should also be in, and with respect to, the coordinate space of the page, master, or group shape that contains the shape being hit tested.

The following are possible values returned by the  **HitTest** method, and are declared by the Visio type library in **VisHitTestResults** .



|**Constant**|**Value**|
|:-----|:-----|
| **visHitOutside**|0|
| **visHitOnBoundary**|1|
| **visHitInside**|2|
Data graphic callout shapes (and their sub-shapes) that are applied to the parent shape are excluded from hit-test calculations. If the parent shape is itself a data graphic callout shape, its geometry (and that of its sub-shapes) is  _not_ excluded from hit-test calculations.


