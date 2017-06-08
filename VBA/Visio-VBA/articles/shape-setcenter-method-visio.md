---
title: Shape.SetCenter Method (Visio)
keywords: vis_sdr.chm11216555
f1_keywords:
- vis_sdr.chm11216555
ms.prod: visio
api_name:
- Visio.Shape.SetCenter
ms.assetid: 9a3c0597-c255-44ab-9268-938acd3c5a69
ms.date: 06/08/2017
---


# Shape.SetCenter Method (Visio)

Moves a shape so that its pin is positioned at the coordinates represented by  _xPos_ and _yPos_. .


## Syntax

 _expression_ . **SetCenter**( **_xPos_** , **_yPos_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _xPos_|Required| **Double**|The new x-coordinate of the center of rotation (PinX).|
| _yPos_|Required| **Double**|The new y-coordinate of the center of rotation (PinY).|

### Return Value

Nothing


## Remarks

The coordinates represented by the  _xPos_ and _yPos_ arguments are parent coordinates, measured from the origin of the shape's parent (the page or group that contains the shape).

The  **SetCenter** method only moves the point, in parent coordinates, about which the shape rotates. It does not change the point, in local coordinates, about which the shape rotates. The overall effect is to move the shape with respect to its parent shape (or the page).


