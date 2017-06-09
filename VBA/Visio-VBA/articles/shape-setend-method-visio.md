---
title: Shape.SetEnd Method (Visio)
keywords: vis_sdr.chm11216570
f1_keywords:
- vis_sdr.chm11216570
ms.prod: visio
api_name:
- Visio.Shape.SetEnd
ms.assetid: 5f2c7b85-52b3-9147-a989-b2dce61c3493
ms.date: 06/08/2017
---


# Shape.SetEnd Method (Visio)

Moves the endpoint of a one-dimensional (1-D) shape to the coordinates represented by  _xPos_ and _yPos_.


## Syntax

 _expression_ . **SetEnd**( **_xPos_** , **_yPos_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _xPos_|Required| **Double**|The new x-coordinate of the endpoint.|
| _yPos_|Required| **Double**|The new y-coordinate of the endpoint.|

### Return Value

Nothing


## Remarks

The  **SetEnd** method applies only to 1-D shapes. If the indicated shape is a 2-D shape, an error is returned.

The coordinates represented by the  _xPos_ and _yPos_ arguments are parent coordinates, measured from the origin of the shape's parent (the page or group that contains the shape).


