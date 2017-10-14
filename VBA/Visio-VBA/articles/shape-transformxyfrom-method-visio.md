---
title: Shape.TransformXYFrom Method (Visio)
keywords: vis_sdr.chm11216600
f1_keywords:
- vis_sdr.chm11216600
ms.prod: visio
api_name:
- Visio.Shape.TransformXYFrom
ms.assetid: 4676e464-83c7-7ff6-e742-becc41436259
ms.date: 06/08/2017
---


# Shape.TransformXYFrom Method (Visio)

Transforms a point expressed in the local coordinate system of one  **Shape** object from an equivalent point expressed in the local coordinate system of another **Shape** object.


## Syntax

 _expression_ . **TransformXYFrom**( **_OtherShape_** , **_x_** , **_y_** , **_xprime_** , **_yprime_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _OtherShape_|Required| **[IVSHAPE]**|An expression that returns a  **Shape** object whose local coordinate system you are transforming the point from.|
| _x_|Required| **Double**| _x_-coordinate corresponding to  _x_ in the _OtherShape_ coordinate system.|
| _y_|Required| **Double**| _y_-coordinate corresponding to  _y_ in the _OtherShape_ coordinate system.|
| _xprime_|Required| **Double**| _x_-coordinate in coordinate system of  _object._|
| _yprime_|Required| **Double**| _y_-coordinate in coordinate system of  _object._|

### Return Value

Nothing


## Remarks

The points  _x_,  _y_,  _xprime_, and  _yprime_ are all treated as internal drawing units.

An exception is raised if object is not a  **Shape** object of a **Page** or **Master** object, or if _OtherShape_ is not in the same **Page** or **Master** object as _object_.


