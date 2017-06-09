---
title: Shape.TransformXYTo Method (Visio)
keywords: vis_sdr.chm11216605
f1_keywords:
- vis_sdr.chm11216605
ms.prod: visio
api_name:
- Visio.Shape.TransformXYTo
ms.assetid: dc85cf08-0d83-34ff-8389-94a0f5f05c5e
ms.date: 06/08/2017
---


# Shape.TransformXYTo Method (Visio)

Transforms a point expressed in the local coordinate system of one  **Shape** object to an equivalent point expressed in the local coordinate system of another **Shape** object.


## Syntax

 _expression_ . **TransformXYTo**( **_OtherShape_** , **_x_** , **_y_** , **_xprime_** , **_yprime_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _OtherShape_|Required| **[IVSHAPE]**|An expression that returns a  **Shape** object whose local coordinate system you are transforming the point to.|
| _x_|Required| **Double**| _x_-coordinate in coordinate system of  _object._|
| _y_|Required| **Double**| _y_-coordinate in coordinate system of  _object._|
| _xprime_|Required| **Double**| _x_-coordinate corresponding to  _x_in the  _OtherShape_coordinate system.|
| _yprime_|Required| **Double**| _y_-coordinate corresponding to  _y_in the  _OtherShape_coordinate system.|

### Return Value

Nothing


## Remarks

The points  _x_,  _y_,  _xprime_ and _yprime_ are all treated as internal drawing units.

An exception is raised if object is not a  **Shape** object of a **Page** or **Master** object, or if _OtherShape_ is not in the same **Page** or **Master** object as _object_.


