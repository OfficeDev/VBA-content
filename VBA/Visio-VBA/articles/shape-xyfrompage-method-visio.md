---
title: Shape.XYFromPage Method (Visio)
keywords: vis_sdr.chm11216645
f1_keywords:
- vis_sdr.chm11216645
ms.prod: visio
api_name:
- Visio.Shape.XYFromPage
ms.assetid: 85b04e0b-04e1-a5b5-f6ff-393c57751946
ms.date: 06/08/2017
---


# Shape.XYFromPage Method (Visio)

Transforms a point expressed in the local coordinate system of its  **Page** or **Master** object to an equivalent point expressed in the local coordinate system of the **Shape** object.


## Syntax

 _expression_ . **XYFromPage**( **_x_** , **_y_** , **_xprime_** , **_yprime_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _x_|Required| **Double**| _x_-coordinate corresponding to  _x_ in the **Page** or **Master** object's coordinate system.|
| _y_|Required| **Double**| _y_-coordinate corresponding to  _y_ in the **Page** or **Master** object's coordinate system.|
| _xprime_|Required| **Double**| _x_-coordinate in coordinate system of  _object._|
| _yprime_|Required| **Double**| _y_-coordinate in coordinate system of  _object._|

### Return Value

Nothing


## Remarks

The points  _x_,  _y_,  _xprime_, and  _yprime_ are all treated as internal drawing units.

An exception is raised if object is not a  **Shape** object of a **Page** or **Master** object.


