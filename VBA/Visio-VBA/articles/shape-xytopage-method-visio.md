---
title: Shape.XYToPage Method (Visio)
keywords: vis_sdr.chm11216650
f1_keywords:
- vis_sdr.chm11216650
ms.prod: visio
api_name:
- Visio.Shape.XYToPage
ms.assetid: 4a230d63-57a8-3b69-6425-2dca6a2014eb
ms.date: 06/08/2017
---


# Shape.XYToPage Method (Visio)

Transforms a point expressed in the local coordinate system of a  **Shape** object to an equivalent point expressed in the local coordinate system of its **Page** or **Master** object.


## Syntax

 _expression_ . **XYToPage**( **_x_** , **_y_** , **_xprime_** , **_yprime_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _x_|Required| **Double**| _x_-coordinate in coordinate system of  _object._|
| _y_|Required| **Double**| _y_-coordinate in coordinate system of  _object._|
| _xprime_|Required| **Double**| _x_-coordinate corresponding to  _x_ in the **Page** or **Master** object's coordinate system.|
| _yprime_|Required| **Double**| _y_-coordinate corresponding to  _y_ in the **Page** or **Master** object's coordinate system.|

### Return Value

Nothing


## Remarks

The points  _x_,  _y_,  _xprime_, and  _yprime_ are all treated as internal drawing units.

An exception is raised if object is not a  **Shape** object of a **Page** or **Master** object.


