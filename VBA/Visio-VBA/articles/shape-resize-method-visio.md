---
title: Shape.Resize Method (Visio)
keywords: vis_sdr.chm11262205
f1_keywords:
- vis_sdr.chm11262205
ms.prod: visio
api_name:
- Visio.Shape.Resize
ms.assetid: ce8e9253-e1bb-e542-30eb-f9ac2e4305da
ms.date: 06/08/2017
---


# Shape.Resize Method (Visio)

Resizes the shape by moving shape handles as specified.


## Syntax

 _expression_ . **Resize**( **_Direction_** , **_Distance_** , **_UnitCode_** )

 _expression_ A variable that represents a **[Shape](shape-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Direction_|Required| **[VisResizeDirection](visresizedirection-enumeration-visio.md)**|The direction that corresponds to the shape handle. See Remarks for possible values.|
| _Distance_|Required| **Double**|The distance to move the shape edge or corner; positive values move outward and negative values move inward.|
| _UnitCode_|Required| **[VisUnitCodes](visunitcodes-enumeration-visio.md)**|The units of measure to use for the resize distance.|

### Return Value

 **Nothing**


## Remarks

 _Direction_ must be one of the following **VisResizeDirection** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visResizeDirE**|0|Right, middle shape handle.|
| **visResizeDirNE**|1|Right, top shape handle.|
| **visResizeDirN**|2|Center, top shape handle.|
| **visResizeDirNW**|3|Left, top shape handle.|
| **visResizeDirW**|4|Left, middle shape handle.|
| **visResizeDirSW**|5|Left, bottom shape handle.|
| **visResizeDirS**|6|Center, bottom shape handle.|
| **visResizeDirSE**|7|Right, bottom shape handle.|
When you resize a shape in a diagonal direction (that is, NE, NW, SE, or SW), the specified distance is applied along both the horizontal and vertical axes (as opposed to along the compass direction.)


