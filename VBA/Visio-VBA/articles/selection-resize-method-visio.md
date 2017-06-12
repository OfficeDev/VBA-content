---
title: Selection.Resize Method (Visio)
keywords: vis_sdr.chm11162205
f1_keywords:
- vis_sdr.chm11162205
ms.prod: visio
api_name:
- Visio.Selection.Resize
ms.assetid: 4fc41631-adb4-9c5a-570f-e8ccaa2701eb
ms.date: 06/08/2017
---


# Selection.Resize Method (Visio)

Resizes the selection by moving shape handles as specified.


## Syntax

 _expression_ . **Resize**( **_Direction_** , **_Distance_** , **_UnitCode_** )

 _expression_ A variable that represents a **[Selection](selection-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Direction_|Required| **[VisResizeDirection](visresizedirection-enumeration-visio.md)**|The direction that corresponds to the shape handle. See Remarks for possible values.|
| _Distance_|Required| **Double**|The distance to move the selection edge or corner, where positive values move outward, and negative values move inward.|
| _UnitCode_|Required| **[VisUnitCodes](visunitcodes-enumeration-visio.md)**|The unit of measure for the resize distance.|

### Return Value

 **Nothing**


## Remarks

The  _Direction_ parameter must be one of the following **VisResizeDirection** constants.



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
When you resize a selection in a diagonal direction (that is, NE, NW, SE, or SW), the specified distance is applied along both the horizontal and vertical axes (as opposed to along the compass direction).


