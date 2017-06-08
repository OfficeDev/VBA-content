---
title: ShapeRange.ScaleHeight Method (PowerPoint)
keywords: vbapp10.chm548010
f1_keywords:
- vbapp10.chm548010
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.ScaleHeight
ms.assetid: 3e86cfd8-1df6-a164-d19b-8d53b7b52dc0
ms.date: 06/08/2017
---


# ShapeRange.ScaleHeight Method (PowerPoint)

Scales the height of the shapes in the range by a specified factor. 


## Syntax

 _expression_. **ScaleHeight**( **_Factor_**, **_RelativeToOriginalSize_**, **_fScale_** )

 _expression_ A variable that represents a **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Factor_|Required|**Single**|Specifies the ratio between the height of the shapes after you resize them and their current or original height. For example, to make shapes 50 percent larger, specify 1.5 for this parameter.|
| _RelativeToOriginalSize_|Required|**MsoTriState**|Specifies whether shapes are scaled relative to their current or original sizes.|
| _fScale_|Optional|**MsoScaleFrom**|The parts of the shapes that retain their position when the shapes are scaled.|

### Return Value

Nothing


## Remarks

For pictures and OLE objects, you can indicate whether you want to scale the shapes relative to their original sizes or relative to their current sizes. Shapes other than pictures and OLE objects are always scaled relative to their current height.

The  _RelativeToOriginalSize_ parameter value can be one of the following **MsoTriState** constants. You can specify **msoTrue** for this argument only if the specified shapes are pictures or OLE objects.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|Scales the shapes relative to their current sizes. |
|**msoTrue**|Scales the shapes relative to their original sizes. |
The  _fScale_ parameter value can be one of the following **MsoTriState** constants. The default is **msoScaleFromTopLeft**.


||
|:-----|
|**msoScaleFromBottomRight**|
|**msoScaleFromMiddle**|
|**msoScaleFromTopLeft**|

## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

