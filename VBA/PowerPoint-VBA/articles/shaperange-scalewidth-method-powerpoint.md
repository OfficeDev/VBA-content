---
title: ShapeRange.ScaleWidth Method (PowerPoint)
keywords: vbapp10.chm548011
f1_keywords:
- vbapp10.chm548011
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.ScaleWidth
ms.assetid: 868f56cb-6a3a-902e-b6a9-2a9229936b41
ms.date: 06/08/2017
---


# ShapeRange.ScaleWidth Method (PowerPoint)

Scales the width of the shapes in the range by a specified factor. 


## Syntax

 _expression_. **ScaleWidth**( **_Factor_**, **_RelativeToOriginalSize_**, **_fScale_** )

 _expression_ A variable that represents a **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Factor_|Required|**Single**|Specifies the ratio between the width of the shapes after you resize them and the current or original width. For example, to make all the shapes in the range 50 percent larger, specify 1.5 for this parameter.|
| _RelativeToOriginalSize_|Required|**MsoTriState**|Specifies whether shapes are scaled relative to their current or original size.|
| _fScale_|Optional|**MsoScaleFrom**|The parts of the shapes that retain their positions when the shapes are scaled.|

## Remarks

For pictures and OLE objects, you can indicate whether you want to scale the shapes relative to the original size or relative to the current size. Shapes other than pictures and OLE objects are always scaled relative to their current width.

The  _RelativeToOriginalSize_ parameter value can be one of the following **MsoTriState** constants. You can specify **msoTrue** for this parameter only if the specified shapes are pictures or OLE objects.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|Scales the shape relative to its current size. |
|**msoTrue**| Scales the shape relative to its original size.|
The  _fScale_ parameter value can be one of the following **MsoScaleFrom** constants. The default is **msoScaleFromTopLeft**.


||
|:-----|
|**msoScaleFromBottomRight**|
|**msoScaleFromMiddle**|
|**msoScaleFromTopLeft**|

## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

