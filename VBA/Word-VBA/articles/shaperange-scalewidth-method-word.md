---
title: ShapeRange.ScaleWidth Method (Word)
keywords: vbawd10.chm162856984
f1_keywords:
- vbawd10.chm162856984
ms.prod: word
api_name:
- Word.ShapeRange.ScaleWidth
ms.assetid: 3742c9a7-7a97-9e24-299c-b7567eedb9d1
ms.date: 06/08/2017
---


# ShapeRange.ScaleWidth Method (Word)

Scales the width of a shape by a specified factor.


## Syntax

 _expression_ . **ScaleWidth**( **_Factor_** , **_RelativeToOriginalSize_** , **_Scale_** )

 _expression_ Required. A variable that represents a **[ShapeRange](shaperange-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Factor_|Required| **Single**|Specifies the ratio between the width of the shape after you resize it and the current or original width. For example, to make a rectangle 50 percent larger, specify 1.5 for this argument.|
| _RelativeToOriginalSize_|Required| **MsoTriState**| **True** to scale the shape relative to its original size. **False** to scale it relative to its current size. You can specify **True** for this argument only if the specified shape is a picture or an OLE object.|
| _Scale_|Optional| **MsoScaleFrom**|The part of the shape that retains its position when the shape is scaled.|

## Remarks

For pictures and OLE objects, you can indicate whether you want to scale the range of shapes relative to the original size or relative to the current size. Shapes other than pictures and OLE objects are always scaled relative to their current width.


## See also


#### Concepts


[ShapeRange Collection Object](shaperange-object-word.md)

