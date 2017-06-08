---
title: ShapeRange.Align Method (Word)
keywords: vbawd10.chm162856970
f1_keywords:
- vbawd10.chm162856970
ms.prod: word
api_name:
- Word.ShapeRange.Align
ms.assetid: 99cf934c-0a65-b283-f7a5-28674e5cb39f
ms.date: 06/08/2017
---


# ShapeRange.Align Method (Word)

Aligns the shapes in the specified range of shapes.


## Syntax

 _expression_ . **Align**( **_Align_** , **_RelativeTo_** )

 _expression_ Required. A variable that represents a **[ShapeRange](shaperange-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Align_|Required| **MsoAlignCmd**|Specifies the way the shapes in the specified shape range are to be aligned.|
| _RelativeTo_|Required| **Long**| **True** to align shapes relative to the edge of the document. **False** to align shapes relative to one another.|

## Example

This example aligns the left edges of all the shapes in the selection of shapes in the active document with the left edge of the leftmost shape in the range.


```vb
Set myShapeRange = Selection.ShapeRange 
myShapeRange.Align msoAlignLefts, False
```


## See also


#### Concepts


[ShapeRange Collection Object](shaperange-object-word.md)

