---
title: ShapeRange.ZOrder Method (Word)
keywords: vbawd10.chm162856988
f1_keywords:
- vbawd10.chm162856988
ms.prod: word
api_name:
- Word.ShapeRange.ZOrder
ms.assetid: 7f9a1a08-ac21-8866-9bf7-6a850200e2fd
ms.date: 06/08/2017
---


# ShapeRange.ZOrder Method (Word)

Moves the specified shape range in front of or behind other shapes in the collection (that is, changes the shape range's position in the z-order).


## Syntax

 _expression_ . **ZOrder**( **_ZOrderCmd_** )

 _expression_ An expression that returns a **[ShapeRange](shaperange-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ZOrderCmd_|Required| **MsoZOrderCmd**|Specifies where to move the specified shape range relative to the other shapes.|

### Return Value

Nothing


## Remarks

Use the  **[ZOrderPosition](shaperange-zorderposition-property-word.md)** property to determine a shape range's current position in the z-order.


## See also


#### Concepts


[ShapeRange Collection Object](shaperange-object-word.md)

