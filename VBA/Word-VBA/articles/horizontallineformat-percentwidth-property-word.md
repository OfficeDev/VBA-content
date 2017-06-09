---
title: HorizontalLineFormat.PercentWidth Property (Word)
keywords: vbawd10.chm165543938
f1_keywords:
- vbawd10.chm165543938
ms.prod: word
api_name:
- Word.HorizontalLineFormat.PercentWidth
ms.assetid: 40c51a99-aeda-9250-bb94-ee983ef3c33c
ms.date: 06/08/2017
---


# HorizontalLineFormat.PercentWidth Property (Word)

Returns or sets the length of the specified horizontal line expressed as a percentage of the window width. Read/write  **Single** .


## Syntax

 _expression_ . **PercentWidth**

 _expression_ An expression that returns a **[HorizontalLineFormat](horizontallineformat-object-word.md)** object.


## Remarks

Setting this property also sets the  **[WidthType](horizontallineformat-widthtype-property-word.md)** property to **wdHorizontalLinePercentWidth** .


## Example

This example adds a horizontal line and sets its length to 50% of the window width.


```
Selection.InlineShapes.AddHorizontalLineStandard 
ActiveDocument.InlineShapes(1) _ 
 .HorizontalLineFormat.PercentWidth = 50
```


## See also


#### Concepts


[HorizontalLineFormat Object](horizontallineformat-object-word.md)

