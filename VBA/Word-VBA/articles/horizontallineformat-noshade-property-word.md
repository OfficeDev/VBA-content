---
title: HorizontalLineFormat.NoShade Property (Word)
keywords: vbawd10.chm165543939
f1_keywords:
- vbawd10.chm165543939
ms.prod: word
api_name:
- Word.HorizontalLineFormat.NoShade
ms.assetid: 90728761-cdfa-fd2c-db00-44ca78a34017
ms.date: 06/08/2017
---


# HorizontalLineFormat.NoShade Property (Word)

 **True** if Microsoft Word draws the specified horizontal line without 3-D shading. Read/write **Boolean** .


## Syntax

 _expression_ . **NoShade**

 _expression_ An expression that returns a **[HorizontalLineFormat](horizontallineformat-object-word.md)** object.


## Remarks

You can only use this property with horizontal lines that are not based on an existing image file.


## Example

This example adds a horizontal line without any 3-D shading.


```vb
Selection.InlineShapes.AddHorizontalLineStandard 
ActiveDocument.InlineShapes(1) _ 
 .HorizontalLineFormat.NoShade = True
```


## See also


#### Concepts


[HorizontalLineFormat Object](horizontallineformat-object-word.md)

