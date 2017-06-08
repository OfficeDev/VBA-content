---
title: View.ShadeEditableRanges Property (Word)
keywords: vbawd10.chm161808431
f1_keywords:
- vbawd10.chm161808431
ms.prod: word
api_name:
- Word.View.ShadeEditableRanges
ms.assetid: f079c812-024d-6568-4643-4c2df96fd59d
ms.date: 06/08/2017
---


# View.ShadeEditableRanges Property (Word)

Returns or sets a  **Long** that represents whether shading is applied to the ranges in a document for which users have permission to modify. .


## Syntax

 _expression_ . **ShadeEditableRanges**

 _expression_ An expression that returns a **[View](view-object-word.md)** object.


## Remarks

 **True** shades the ranges in a document that users can modify. Range shading is on by default. When range shading is on, or when you set the property to **True** , the **ShadeEditableRanges** property returns a value of -1. When you set the **ShadeEditableRanges** property to **False** it returns a value of 0. The values have no meaning beyond indicating whether the property is **True** or **False** .


## Example

The following example shades all ranges for which users have permission to modify.


```vb
ActiveWindow.View.ShadeEditableRanges = True
```


## See also


#### Concepts


[View Object](view-object-word.md)

