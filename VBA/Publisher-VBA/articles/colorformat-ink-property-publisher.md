---
title: ColorFormat.Ink Property (Publisher)
keywords: vbapb10.chm2555911
f1_keywords:
- vbapb10.chm2555911
ms.prod: publisher
api_name:
- Publisher.ColorFormat.Ink
ms.assetid: 53851337-fdce-7b72-5626-50bce370457b
ms.date: 06/08/2017
---


# ColorFormat.Ink Property (Publisher)

Returns or sets a  **Long** indicating whether the specified color is a spot color, and if so, the spot plate to which it belongs. Valid values are **pbInkNone** (default; meaning that the color is not a spot color) or a number between 1 and _n_ where _n_ is the number of spot plates. Read/write.


## Syntax

 _expression_. **Ink**

 _expression_A variable that represents an  **ColorFormat** object.


### Return Value

Long


## Example

The following example specifies that the color of the first text range on page one of the active publication should be assigned to spot plate two.


```vb
ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Font.Color.Ink = 2
```


