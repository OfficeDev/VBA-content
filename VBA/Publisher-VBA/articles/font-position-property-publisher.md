---
title: Font.Position Property (Publisher)
keywords: vbapb10.chm5373988
f1_keywords:
- vbapb10.chm5373988
ms.prod: publisher
api_name:
- Publisher.Font.Position
ms.assetid: 24573faf-1627-3b10-5a8e-2f76a9f8831d
ms.date: 06/08/2017
---


# Font.Position Property (Publisher)

Returns or sets a  **Variant** representing the font position relative to the baseline of the text in the specified range. Positive values move the text above the normal baseline, negative values move the text below the baseline. Indeterminate values are returned as -9999.0. Read/write.


## Syntax

 _expression_. **Position**

 _expression_A variable that represents a  **Font** object.


## Example

This example adjusts the text in the second story to 5 points below the baseline.


```vb
Sub Position() 
 
 Application.ActiveDocument.Stories(2).TextRange.Font.Position = -5 
 
End Sub
```


