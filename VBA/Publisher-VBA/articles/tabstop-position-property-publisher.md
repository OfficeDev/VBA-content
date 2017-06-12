---
title: TabStop.Position Property (Publisher)
keywords: vbapb10.chm5636099
f1_keywords:
- vbapb10.chm5636099
ms.prod: publisher
api_name:
- Publisher.TabStop.Position
ms.assetid: 1ca7831a-6662-036e-8ba2-5784bc95fe8d
ms.date: 06/08/2017
---


# TabStop.Position Property (Publisher)

Returns or sets a  **Variant** representing the font position relative to the baseline of the text in the specified range. Positive values move the text above the normal baseline, negative values move the text below the baseline. Indeterminate values are returned as -9999.0. Read/write.


## Syntax

 _expression_. **Position**

 _expression_A variable that represents a  **TabStop** object.


## Example

This example adjusts the text in the second story to 5 points below the baseline.


```vb
Sub Position() 
 
 Application.ActiveDocument.Stories(2).TextRange.Font.Position = -5 
 
End Sub
```


