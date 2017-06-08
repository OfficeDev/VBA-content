---
title: HorizontalLineFormat Object (Word)
keywords: vbawd10.chm2526
f1_keywords:
- vbawd10.chm2526
ms.prod: word
api_name:
- Word.HorizontalLineFormat
ms.assetid: 55296fc7-9b7e-dcdb-00e0-901015cf0efb
ms.date: 06/08/2017
---


# HorizontalLineFormat Object (Word)

Represents horizontal line formatting.


## Remarks

Use the  **HorizontalLineFormat** property to return a **HorizontalLineFormat** object. This example sets the alignment for a new horizontal line.


```
Selection.InlineShapes.AddHorizontalLineStandard 
ActiveDocument.InlineShapes(1) _ 
 .HorizontalLineFormat.Alignment = _ 
 wdHorizontalLineAlignLeft
```

This example adds a horizontal line without any 3-D shading.




```vb
Selection.InlineShapes.AddHorizontalLineStandard 
ActiveDocument.InlineShapes(1) _ 
 .HorizontalLineFormat.NoShade = True
```

This example adds a horizontal line and sets its length to 50% of the window width.




```
Selection.InlineShapes.AddHorizontalLineStandard 
ActiveDocument.InlineShapes(1) _ 
 .HorizontalLineFormat.PercentWidth = 50
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


