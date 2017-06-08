---
title: ChartTitle Object (Word)
keywords: vbawd10.chm996
f1_keywords:
- vbawd10.chm996
ms.prod: word
api_name:
- Word.ChartTitle
ms.assetid: fc8ca540-0a29-123b-2fdf-b16aaa1f940c
ms.date: 06/08/2017
---


# ChartTitle Object (Word)

Represents the chart title.


## Remarks

Use the  **[ChartTitle](chart-charttitle-property-word.md)** property to return the **ChartTitle** object.

The  **ChartTitle** object does not exist and cannot be used unless the **[HasTitle](chart-hastitle-property-word.md)** property for the chart is **True** .


## Example

 The following example adds a title to embedded chart one on the worksheet named "Sheet1."


```vb
With ActiveDocument.InlineShapes(1).Chart 
 .HasTitle = True 
 .ChartTitle.Text = "February Sales" 
End With
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

