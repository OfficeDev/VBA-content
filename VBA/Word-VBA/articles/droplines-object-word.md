---
title: DropLines Object (Word)
keywords: vbawd10.chm1602
f1_keywords:
- vbawd10.chm1602
ms.prod: word
api_name:
- Word.DropLines
ms.assetid: 4691b002-8512-7cd3-5a20-561232e18d88
ms.date: 06/08/2017
---


# DropLines Object (Word)

Represents the drop lines in a chart group.


## Remarks

Drop lines connect the points in the chart with the x-axis. Only line and area chart groups can have drop lines. This object is not a collection. There is no object that represents a single drop line; you either enable drop lines for all points in a chart group or you disable them.

If the  **[HasDropLines](chartgroup-hasdroplines-property-word.md)** property is **False** , most properties of the **DropLines** object are disabled.


## Example

Use the  **[DropLines](chartgroup-droplines-property-word.md)** property to return the **DropLines** object. The following example enables drop lines for chart group one of the first chart in the active document and then sets the drop line color to red.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.ChartGroups(1) 
 .HasDropLines = True 
 .DropLines.Border.ColorIndex = 3 
 End With 
 End If 
End With
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

