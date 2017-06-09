---
title: TickLabels Object (Word)
keywords: vbawd10.chm2549
f1_keywords:
- vbawd10.chm2549
ms.prod: word
api_name:
- Word.TickLabels
ms.assetid: d94e90dc-0b0e-f4af-078e-6f2b97729db5
ms.date: 06/08/2017
---


# TickLabels Object (Word)

Represents the tick-mark labels associated with tick marks on a chart axis.


## Remarks

This object is not a collection. There is no object that represents a single tick-mark label; you must return all the tick-mark labels as a unit.

Tick-mark label text for the category axis comes from the name of the associated category in the chart. The default tick-mark label text for the category axis is the number that indicates the position of the category relative to the left end of this axis. To change the number of unlabeled tick marks between tick-mark labels, you must change the  **[TickLabelSpacing](axis-ticklabelspacing-property-word.md)** property for the category axis.

Tick-mark label text for the value axis is calculated based on the  **[MajorUnit](axis-majorunit-property-word.md)** , **[MinimumScale](axis-minimumscale-property-word.md)** , and **[MaximumScale](axis-maximumscale-property-word.md)** properties of the value axis. To change the tick-mark label text for the value axis, you must change the values of these properties.


## Example

Use the  **[TickLabels](axis-ticklabels-property-word.md)** property to return the **TickLabels** object. The following example sets the number format for the tick-mark labels on the value axis for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Axes(xlValue).TickLabels.NumberFormat = "0.00" 
 End If 
End With
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


