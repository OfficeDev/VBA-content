---
title: LegendKey Object (Word)
keywords: vbawd10.chm4062
f1_keywords:
- vbawd10.chm4062
ms.prod: word
api_name:
- Word.LegendKey
ms.assetid: 07578528-3e73-7898-47dc-296aefb854f0
ms.date: 06/08/2017
---


# LegendKey Object (Word)

Represents a legend key in a chart legend.


## Remarks

 Each legend key is a graphic that visually links a legend entry with its associated series or trendline in the chart. The legend key is linked to its associated series or trendline in such a way that changing the formatting of one simultaneously changes the formatting of the other.


## Example

Use the  **[LegendKey](legendentry-legendkey-property-word.md)** property to return the **LegendKey** object. The following example changes the marker background color for the legend entry at the top of the legend for the first chart in the active document. This simultaneously changes the format of every point in the series associated with this legend entry. The associated series must support data markers.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Legend.LegendEntries(1).LegendKey _ 
 .MarkerBackgroundColorIndex = 5 
 End If 
End With 

```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


