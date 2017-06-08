---
title: LegendEntry Object (Word)
keywords: vbawd10.chm73
f1_keywords:
- vbawd10.chm73
ms.prod: word
api_name:
- Word.LegendEntry
ms.assetid: 9f793578-cb9b-faa1-f0a1-ea0f9e90dc6f
ms.date: 06/08/2017
---


# LegendEntry Object (Word)

Represents a legend entry in a chart legend.


## Remarks

 The **LegendEntry** object is a member of the **[LegendEntries](legendentries-object-word.md)** collection. The **LegendEntries** collection contains all the **LegendEntry** objects in the legend.

 Each legend entry has two parts:




- The text of the entry, which is the name of the series or trendline associated with the legend entry.
    
- The entry marker, which visually links the legend entry with its associated series or trendline in the chart.
    


The formatting properties for the entry marker and its associated series or trendline are contained in the  **[LegendKey](legendkey-object-word.md)** object.

The text of a legend entry cannot be changed.  **LegendEntry** objects support font formatting, and they can be deleted. No pattern formatting is supported for legend entries. The position and size of entries is fixed.

There is no direct way to return the series or trendline that corresponds to the legend entry.

After legend entries have been deleted, the only way to restore them is to remove and re-create the legend that contained them by setting the  **[HasLegend](chart-haslegend-property-word.md)** property for the chart to **False** and then back to **True** .


## Example

Use  **[LegendEntries](legend-legendentries-method-word.md)** ( _index_ ), where _index_ is the legend entry index number, to return a single **LegendEntry** object. You cannot return legend entries by name.

The index number represents the position of the legend entry in the legend.  `LegendEntries(1)` is at the top of the legend, and `LegendEntries(LegendEntries.Count)` is at the bottom. The following example changes the font for the text of the legend entry at the top of the legend (this is usually the legend for series one) for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Legend.LegendEntries(1).Font.Italic = True 
 End If 
End With
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


