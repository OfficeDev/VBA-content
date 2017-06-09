---
title: LegendEntry Object (PowerPoint)
keywords: vbapp10.chm711000
f1_keywords:
- vbapp10.chm711000
ms.prod: powerpoint
api_name:
- PowerPoint.LegendEntry
ms.assetid: c92ddccd-92a3-bec9-cdcd-efd82c77706b
ms.date: 06/08/2017
---


# LegendEntry Object (PowerPoint)

Represents a legend entry in a chart legend.


## Remarks

 The **LegendEntry** object is a member of the **[LegendEntries](legendentries-object-powerpoint.md)** collection. The **LegendEntries** collection contains all the **LegendEntry** objects in the legend.

 Each legend entry has two parts:




- The text of the entry, which is the name of the series or trendline associated with the legend entry.
    
- The entry marker, which visually links the legend entry with its associated series or trendline in the chart.
    


The formatting properties for the entry marker and its associated series or trendline are contained in the  **[LegendKey](legendkey-object-powerpoint.md)** object.

The text of a legend entry cannot be changed.  **LegendEntry** objects support font formatting, and they can be deleted. No pattern formatting is supported for legend entries. The position and size of entries is fixed.

There is no direct way to return the series or trendline that corresponds to the legend entry.

After legend entries have been deleted, the only way to restore them is to remove and re-create the legend that contained them by setting the  **[HasLegend](chart-haslegend-property-powerpoint.md)** property for the chart to **False** and then back to **True**.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use  **[LegendEntries](legend-legendentries-method-powerpoint.md)** ( _index_ ), where _index_ is the legend entry index number, to return a single **LegendEntry** object. You cannot return legend entries by name.

The index number represents the position of the legend entry in the legend.  `LegendEntries(1)` is at the top of the legend, and `LegendEntries(LegendEntries.Count)` is at the bottom. The following example changes the font for the text of the legend entry at the top of the legend (this is usually the legend for series one) for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Legend.LegendEntries(1).Font.Italic = True

    End If

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

