---
title: LegendEntries Collection (Excel)
keywords: vbagr10.chm5207606
f1_keywords:
- vbagr10.chm5207606
ms.prod: excel
ms.assetid: 98f5f860-7648-e3a6-f2af-985b383fed24
ms.date: 06/08/2017
---


# LegendEntries Collection (Excel)

A collection of all the  **[LegendEntry](legendentry-object.md)** objects in the specified chart legend. Each legend entry has two parts: the text of the entry, which is the name of the series or trendline associated with the entry; and the entry marker, which visually links the legend entry with its associated series or trendline in the chart. The formatting properties for the entry marker and its associated series or trendline are contained in the  **[LegendKey](legendkey-object.md)** object.


## Using the LegendEntries Collection

Use the  **LegendEntries** method to return the **LegendEntries** collection. The following example loops through the collection of legend entries in the chart and changes their font color to blue.


```vb
With myChart.Legend 
 For i = 1 To .LegendEntries.Count 
 .LegendEntries(i).Font.ColorIndex = 5 
 Next 
End With
```

Use  **LegendEntries**( _index_), where  _index_ is the legend entry's index number, to return a single **LegendEntry** object. You cannot return legend entries by name.

The index number represents the position of the legend entry in the legend.  `LegendEntries(1)` is at the top of the legend; `LegendEntries(LegendEntries.Count)` is at the bottom. The following example changes the font style to italic for the text of the legend entry at the top of the legend (this is usually the legend for series one) in is at the top of the legend; `LegendEntries(LegendEntries.Count)` is at the bottom. The following example changes the font style to italic for the text of the legend entry at the top of the legend (this is usually the legend for series one) in `myChart`.




```vb
myChart.Legend.LegendEntries(1).Font.Italic = True
```


