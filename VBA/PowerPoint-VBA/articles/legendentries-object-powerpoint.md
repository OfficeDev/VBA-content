---
title: LegendEntries Object (PowerPoint)
keywords: vbapp10.chm710000
f1_keywords:
- vbapp10.chm710000
ms.prod: powerpoint
api_name:
- PowerPoint.LegendEntries
ms.assetid: ac65aeaa-8a1c-57d7-499f-1c0b57dd02fd
ms.date: 06/08/2017
---


# LegendEntries Object (PowerPoint)

A collection of all the  **[LegendEntry](legendentry-object-powerpoint.md)** objects in the specified chart legend.


## Remarks

 Each legend entry has two parts:




- The text of the entry, which is the name of the series or trendline associated with the legend entry.
    
- The entry marker, which visually links the legend entry with its associated series or trendline in the chart.
    


The formatting properties for the entry marker and its associated series or trendline are contained in the  **[LegendKey](legendkey-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **[LegendEntries](legend-legendentries-method-powerpoint.md)** method to return the **LegendEntries** collection. The following example loops through the collection of legend entries for the first chart in the active document and changes their font color.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Legend

            For i = 1 To .LegendEntries.Count

                .LegendEntries(i).Font.ColorIndex = 5

            Next

        End With

    End If

End With


```




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use  **[LegendEntries](legend-legendentries-method-powerpoint.md)** ( _index_ ), where _index_ is the legend entry index number, to return a single **LegendEntry** object. You cannot return legend entries by name.

The index number represents the position of the legend entry in the legend.  `LegendEntries(1)` is at the top of the legend; `LegendEntries(LegendEntries.Count)` is at the bottom. The following example changes the font style for the text of the legend entry at the top of the legend (this is usually the legend for series one) for the first chart in the active document to italic.




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

