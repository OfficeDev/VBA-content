
# LegendEntry Object (PowerPoint)

 **Last modified:** July 28, 2015

Represents a legend entry in a chart legend.

## Remarks

 The **LegendEntry** object is a member of the ** [LegendEntries](ac65aeaa-8a1c-57d7-499f-1c0b57dd02fd.md)** collection. The **LegendEntries** collection contains all the **LegendEntry** objects in the legend.

 Each legend entry has two parts:




- The text of the entry, which is the name of the series or trendline associated with the legend entry.
    
- The entry marker, which visually links the legend entry with its associated series or trendline in the chart.
    


The formatting properties for the entry marker and its associated series or trendline are contained in the  ** [LegendKey](98e8b9c3-b53e-9595-9389-6f92a6d730f4.md)** object.

The text of a legend entry cannot be changed.  **LegendEntry** objects support font formatting, and they can be deleted. No pattern formatting is supported for legend entries. The position and size of entries is fixed.

There is no direct way to return the series or trendline that corresponds to the legend entry.

After legend entries have been deleted, the only way to restore them is to remove and re-create the legend that contained them by setting the  ** [HasLegend](084f7de3-b0ed-d7b3-3b24-465e74afa167.md)** property for the chart to **False** and then back to **True**.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use  ** [LegendEntries](a6110ddf-76dd-efc9-c6ce-abb260f9534c.md)**( _index_), where  _index_ is the legend entry index number, to return a single **LegendEntry** object. You cannot return legend entries by name.

The index number represents the position of the legend entry in the legend.  `LegendEntries(1)` is at the top of the legend, and `LegendEntries(LegendEntries.Count)` is at the bottom. The following example changes the font for the text of the legend entry at the top of the legend (this is usually the legend for series one) for the first chart in the active document.




```
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Legend.LegendEntries(1).Font.Italic = True

    End If

End With
```


## See also


#### Concepts


 [PowerPoint Object Model Reference](00acd64a-5896-0459-39af-98df2849849e.md)
#### Other resources


 [LegendEntry Object Members](408ad572-e777-f74a-4ab9-d70b43901c7e.md)
