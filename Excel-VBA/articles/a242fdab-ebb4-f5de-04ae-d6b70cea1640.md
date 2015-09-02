
# LegendEntry Object

 **Last modified:** July 28, 2015

Represents a legend entry in the specified chart legend. The  **LegendEntry** object is a member of the ** [LegendEntries](98f5f860-7648-e3a6-f2af-985b383fed24.md)**collection, which contains all the  **LegendEntry** objects in the legend.

Each legend entry has two parts: the text of the entry, which is the name of the series associated with the entry; and an entry marker, which visually links the legend entry with its associated series or trendline in the chart. Formatting properties for the entry marker and its associated series or trendline are contained in the  ** [LegendKey](ab90cb64-1f81-dfcb-7542-cba68964acba.md)**object.

You cannot change the text of a legend entry.  **LegendEntry** objects support font formatting, and they can be deleted. No pattern formatting is supported for legend entries. The position and size of entries is fixed.

## Using the LegendEntry Object

Use  **LegendEntries**( _index_), where  _index_ is the legend entry's index number, to return a single **LegendEntry** object. You cannot return legend entries by name.

The index number represents the position of the legend entry in the legend.  `LegendEntries(1)` is at the top of the legend, and is at the top of the legend, and `LegendEntries(LegendEntries.Count)` is at the bottom. The following example changes the font style for the text of the legend entry at the top of the legend (this is usually the legend for series one).




```
myChart.Legend.LegendEntries(1).Font.Italic = True
```


## Remarks

There's no direct way to return the series or trendline that corresponds to a particular legend entry.

After legend entries have been deleted, the only way to restore them is to remove and then recreate the legend that contained them by setting the  ** [HasLegend](b4dbef39-9d83-2f6e-fe06-8ca38cceeeec.md)**property for the chart to  **False** and then back to **True**.

