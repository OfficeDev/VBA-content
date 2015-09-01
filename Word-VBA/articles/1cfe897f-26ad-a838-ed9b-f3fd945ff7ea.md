
# Trendline Object (Word)

 **Last modified:** July 28, 2015

Represents a trendline in a chart.

## Remarks

A trendline shows the trend, or direction, of data in a series. The  **Trendline** object is a member of the ** [Trendlines](06c20a75-4afc-03f5-1eec-eee1559d3f52.md)** collection. The **Trendlines** collection contains all the **Trendline** objects for a single series.


## Example

Use  ** [Trendlines](300dca01-097f-8a3d-4f63-a1841a92098e.md)**(Index), where Index is the trendline index number, to return a single  **Trendline** object.

The index number denotes the order in which the trendlines were added to the series.  `Trendlines(1)` is the first trendline added to the series, and `Trendlines(Trendlines.Count)` is the last one added.

The following example changes the trendline type for the first series of the first chart in the active document. If the series has no trendline, this example will fail.




```
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).Trendlines(1).Type = xlMovingAvg 
 End If 
End With
```


## See also


#### Concepts


 [Word Object Model Reference](be452561-b436-bb9b-6f94-3faa9a74a6fd.md)
#### Other resources


 [Trendline Object Members](02d1ce95-ff74-859a-70b2-cd914c334083.md)
