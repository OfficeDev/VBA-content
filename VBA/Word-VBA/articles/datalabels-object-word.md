---
title: DataLabels Object (Word)
ms.prod: word
api_name:
- Word.DataLabels
ms.assetid: a7676f18-b1f2-1e11-9489-863cb85c1669
ms.date: 06/08/2017
---


# DataLabels Object (Word)

A collection of all the  **[DataLabel](datalabel-object-word.md)** objects for the specified series.


## Remarks

 Each **DataLabel** object represents a data label for a point or trendline. For a series without definable points (such as an area series), the **DataLabels** collection contains a single data label.


## Example

Use the  **[DataLabels](series-datalabels-method-word.md)** method to return the **DataLabels** collection. The following example sets the number format for data labels on the first series of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With Chart.SeriesCollection(1) 
 .HasDataLabels = True 
 .DataLabels.NumberFormat = "##.##" 
 End With 
 End If 
End With 

```

Use  **[DataLabels](series-datalabels-method-word.md)** ( _Index_ ), where _Index_ is the data label index number, to return a single **DataLabel** object. The following example sets the number format for the fifth data label in the first series of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With Chart.SeriesCollection(1).DataLabels(5) 
 .NumberFormat = "0.000" 
 End With 
 End If 
End With
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


