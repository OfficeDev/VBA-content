---
title: SeriesCollection Object (Word)
keywords: vbawd10.chm2295
f1_keywords:
- vbawd10.chm2295
ms.prod: word
api_name:
- Word.SeriesCollection
ms.assetid: 785d61ff-96c9-b9b0-ed98-e992d9adeda6
ms.date: 06/08/2017
---


# SeriesCollection Object (Word)

Represents a collection of all the  **[Series](series-object-word.md)** objects in the specified chart or chart group.


## Remarks

Use the  **[SeriesCollection](chart-seriescollection-method-word.md)** method to return the **SeriesCollection** collection.


## Example

 Use the **[Extend](seriescollection-extend-method-word.md)** method to extend an existing series. The following example adds the data in cells C6:C10 in the chart's worksheet to an existing series in the series collection of the chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection.Extend "='Sheet1'!$C$6:$C$10" 
 End If 
End With
```

Use the  **[Add](seriescollection-add-method-word.md)** method to create a new series and add it to the chart. The following example adds the data from cells D1:D5 in the chart's worksheet as a new series to the chart.




```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection.Add "='Sheet1'!$D$1:$D$5" 
 End If 
End With
```

Use  **SeriesCollection** ( _Index_ ), where _Index_ is the series index number or name, to return a single **Series** object. The following example sets the color of the interior for the first series in embedded chart one on Sheet1.




```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).Interior.Color = RGB(255, 0, 0) 
 End If 
End With
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

