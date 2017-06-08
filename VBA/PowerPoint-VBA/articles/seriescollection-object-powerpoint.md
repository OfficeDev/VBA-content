---
title: SeriesCollection Object (PowerPoint)
keywords: vbapp10.chm717000
f1_keywords:
- vbapp10.chm717000
ms.prod: powerpoint
api_name:
- PowerPoint.SeriesCollection
ms.assetid: 6277f9e0-0198-0773-9c54-f2d009c0ba7a
ms.date: 06/08/2017
---


# SeriesCollection Object (PowerPoint)

Represents a collection of all the  **[Series](series-object-powerpoint.md)** objects in the specified chart or chart group.


## Remarks

Use the  **[SeriesCollection](http://msdn.microsoft.com/library/8adeb8b4-ba4f-6cdf-33bf-dceb1845dfb8%28Office.15%29.aspx)** method to return the **SeriesCollection** collection.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

 Use the **[Extend](http://msdn.microsoft.com/library/f5ac6da3-90c7-d938-9a95-e87d228d901d%28Office.15%29.aspx)** method to extend an existing series. The following example adds the data in cells C6:C10 in the chart's worksheet to an existing series in the series collection of the chart.




```
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection.Extend "='Sheet1'!$C$6:$C$10"

    End If

End With
```




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **[Add](http://msdn.microsoft.com/library/29dd05a7-a707-78ff-fc06-1085e065eb3c%28Office.15%29.aspx)** method to create a new series and add it to the chart. The following example adds the data from cells D1:D5 in the chart's worksheet as a new series to the chart.




```
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection.Add "='Sheet1'!$D$1:$D$5"

    End If

End With
```




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use  **SeriesCollection** ( _Index_ ), where _Index_ is the series index number or name, to return a single **Series** object. The following example sets the color of the interior for the first series in embedded chart one the active document.




```
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).Interior.Color = RGB(255, 0, 0)

    End If

End With
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/29dd05a7-a707-78ff-fc06-1085e065eb3c%28Office.15%29.aspx)|
|[Extend](http://msdn.microsoft.com/library/f5ac6da3-90c7-d938-9a95-e87d228d901d%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/ae34ad0d-1b0a-decb-24e8-3d1c51652f72%28Office.15%29.aspx)|
|[NewSeries](http://msdn.microsoft.com/library/37a94558-02d9-7f0b-e881-0d9c5a9d4787%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/c872de5e-2a1c-fe96-9966-28e7d30f46c2%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/527e7502-d84e-8884-b0df-7d44cbe89f3f%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/0d767309-d866-9ec5-5ff0-9c4b7e54c8fc%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/f5d40a16-5a35-3560-1f59-ffdba6d95807%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
