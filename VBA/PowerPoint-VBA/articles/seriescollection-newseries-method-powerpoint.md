---
title: SeriesCollection.NewSeries Method (PowerPoint)
keywords: vbapp10.chm66653
f1_keywords:
- vbapp10.chm66653
ms.prod: powerpoint
api_name:
- PowerPoint.SeriesCollection.NewSeries
ms.assetid: 37a94558-02d9-7f0b-e881-0d9c5a9d4787
ms.date: 06/08/2017
---


# SeriesCollection.NewSeries Method (PowerPoint)

Creates a new series.


## Syntax

 _expression_. **NewSeries**

 _expression_ A variable that represents a **[SeriesCollection](seriescollection-object-powerpoint.md)** object.


### Return Value

A  **[Series](series-object-powerpoint.md)** object that represents the new series.


## Remarks

This method is not available for PivotChart charts.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example adds a new series to the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        Set ns = .Chart.SeriesCollection.NewSeries

    End If

End With
```


## See also


#### Concepts


[SeriesCollection Object](seriescollection-object-powerpoint.md)

