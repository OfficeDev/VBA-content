---
title: ChartGroup.SeriesCollection Method (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroup.SeriesCollection
ms.assetid: 5d20f5b2-cd4c-06b6-a49c-0ab331157b2f
ms.date: 06/08/2017
---


# ChartGroup.SeriesCollection Method (PowerPoint)

Returns all the series in the chart group.


## Syntax

 _expression_. **SeriesCollection**( **_Index_** )

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-powerpoint.md)** object.


### Return Value

A  **[SeriesCollection](seriescollection-object-powerpoint.md)** object that represents all the series in the chart group.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example turns on data labels for the first series of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.ChartGroups(1). _
            SeriesCollection(1).HasDataLabels = True
    End If
End With
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-powerpoint.md)

