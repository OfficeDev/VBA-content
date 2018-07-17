---
title: DataLabel.ShowSeriesName Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.DataLabel.ShowSeriesName
ms.assetid: 5d6eac40-c951-763d-7b1d-f7e69ea88407
ms.date: 06/08/2017
---


# DataLabel.ShowSeriesName Property (PowerPoint)

 **True** to show the series name for the data labels on a chart. **False** to hide the series name. Read/write **Boolean**.


## Syntax

 _expression_. **ShowSeriesName**

 _expression_ A variable that represents a **[DataLabel](datalabel-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example enables the series name to be shown for the data labels of the first series on the first chart.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.SeriesCollection(1).DataLabels. _
            ShowSeriesName = True
    End If
End With
```


## See also


#### Concepts


[DataLabel Object](datalabel-object-powerpoint.md)

