---
title: DataLabels.ShowPercentage Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.DataLabels.ShowPercentage
ms.assetid: c125433f-7166-871e-f433-9320b1613a70
ms.date: 06/08/2017
---


# DataLabels.ShowPercentage Property (PowerPoint)

 **True** to display the percentage value for the data labels on a chart. **False** to hide the value. Read/write **Boolean**.


## Syntax

 _expression_. **ShowPercentage**

 _expression_ A variable that represents a **[DataLabels](datalabels-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example enables the percentage value to be shown for the data labels of the first series on the first chart.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.SeriesCollection(1).DataLabels. _
            ShowPercentage = True
    End If
End With
```


## See also


#### Concepts


[DataLabels Object](datalabels-object-powerpoint.md)

