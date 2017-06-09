---
title: DataLabels.ShowValue Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.DataLabels.ShowValue
ms.assetid: e0c739f6-286b-1267-49c0-484b7d1bca16
ms.date: 06/08/2017
---


# DataLabels.ShowValue Property (PowerPoint)

 **True** to display the data label values for a specified chart. **False** to hide the values. Read/write **Boolean**.


## Syntax

 _expression_. **ShowValue**

 _expression_ A variable that represents a **[DataLabels](datalabels-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example enables the value to be shown for the data labels of the first series in the first chart.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.SeriesCollection(1).DataLabels. _
            ShowValue = True
    End If
End With
```


## See also


#### Concepts


[DataLabels Object](datalabels-object-powerpoint.md)

