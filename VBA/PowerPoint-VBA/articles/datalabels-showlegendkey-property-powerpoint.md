---
title: DataLabels.ShowLegendKey Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.DataLabels.ShowLegendKey
ms.assetid: 84080842-b3ab-0617-e0b2-04e865f968c9
ms.date: 06/08/2017
---


# DataLabels.ShowLegendKey Property (PowerPoint)

 **True** if the data label legend key is visible. Read/write **Boolean**.


## Syntax

 _expression_. **ShowLegendKey**

 _expression_ A variable that represents a **[DataLabels](datalabels-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the data labels for series one of the first chart in the active document to show values and the legend key.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.SeriesCollection(1).DataLabels. _
            ShowLegendKey = True
        .Chart.SeriesCollection(1).DataLabels.Type = xlShowValue
    End If
End With
```


## See also


#### Concepts


[DataLabels Object](datalabels-object-powerpoint.md)

