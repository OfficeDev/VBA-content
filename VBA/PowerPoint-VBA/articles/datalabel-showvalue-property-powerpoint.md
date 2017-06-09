---
title: DataLabel.ShowValue Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.DataLabel.ShowValue
ms.assetid: 2d4ca0a0-9b2c-7477-214b-322283e2c082
ms.date: 06/08/2017
---


# DataLabel.ShowValue Property (PowerPoint)

 **True** to display a specified chart's data label values. **False** to hide the values. Read/write **Boolean**.


## Syntax

 _expression_. **ShowValue**

 _expression_ A variable that represents a **[DataLabel](datalabel-object-powerpoint.md)** object.


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


[DataLabel Object](datalabel-object-powerpoint.md)

