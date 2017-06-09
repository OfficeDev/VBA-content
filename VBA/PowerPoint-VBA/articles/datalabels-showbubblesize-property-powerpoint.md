---
title: DataLabels.ShowBubbleSize Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.DataLabels.ShowBubbleSize
ms.assetid: 78cb2f6f-f13c-9cc6-9842-ba8000273165
ms.date: 06/08/2017
---


# DataLabels.ShowBubbleSize Property (PowerPoint)

 **True** to show the bubble size for the data labels on a chart. **False** to hide the bubble size. Read/write **Boolean**.


## Syntax

 _expression_. **ShowBubbleSize**

 _expression_ A variable that represents a **[DataLabels](datalabels-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example shows the bubble size for the data labels of the first series on the first chart.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.SeriesCollection(1).DataLabels. _
            ShowBubbleSize = True
    End If
End With
```


## See also


#### Concepts


[DataLabels Object](datalabels-object-powerpoint.md)

