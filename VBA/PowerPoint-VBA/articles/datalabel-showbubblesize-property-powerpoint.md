---
title: DataLabel.ShowBubbleSize Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.DataLabel.ShowBubbleSize
ms.assetid: a6bbef53-ff4a-7766-2a6b-f9b5907bebf3
ms.date: 06/08/2017
---


# DataLabel.ShowBubbleSize Property (PowerPoint)

 **True** to show the bubble size for the data labels on a chart. **False** to hide the bubble size. Read/write **Boolean**.


## Syntax

 _expression_. **ShowBubbleSize**

 _expression_ A variable that represents a **[DataLabel](datalabel-object-powerpoint.md)** object.


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


[DataLabel Object](datalabel-object-powerpoint.md)

