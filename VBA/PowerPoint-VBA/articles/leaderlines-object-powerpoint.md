---
title: LeaderLines Object (PowerPoint)
keywords: vbapp10.chm708000
f1_keywords:
- vbapp10.chm708000
ms.prod: powerpoint
api_name:
- PowerPoint.LeaderLines
ms.assetid: 2357c570-0f68-8bb4-910a-e88c00ed9884
ms.date: 06/08/2017
---


# LeaderLines Object (PowerPoint)

Represents leader lines on a chart. Leader lines connect data labels to data points.


## Remarks

 This object is not a collection; there is no object that represents a single leader line.

This object applies only to pie charts.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **[LeaderLines](series-leaderlines-property-powerpoint.md)** property to return the **LeaderLines** object. The following example adds data labels and blue leader lines to series one on the first chart in the active document. If no leader lines are visible, this example code will fail. In this situation, you can manually drag one of the data labels away from the pie chart to make a leader line show up.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.SeriesCollection(1)

            .HasDataLabels = True

            .DataLabels.Position = xlLabelPositionBestFit

            .HasLeaderLines = True

            .LeaderLines.Border.ColorIndex = 5

        End With

    End If

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

