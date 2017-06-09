---
title: Series.LeaderLines Property (PowerPoint)
keywords: vbapp10.chm67202
f1_keywords:
- vbapp10.chm67202
ms.prod: powerpoint
api_name:
- PowerPoint.Series.LeaderLines
ms.assetid: f5c706e0-c6df-ae45-9f34-b7f6b4200326
ms.date: 06/08/2017
---


# Series.LeaderLines Property (PowerPoint)

Returns the leader lines for the series. Read-only  **[LeaderLines](leaderlines-object-powerpoint.md)**.


## Syntax

 _expression_. **LeaderLines**

 _expression_ A variable that represents a **[Series](series-object-powerpoint.md)** object.


## Remarks

This property applies only to pie charts.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example adds data labels and blue leader lines to series one on the first pie chart in the active document. If no leader lines are visible, this example code will fail. In this situation, you can manually drag one of the data labels away from the pie chart to make a leader line show up.




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


[Series Object](series-object-powerpoint.md)

