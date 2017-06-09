---
title: Trendline.DisplayRSquared Property (PowerPoint)
keywords: vbapp10.chm65725
f1_keywords:
- vbapp10.chm65725
ms.prod: powerpoint
api_name:
- PowerPoint.Trendline.DisplayRSquared
ms.assetid: e2899b19-c35f-b648-42af-d7fd75d51653
ms.date: 06/08/2017
---


# Trendline.DisplayRSquared Property (PowerPoint)

 **True** if the R-squared value of the trendline is displayed on the chart (in the same data label as the equation). Read/write **Boolean**.


## Syntax

 _expression_. **DisplayRSquared**

 _expression_ A variable that represents a **[Trendline](trendline-object-powerpoint.md)** object.


## Remarks

Setting this property to  **True** automatically turns on data labels.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example displays the R-squared value and equation for trendline one of the first chart in the active document. You should run the example on a 2-D column chart that has a trendline for the first series.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.SeriesCollection(1).Trendlines(1)

            .DisplayRSquared = True

            .DisplayEquation = True

        End With

    End If

End With


```


## See also


#### Concepts


[Trendline Object](trendline-object-powerpoint.md)

