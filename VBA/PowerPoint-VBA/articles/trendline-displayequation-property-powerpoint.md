---
title: Trendline.DisplayEquation Property (PowerPoint)
keywords: vbapp10.chm65726
f1_keywords:
- vbapp10.chm65726
ms.prod: powerpoint
api_name:
- PowerPoint.Trendline.DisplayEquation
ms.assetid: dad5ea14-3165-df06-33b6-b90ddedaab39
ms.date: 06/08/2017
---


# Trendline.DisplayEquation Property (PowerPoint)

 **True** if the equation for the trendline is displayed on the chart (in the same data label as the R-squared value). Read/write **Boolean**.


## Syntax

 _expression_. **DisplayEquation**

 _expression_ A variable that represents a **[Trendline](trendline-object-powerpoint.md)** object.


## Remarks

Setting this property to  **True** automatically enables data labels.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example displays the R-squared value and equation for the first trendline of the first chart in the active document. You should run the example on a 2-D column chart that has a trendline for the first series.




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

