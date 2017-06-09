---
title: Trendline.Forward2 Property (PowerPoint)
keywords: vbapp10.chm68187
f1_keywords:
- vbapp10.chm68187
ms.prod: powerpoint
api_name:
- PowerPoint.Trendline.Forward2
ms.assetid: d5968c1f-de77-a03f-44b2-f91d6638a6ae
ms.date: 06/08/2017
---


# Trendline.Forward2 Property (PowerPoint)

Returns or sets the number of periods (or units on a scatter chart) that the trendline extends forward. Read/write  **Double**.


## Syntax

 _expression_. **Forward2**

 _expression_ A variable that represents a **[Trendline](trendline-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the number of units that the trendline for the first chart in the active document extends forward and backward. You should run the example on a 2-D column chart that contains a single series that has a trendline.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.SeriesCollection(1).Trendlines(1)

            .Forward2 = 5

            .Backward2 = .5

        End With

    End If

End With
```


## See also


#### Concepts


[Trendline Object](trendline-object-powerpoint.md)

