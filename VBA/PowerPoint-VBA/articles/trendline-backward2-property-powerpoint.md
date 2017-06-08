---
title: Trendline.Backward2 Property (PowerPoint)
keywords: vbapp10.chm68186
f1_keywords:
- vbapp10.chm68186
ms.prod: powerpoint
api_name:
- PowerPoint.Trendline.Backward2
ms.assetid: 76415c6a-2c7a-67b5-44a8-23eb768674e5
ms.date: 06/08/2017
---


# Trendline.Backward2 Property (PowerPoint)

Returns or sets the number of periods (or units on a scatter chart) that the trendline extends backward. Read/write  **Double**.


## Syntax

 _expression_. **Backward2**

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

