---
title: Trendline.Type Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Trendline.Type
ms.assetid: 15eb494c-8e11-491a-5bf1-d7d0ea337e92
ms.date: 06/08/2017
---


# Trendline.Type Property (PowerPoint)

Returns or sets the trendline type. Read/write  **[XlTrendlineType](xltrendlinetype-enumeration-powerpoint.md)**.


## Syntax

 _expression_. **Type**

 _expression_ A variable that represents a **[Trendline](trendline-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example changes the trendline type for the first series of the first chart in the active document. If the series has no trendline, this example fails.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).Trendlines(1).Type = xlMovingAvg

    End If

End With
```


## See also


#### Concepts


[Trendline Object](trendline-object-powerpoint.md)

