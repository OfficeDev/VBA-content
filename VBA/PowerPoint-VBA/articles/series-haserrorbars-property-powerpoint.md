---
title: Series.HasErrorBars Property (PowerPoint)
keywords: vbapp10.chm65696
f1_keywords:
- vbapp10.chm65696
ms.prod: powerpoint
api_name:
- PowerPoint.Series.HasErrorBars
ms.assetid: 658e45b6-0c1c-af50-491a-d88468782227
ms.date: 06/08/2017
---


# Series.HasErrorBars Property (PowerPoint)

 **True** if the series has error bars. Read/write **Boolean**.


## Syntax

 _expression_. **HasErrorBars**

 _expression_ A variable that represents a **[Series](series-object-powerpoint.md)** object.


## Remarks

This property is not available for 3-D charts. 


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example removes error bars from series one for the first chart in the active document. You should run the example on a 2-D line chart that has error bars for series one.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).HasErrorBars = False

    End If

End With
```


## See also


#### Concepts


[Series Object](series-object-powerpoint.md)

