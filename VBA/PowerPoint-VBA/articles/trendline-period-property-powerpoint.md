---
title: Trendline.Period Property (PowerPoint)
keywords: vbapp10.chm65720
f1_keywords:
- vbapp10.chm65720
ms.prod: powerpoint
api_name:
- PowerPoint.Trendline.Period
ms.assetid: 7482f5c1-412f-8653-b5f3-1b672125b3e5
ms.date: 06/08/2017
---


# Trendline.Period Property (PowerPoint)

Returns or sets the period for the moving-average trendline. Read/write  **Long**.


## Syntax

 _expression_. **Period**

 _expression_ A variable that represents a **[Trendline](trendline-object-powerpoint.md)** object.


## Remarks

This property can have a value from 2 through 255. 


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the period for the moving-average trendline on the first chart in the active document. You should run the example on a 2-D column chart that has a single series that contains 10 data points and a moving-average trendline.




```vb
With ActiveDocument.InlineShapes(1) 
    If .HasChart Then 
        With .Chart.SeriesCollection(1).Trendlines(1) 
            If .Type = xlMovingAvg Then .Period = 5 
        End With 
    End If 
End With
```


