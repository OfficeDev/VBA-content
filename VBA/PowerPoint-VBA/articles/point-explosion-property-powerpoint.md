---
title: Point.Explosion Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Point.Explosion
ms.assetid: de7d81aa-bbee-3af5-f38a-74ff7b11c88e
ms.date: 06/08/2017
---


# Point.Explosion Property (PowerPoint)

Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Read/write  **Long**.


## Syntax

 _expression_. **Explosion**

 _expression_ A variable that represents a **[Point](point-object-powerpoint.md)** object.


## Remarks

This property returns 0 (zero) if there is no explosion (the tip of the slice is in the center of the pie). 


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the explosion value for point two of the first chart in the active document. You should run the example on a pie chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).Points(2).Explosion = 20

    End If

End With
```


## See also


#### Concepts


[Point Object](point-object-powerpoint.md)

