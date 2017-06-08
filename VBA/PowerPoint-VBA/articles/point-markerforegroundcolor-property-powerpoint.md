---
title: Point.MarkerForegroundColor Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Point.MarkerForegroundColor
ms.assetid: cf0dc6ba-eb97-164d-6c95-d13b75805931
ms.date: 06/08/2017
---


# Point.MarkerForegroundColor Property (PowerPoint)

Sets the marker foreground color as an RGB value or returns the corresponding color index value. Read/write  **Long**.


## Syntax

 _expression_. **MarkerForegroundColor**

 _expression_ A variable that represents a **[Point](point-object-powerpoint.md)** object.


## Remarks

This property applies only to line, scatter, and radar charts. 


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the marker background and foreground colors for the second point in series one for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.SeriesCollection(1).Points(2)

            ' Set the background color to green.

            .MarkerBackgroundColor = RGB(0,255,0)



            ' Set the foreground color to red.

            .MarkerForegroundColor = RGB(255,0,0)

        End With

    End If

End With


```


## See also


#### Concepts


[Point Object](point-object-powerpoint.md)

