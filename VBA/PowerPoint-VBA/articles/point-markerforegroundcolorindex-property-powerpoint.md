---
title: Point.MarkerForegroundColorIndex Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Point.MarkerForegroundColorIndex
ms.assetid: 9fb6b350-3eee-305c-dd64-6e3ac009aabc
ms.date: 06/08/2017
---


# Point.MarkerForegroundColorIndex Property (PowerPoint)

Returns or sets the marker foreground color as an index into the current color palette, or as one of the following  **[XlColorIndex](xlcolorindex-enumeration-powerpoint.md)** constants: **xlColorIndexAutomatic** or **xlColorIndexNone**. Read/write **Long**.


## Syntax

 _expression_. **MarkerForegroundColorIndex**

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

            .MarkerBackgroundColorIndex = 4



            ' Set the foreground color to red.

            .MarkerForegroundColorIndex = 3

        End With

    End If

End With


```


## See also


#### Concepts


[Point Object](point-object-powerpoint.md)

