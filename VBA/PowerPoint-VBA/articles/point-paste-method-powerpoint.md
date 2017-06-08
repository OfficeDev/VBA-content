---
title: Point.Paste Method (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Point.Paste
ms.assetid: 4f6304f2-8cb6-8956-38ff-8718a25aa3ef
ms.date: 06/08/2017
---


# Point.Paste Method (PowerPoint)

Pastes a picture from the Clipboard as the marker on the selected point.


## Syntax

 _expression_. **Paste**

 _expression_ A variable that represents a **[Point](point-object-powerpoint.md)** object.


## Remarks

You can use this method on column, bar, line, or radar charts, and it sets the  **[MarkerStyle](point-markerstyle-property-powerpoint.md)** property to **xlMarkerStylePicture**.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example pastes a picture from the Clipboard into point one in series one for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).Points(1).Paste

    End If

End With


```


## See also


#### Concepts


[Point Object](point-object-powerpoint.md)

