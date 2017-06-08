---
title: TickLabels.Offset Property (PowerPoint)
keywords: vbapp10.chm719012
f1_keywords:
- vbapp10.chm719012
ms.prod: powerpoint
api_name:
- PowerPoint.TickLabels.Offset
ms.assetid: 1bb539a8-a777-e3ff-d1c8-da33b87a2f3f
ms.date: 06/08/2017
---


# TickLabels.Offset Property (PowerPoint)

Returns or sets the distance between the levels of labels, and the distance between the first level and the axis line. Read/write  **Long**.


## Syntax

 _expression_. **Offset**

 _expression_ A variable that represents a **[TickLabels](ticklabels-object-powerpoint.md)** object.


## Remarks

 The default distance is 100 percent, which represents the default spacing between the axis labels and the axis line. The value can be an integer percentage from 0 through 1000, relative to the axis label's font size.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the label spacing of the category axis for the first chart in the active document to twice the current setting, if the offset is less than 500.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlCategory).TickLabels

            If .Offset < 500 then

                .Offset = .Offset * 2

            End If

        End With

    End If

End With
```


## See also


#### Concepts


[TickLabels Object](ticklabels-object-powerpoint.md)

